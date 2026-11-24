# ADR 0001: Explicit ownership and scoped cleanup of pyxl workbooks

- Status: Accepted
- Date: 2026-11-24
- Scope: `noxl.engines.pyxl`
- Implementation status: Implemented in [`PyxlSession`](../../src/noxl/engines/pyxl/session.py)
  and the pyxl convenience helpers, with [lifetime tests](../../tests/test_pyxl_ownership.py).

## Context

Before this decision, the module mixed source resolution with resource ownership:

| Previous code | Lifetime behavior |
| --- | --- |
| [`resolve_pyxl_wb`](../../src/noxl/engines/pyxl/utils.py) and [`workbook.resolve`](../../src/noxl/engines/pyxl/workbook/read.py) | Return an existing workbook unchanged or load a path, without conveying who must close the result. |
| [`get_visible_sheetnames` and `get_sheetinfo`](../../src/noxl/engines/pyxl/utils.py) | Close by default, including caller-supplied workbooks. Cleanup occurs only on the success path. |
| [`find_sheetname` and `find_header`](../../src/noxl/engines/pyxl/utils.py) | Use `finally`, but close resolved workbook inputs by default regardless of their origin. The worksheet branch of `find_header` borrows its input without closing its parent. |
| [`worksheet.rows`](../../src/noxl/engines/pyxl/worksheet/read.py) | Borrows worksheet inputs, but closes workbook inputs by default. `values_only=False` returns cells, not detached values. |
| [`read_mapped_data`](../../src/noxl/engines/pyxl/read_data.py) | Opens `options.path` in read-only mode. `close_wb=False` leaves the workbook unclosed without returning it to the caller. |
| Other worksheet readers and [`worksheet.write`](../../src/noxl/engines/pyxl/worksheet/write.py) | Accept worksheets and do not manage their parent workbooks. This is already the desired borrowing model. |

Consequently, a helper could close a workbook that a caller or another helper
still needed. Conversely, callers could disable cleanup without receiving any
resource they could subsequently close. A `close_wb` boolean describes an action,
not the authority to perform that action.

The existing [pyxl consideration](../../src/noxl/engines/pyxl/.crf.txt)
suggests returning an ownership flag from the resolver. The
[`ExcelSession`](../../src/noxl/engines/xlcom/session.py) implementation offers
another useful precedent: an explicit context owns the workbooks it opens.
The pyxl design should reuse that principle, not COM-specific process
management, path-keyed bookkeeping, or save behavior.

### Verified openpyxl behavior

Pre-refactor checks against the installed openpyxl 3.1.5 established that:

- Passing a read-only workbook to `get_sheetinfo(wb)` closes its archive.
  A subsequent read raises
  `ValueError: Attempt to use ZIP archive that was already closed`.
- Injecting a processing failure after resolution in either metadata helper
  results in zero calls to `wb.close()`.
- Successful normal-mode loading closes the input archive before returning.
  `Workbook.close()` closes a retained `_archive` when present; it does not
  invalidate a normal in-memory workbook or save modifications.
- Two loads of the same path can produce distinct workbooks with different
  `data_only` settings. A path is therefore not a sufficient ownership key.
- The raw resolver returns a supplied workbook unchanged; its load options
  do not reconfigure that workbook.
- `rows(ws, values_only=False)` returns cells whose parent is the supplied
  worksheet, even though the outer result is an eager list.

This distinction explains why premature closure can appear harmless in normal
mode while breaking read-only workflows.

## Decision

**The scope that opens a workbook owns its cleanup. Receiving a workbook or
worksheet as an argument only borrows it.**

Use an explicit `PyxlSession` for caller-controlled lifetimes, plus one internal
context manager for path-based convenience functions. Both use the same session
cleanup mechanism.

### Ownership rules

| Acquisition | Owner | Cleanup boundary |
| --- | --- | --- |
| `session.open_workbook(path, ...)` | That session | Session exit, whether successful or exceptional |
| A convenience helper receives a path | A short-lived internal session | Before the helper returns or propagates an error |
| A helper receives an existing workbook | The existing owner, unchanged | Never closed by the helper |
| A helper receives a worksheet | The owner of its parent workbook | Neither worksheet nor parent closed by the helper |
| Legacy raw resolver loads a path | Its caller | Caller explicitly closes the returned workbook |

Additional invariants:

1. Helpers never acquire ownership merely by resolving an existing object.
   This also applies to workbooks opened by an outer session.
2. Passing a session-owned workbook into another helper does not register it
   with a second owner or extend its lifetime.
3. Ownership is over workbook objects, not filenames. Repeated opens are
   independent acquisitions; there is no implicit cache or path deduplication.
4. Closing never saves. Callers explicitly invoke `wb.save(destination)` while
   the owning scope is active. The session is not a transaction or rollback
   mechanism.
5. File-backed worksheets, cells, and iterators must be used within their
   owner's lifetime. Detached values and metadata may outlive that owner.
   Resource cleanup does not promise to invalidate every Python object.
6. Preserve current loading defaults and result contents unless a behavioral
   change is explicitly identified below. Ownership does not imply
   `read_only=True`, and `data_only=True` is not a saving policy.

### Minimal public session

[`session.py`](../../src/noxl/engines/pyxl/session.py) defines `PyxlSession`,
exported from the pyxl package. Its public surface is:

- `__enter__() -> PyxlSession`
- `open_workbook(path, /, read_only=False, data_only=True) -> PyxlWorkbook`
- `__exit__(exc_type, exc, tb) -> bool`

`path` accepts the existing string/`Path` source forms, not existing workbooks.
There is deliberately no implicit adoption, ownership transfer, global current
session, workbook proxy, or session argument on every helper.

A session is single-use: new, active, then closed. Opening outside the active
context, re-entering an active session, or reusing a closed session raises
`RuntimeError`. Multiple distinct sessions may be nested. Workbooks returned
by `open_workbook` remain ordinary openpyxl objects, not wrappers.

Record each successfully opened workbook immediately and retain a strong
reference until cleanup. Use a standard-library `ExitStack` to schedule cleanup
in reverse acquisition order. Do not copy the xlcom dictionary keyed by path:
that would overwrite ownership records for independent opens of the same file.
No early-close/adopt/release APIs are needed initially; use a smaller nested
session when an earlier cleanup boundary is required.

On exit:

- Attempt cleanup of every registered workbook, even when another close fails.
- Drain cleanup registrations so repeated cleanup cannot close them again.
- With no body exception, raise an `ExceptionGroup` for ordinary cleanup
  failures after all close attempts.
- With a body exception, preserve that exception as primary, log each cleanup
  failure using the existing `lute`/`LogMsg.err` convention, and attach cleanup
  details with `add_note`. Do not replace the operation's failure or silently
  discard cleanup failures.
- Do not suppress the body exception, save workbooks, run garbage collection,
  or add process-management behavior.

The error policy must be implemented explicitly: bare `ExitStack` callback
propagation alone does not provide the specified aggregation and preservation
of the body exception.

A failed `load_workbook` call must propagate and must not register a nonexistent
workbook. Earlier successful acquisitions remain registered for session exit.
This contract covers workbooks successfully returned by openpyxl; it does not
claim that `Workbook.close()` can repair file handles retained inside a loader
that raises before returning a workbook. Explicit input-stream management, if
needed for that separate failure mode, must preserve filename validation and
the normal-mode early release of the source file. Do not introduce file-like
source support or patch openpyxl private archives in this initial refactor.

### One internal acquisition boundary

`_workbook_context` lives in `session.py`, independently of `utils.py`, to avoid
import cycles. Its ownership flow, with annotations omitted:

```python
@contextmanager
def _workbook_context(source, /, read_only=False, data_only=True):
    if isinstance(source, PyxlWorkbook):
        yield source
    else:
        with PyxlSession() as session:
            yield session.open_workbook(
                source, read_only=read_only, data_only=data_only
            )
```

The implementation annotates the source, loading flags, and generator return type.

All path-or-workbook convenience helpers use this boundary instead of manually
resolving and closing. Their worksheet branches simply operate on the supplied
worksheet. No helper needs an `(owned_workbook, owns_workbook)` tuple or its own
ownership-aware `finally` block.

Loading flags apply only when opening a path. Borrowed workbooks keep their
existing modes; do not silently reopen them to satisfy helper defaults.
Callers needing a specific mode open the workbook explicitly in that mode.

### Usage

```python
from pathlib import Path
from noxl.engines import pyxl

path = Path(r"C:\reports\input.xlsx")

# One-shot path input: the helper's internal session owns cleanup.
info = pyxl.get_sheetinfo(path)

# Explicit lifetime: helpers borrow, and the outer session alone closes.
with pyxl.PyxlSession() as session:
    wb = session.open_workbook(path, read_only=True, data_only=True)
    info = pyxl.get_sheetinfo(wb)
    header = pyxl.find_header(wb.active, required_headers=["key", "value"])
    values = pyxl.read_rows(wb.active, values_only=True)
    # wb remains readable here.

# info, header, and values are usable after cleanup.
```

For multiple workbooks, call `open_workbook` again in the same session.
For a workbook created or loaded outside noxl, retain the original owner's
`try/finally` or other cleanup scope; noxl helpers only borrow it.

### Live return values

The convenience path forms are for detached results: sheet names, sheet
metadata, `HeaderIndex`, mapped data, value-only rows, and row counts.

`row_count` and `rows` share worksheet resolution. The count opens paths in
read-only mode and scans values within the owning scope; existing workbooks
and worksheets are borrowed. An explicit `num_cols` is a width from `origin`.
Without it, the worksheet's column extent applies rather than a detected
header, so headerless data does not require special handling.

`rows(..., values_only=False)` is an important exception to the assumption that
an eager list is detached. Keep that operation available for workbook and
worksheet inputs, whose lifetime is explicit. The API rejects a path
combined with `values_only=False` before opening anything, with a `ValueError`
directing callers to open a session and pass its workbook or worksheet.
This is an intentional compatibility change, not an incidental cleanup edit.
Keep the overloads consistent with the accepted forms.

Future APIs returning worksheets or lazy iterators must similarly require a
borrowed workbook/worksheet or expose their own explicit context boundary.
Do not return them from a temporary owning scope. Finish or close iteration
within the owner scope; closing a workbook is not an iterator-lifetime manager.

## Implemented changes and migration

This is an intentional API change, not a silent patch:

1. Added the session and internal acquisition helper, with focused lifetime tests.
   The session and migrated helpers are published together, so new session users
   cannot accidentally call legacy helpers that close their workbooks.
2. Migrated `get_visible_sheetnames`, `get_sheetinfo`, `find_sheetname`,
   `find_header`, and `worksheet.rows` to the shared boundary. Selection,
   normalization, row contents, and load-mode defaults remain unchanged.
   `workbook.sheetinfo` continues delegating to the migrated helper.
3. Migrated `read_mapped_data` using
   `_workbook_context(options.path, read_only=True, data_only=data_only)`.
   It owns this acquisition and returns detached data. Its call to
   `find_header(ws)` borrows without the former `close_wb=False`.
4. Removed `close_wb` from helper signatures and overloads, and applied
   the live-cell path restriction above. Check downstream consumers before
   releasing these breaking changes. Callers needing reuse must open once and
   pass that object, rather than disable cleanup of an invisible acquisition.
5. Retained the raw `resolve_pyxl_wb` and `workbook.resolve` return shapes for
   compatibility, with their low-level contract documented: a path creates a
   caller-owned workbook; an existing workbook is borrowed. They are no longer
   used internally where cleanup is required. Neither return type changed to a
   context manager or tuple under the same name.
6. Documented migration examples and added focused lifetime tests against
   actual read-only files, not only normal-mode workbooks or mocks.

`MappedDataSearchOptions.path` is currently a serialized `Path`, not a workbook
slot. Keep that shared model and `read_mapped_data`'s path-only acquisition in
this refactor. The initial session does not allow reusing an existing workbook
through that API. A future explicit workbook-source parameter can enable that
if needed; do not place live engine objects into serialized options or imply
that an ambient session intercepts its opens.

The accepted implementation removes the flags without a compatibility shim.
Existing calls supplying `close_wb` now raise `TypeError`; update those callers
instead of silently ignoring their requested behavior. No remaining callers
were found in this repository or the adjacent lute and pyqi source trees.
External consumers must make the same migration before adopting this change.

For example:

```python
# Before: helper-specific cleanup flags.
# info = pyxl.get_sheetinfo(wb, close_wb=False)
# cells = pyxl.read_rows(path, sheetname="Data", values_only=False)

# After: the session owns cleanup, and helpers only borrow.
with pyxl.PyxlSession() as session:
    wb = session.open_workbook(path)
    info = pyxl.get_sheetinfo(wb)
    cells = pyxl.read_rows(wb, sheetname="Data", values_only=False)
    first_cell_value = cells[0][0].value
```

## Alternatives considered

| Alternative | Assessment |
| --- | --- |
| Keep `close_wb` flags | Reject: callers must understand nested implementation details, and flags can close borrowed resources or leak inaccessible owned ones. |
| Return `(wb, owned)` from the resolver | Better than the current ambiguity, but repeats exception-safe cleanup at every call site. A context manager enforces the same rule once. |
| Only add a single-workbook context manager | A reasonable smaller API for one-shot operations, but does not directly cover the requested multi-workbook session. Use it internally through the session instead of duplicating cleanup. |
| Require an explicit session for every convenience call | Consistent but unnecessarily removes useful one-shot path APIs. Short-lived internal sessions preserve them. |
| Cache session workbooks by path | Reject initially: repeated opens can have different modes, and cache eviction adds a second lifetime policy. |
| Automatically adopt supplied workbooks | Reject: passing an argument must not transfer cleanup authority. An explicit transfer API can be considered separately if a real use case appears. |
| Share a generic session base with xlcom | Reject initially: only the ownership principle is shared; COM process, save, and shutdown semantics are materially different. |

## Validation

The [lifetime tests](../../tests/test_pyxl_ownership.py) exercise these acceptance
cases using actual read-only workbooks and injected loading/cleanup failures.
Run them without additional test dependencies:

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -p test_pyxl_ownership.py -v
```

- A session closes every successfully opened workbook once, on both ordinary
  exit and a body exception, in reverse acquisition order.
- Two acquisitions of the same path, including different load options, are
  both cleaned up; neither registration overwrites the other.
- A failing later load propagates, registers nothing, and does not prevent
  cleanup of earlier successful loads when the session unwinds.
- Every migrated path helper closes its owned workbook after success and after
  a processing failure. Use a close spy plus a real read-only archive check.
- Every workbook-accepting helper leaves an external read-only workbook
  readable after success and failure. Worksheet helpers leave its parent open.
- Several helpers can use a session-owned read-only workbook consecutively;
  no helper closes it before outer session exit.
- One close failure does not stop remaining closes. Assert cleanup-only error
  aggregation, and preservation plus logging/notes when the body also fails.
- Session use before entry, after exit, and re-entry fails explicitly.
- No cleanup calls `save`; opening and closing does not change file contents.
- Detached results remain usable after exit. Cell-returning calls accept
  borrowed objects, while path-plus-cells calls fail before loading.
- Loading defaults, sheet selection, header normalization, and returned rows
  match pre-refactor behavior for unaffected calls. Existing workbook load
  modes remain unchanged when borrowed.

## Consequences

Ownership becomes structural rather than a convention repeated in booleans.
Single-file callers retain convenience; multi-step and multi-file callers
control the lifetime explicitly. Nested helpers become safe borrowers.

The tradeoffs are a small new session API, intentional migration of close flags
and path-based live-cell calls, and no caching or ownership-transfer feature.
The implementation does not address unrelated row-scanning, data-validation, writing,
or COM-session issues.
