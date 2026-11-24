# noxl

## pyxl workbook ownership

Use a session when several operations need the same workbook:

```python
from pathlib import Path
from noxl.engines import pyxl

path = Path(r"C:\reports\input.xlsx")

with pyxl.PyxlSession() as session:
    wb = session.open_workbook(path, read_only=True, data_only=True)
    info = pyxl.get_sheetinfo(wb)
    rows = pyxl.read_rows(wb, sheetname="Data")
```

The session closes all workbooks it opens on exit, including exceptional exit.
Helpers receiving an existing workbook or worksheet only borrow it and never
close it. Helpers receiving a path manage their own short-lived session.
Closing does not save; call `wb.save(destination)` explicitly when needed.

### Migrating existing calls

- Remove `close_wb` arguments. To reuse a workbook, open it in a session and pass
  that workbook to each helper.
- `read_rows(path, values_only=False)` now raises `ValueError`. Pass a workbook
  or worksheet instead, and use the returned cells within its owner's lifetime.
- Raw `resolve_pyxl_wb` and `workbook.resolve` still return ordinary workbooks.
  When they load a path, the caller is responsible for closing the result.
- `read_mapped_data` still reads its options' path and always manages cleanup.
  It does not reuse an outer session's workbook implicitly.
- Sessions are single-use. Repeated opens of a path create independent
  workbooks, and every successful acquisition is cleaned up separately.

See [ADR 0001](docs/adr/0001-pyxl-workbook-ownership.md) for the ownership rules,
cleanup error policy, and architectural alternatives.

## Counting worksheet rows

```python
count = pyxl.get_row_count(path, sheetname="Data", origin="B2", num_cols=3)
```

This counts contiguous rows from row 2 in columns B through D, stopping at
the first empty row or the end of the worksheet. Omit `num_cols` to scan the
remaining worksheet columns; there is no header detection. By default, a row
is empty if every selected value is `None` or `""`. Override `is_empty_row`
to use a different rule. With `use_max_row=True`, the helper counts through
`ws.max_row` inclusive, including blanks, without scanning cell values. This
option requires known row dimensions.

You can also pass a workbook and sheet name or pass a worksheet directly.
Path inputs are opened read-only and closed by the helper.

`read_rows(..., max_cols=...)` is unchanged. The new name `num_cols` makes clear
that it is a width from `origin`, not an absolute column index.