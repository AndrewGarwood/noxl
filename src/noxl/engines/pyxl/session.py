"""
*`@file`* `engines/pyxl/session`
"""
from __future__ import annotations as _annotations
from contextlib import ExitStack, contextmanager
from pathlib import Path
from types import TracebackType
from typing import Generator, Literal
from openpyxl import load_workbook
from pyqi import (fs, export)
from lute import (get_logger, LogMsg)
from .types import (PyxlWorkbook, PyxlWorkbookSource)


mlog = get_logger(__file__, name_depth=1)


class PyxlSession():
    """Own workbooks opened within a single context; never save on exit."""

    def __init__(self) -> None:
        self._state: Literal["new", "active", "closed"] = "new"
        self._stack = ExitStack()
        self._close_errors: list[tuple[fs.StrPath, Exception]] = []

    def __enter__(self) -> PyxlSession:
        if self._state != "new":
            raise RuntimeError("PyxlSession is single-use and cannot be re-entered")
        self._state = "active"
        return self

    def open_workbook(
        self,
        path: fs.StrPath,
        /,
        read_only: bool = False,
        data_only: bool = True
    ) -> PyxlWorkbook:
        """Load a path and close the returned workbook when this session exits."""
        if self._state != "active":
            raise RuntimeError("PyxlSession must be active in a with statement")
        if not isinstance(path, (str, Path)):
            raise TypeError(LogMsg.bad_arg(
                path, "path", "PyxlSession opens paths, not existing workbooks", str, Path
            ))
        wb = load_workbook(path, read_only=read_only, data_only=data_only)
        self._stack.callback(self._close_workbook, wb, path)
        return wb

    def _close_workbook(self, wb: PyxlWorkbook, path: fs.StrPath) -> None:
        try:
            wb.close()
        except Exception as e:
            self._close_errors.append((path, e))

    def __exit__(
        self,
        _exc_type: type[BaseException] | None,
        exc: BaseException | None,
        _tb: TracebackType | None
    ) -> Literal[False]:
        self._state = "closed"
        self._stack.close()
        errors, self._close_errors = self._close_errors, []
        if exc is not None:
            for path, error in errors:
                exc.add_note(
                    f"PyxlSession failed closing {path!s}: {type(error).__name__}: {error}"
                )
                mlog.warning(LogMsg.err(
                    error, "Failed closing workbook during PyxlSession cleanup.",
                    workbook=str(path)
                ))
        elif errors:
            for path, error in errors:
                error.add_note(f"PyxlSession failed closing workbook: {path!s}")
            raise ExceptionGroup(
                "Failed closing workbooks during PyxlSession cleanup",
                [error for _, error in errors]
            )
        return False


@contextmanager
def _workbook_context(
    src: PyxlWorkbookSource,
    /,
    read_only: bool = False,
    data_only: bool = True
) -> Generator[PyxlWorkbook, None, None]:
    if isinstance(src, PyxlWorkbook):
        yield src
    else:
        with PyxlSession() as session:
            yield session.open_workbook(
                src, 
                read_only=read_only, 
                data_only=data_only
            )

__all__ = export(
    PyxlSession
)