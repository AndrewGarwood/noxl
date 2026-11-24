"""
*`@file`* `engines/pyxl/worksheet/read`
"""
from __future__ import annotations as _annotations
from contextlib import closing, contextmanager
from pathlib import Path
from typing import (Any, Callable, Generator, Literal, overload)
from lute import (LogMsg, get_logger, annotate)
from pyqi import (export, guard, require)
from openpyxl.utils import (get_column_letter)
from openpyxl.cell import Cell as PyxlCell
from noxl.models import (CellAddress, CellAddressLike, HeaderIndex, normalize_cell_value)
from ..types import (PyxlWorkbook, PyxlWorksheet, PyxlWorkbookSource, Row)
from ..session import _workbook_context

mlog = get_logger(__file__)


@contextmanager
def _worksheet_context(
    src: PyxlWorksheet | PyxlWorkbookSource,
    sheetname: str | None,
    *,
    read_only: bool = False
) -> Generator[PyxlWorksheet, None, None]:
    if isinstance(src, PyxlWorksheet):
        yield src
        return
    if not sheetname:
        raise ValueError(LogMsg.bad_arg(
            sheetname, "sheetname",
            "sheetname must be defined when src is PyxlWorkbookSource",
            str
        ))
    with _workbook_context(src, read_only=read_only) as wb:
        ws = wb[sheetname]
        if not isinstance(ws, PyxlWorksheet):
            raise TypeError(LogMsg(f"Sheet '{sheetname}' is not a worksheet"))
        yield ws

# kinda unnecessary but useful if don't want to type ws.cell(...).value all the time 
def cell_value(ws: PyxlWorksheet, /, row: int, column: int) -> Any:
    """ 
    *`@returns`* `ws.cell(row, column).value`
    """
    return ws.cell(row, column).value


def get_all_rows(ws: PyxlWorksheet, /, values_only: bool=True) -> list[Row]:
    """ *`@returns`* `list(ws.iter_rows(values_only=values_only)) = list[Sequence[Any]]` """
    return list(ws.iter_rows(values_only=values_only))

@overload
def rows(
    src: str | Path,
    /, 
    sheetname: str,
    origin: CellAddressLike="A1", 
    max_cols: int | None=None, 
    values_only: Literal[True]=True
) -> list[Row]: ...

@overload
def rows(
    src: PyxlWorkbook,
    /,
    sheetname: str,
    origin: CellAddressLike="A1",
    max_cols: int | None=None,
    values_only: bool=True
) -> list[Row]: ...

@overload
def rows(
    src: PyxlWorksheet,
    sheetname: str | None=None,
    origin: CellAddressLike | None=None,
    max_cols: int | None=None,
    values_only: bool=True,
) -> list[Row]: ...

def rows(
    src: PyxlWorksheet | PyxlWorkbookSource,
    sheetname: str | None=None,  
    origin: CellAddressLike | None=None, 
    max_cols: int | None=None, 
    values_only: bool=True
) -> list[Row]:
    """Read eager rows, borrowing objects and managing only path acquisitions."""
    if isinstance(src, (str, Path)) and not values_only:
        raise ValueError(LogMsg.bad_arg(
            values_only, "values_only",
            "Cell results require a workbook or worksheet; open a PyxlSession first",
            bool
        ))
    try:
        with _worksheet_context(src, sheetname) as ws:
            origin = CellAddress.from_any(origin) if origin else CellAddress(1, 1)
            result: list[Row] = []
            row_start = max(origin.row, ws.min_row)
            col_start = max(origin.col, ws.min_column)
            col_end = (
                origin.col + max_cols - 1 if max_cols
                else ws.max_column
            )
            for row in ws.iter_rows(
                min_row=row_start,
                min_col=col_start,
                max_col=col_end,
                values_only=values_only
            ):
                result.append(row)
            return result
    except Exception as e:
        raise RuntimeError(LogMsg.err(
            e,
            src=annotate(src),
            sheetname=sheetname,
            origin=str(origin),
            max_cols=max_cols,
            values_only=values_only
        )) from e

def _default_is_empty_value(value: Any, /):
    return (value is None or value == "")

def _is_empty_row(
    row: Row, 
    /,
    *, 
    is_empty_value: Callable[[Any], bool]=_default_is_empty_value
) -> bool:
    """
    - *`@param`* `row (Row := Sequence[Any])` 
        - can be `Sequence[PythonLiteral]` or `Sequence[openpyxl.cell.Cell]`
    - *`@param`* `is_empty_value (Callable[[PyxlCell | PythonLiteral], bool])`
    - *`@returns`* 
    `all(_is_empty_cell(x, is_empty_value=is_empty_value) for x in row)`
    - @consideration: could instead split code / make assumption from sampling
    on whether `row` is `Sequence[PyxlCell]` or is `Sequence[PythonLiteral]` 
    so that don't check `isinstance(x, PyxlCell)` `for x in Row` 
    """
    return all(_is_empty_cell(x, is_empty_value=is_empty_value) for x in row)

def _is_empty_cell(
    cell: PyxlCell | Any, 
    /,
    *, 
    is_empty_value: Callable[[Any], bool]=_default_is_empty_value
) -> bool:
    """ 
    *`@param`* `cell` - an instance of `openpyxl.cell.Cell` or a python literal 
    """
    return (
        is_empty_value(cell) if not isinstance(cell, PyxlCell) 
        else is_empty_value(cell.value)
    )

@overload
def row_count(
    src: PyxlWorkbookSource,
    /,
    *,
    sheetname: str,
    origin: CellAddressLike = "A1",
    num_cols: int | None = None,
    is_empty_row: Callable[[Row], bool] = _is_empty_row,
    use_max_row: bool = False
) -> int: ...


@overload
def row_count(
    src: PyxlWorksheet,
    /,
    *,
    origin: CellAddressLike = "A1",
    num_cols: int | None = None,
    is_empty_row: Callable[[Row], bool] = _is_empty_row,
    use_max_row: bool = False
) -> int: ...


def row_count(
    src: PyxlWorksheet | PyxlWorkbookSource,
    /,
    *,
    sheetname: str | None = None,
    origin: CellAddressLike = "A1",
    num_cols: int | None = None,
    is_empty_row: Callable[[Row], bool] = _is_empty_row,
    use_max_row: bool = False
) -> int:
    """
    - Count **contiguous** non-empty rows from `origin` to first empty row encountered or `ws.max_row`
        - includes `origin.row`.
        - i.e. `relative_row_count`
    - *`@param`* `num_cols (int | None)` 
        - positive width starting at the origin column. If omitted, uses the remaining worksheet columns. 
    - *`@param`* `is_empty_row` 
        - defaults to `_is_empty_row()` where `None` and `""` are considered empty cell values. 
    
    Headers are not detected or skipped. Paths are opened read-only; objects are borrowed.
    """
    start = CellAddress.from_any(origin)
    if (num_cols is not None):
        require.positive_integer(num_cols, "row_count(..., *, <num_cols>, )")
    with _worksheet_context(src, sheetname, read_only=True) as ws:
        last_row = ws.max_row
        if use_max_row:
            if last_row is None:
                raise ValueError(LogMsg(
                    "use_max_row requires a known worksheet max_row (int)",
                    "(Worksheet.max_row returned None)"
                ))
            return max(0, last_row - start.row + 1)
        if (last_row is not None and start.row > last_row):
            return 0
        last_col =(
            ws.max_column if num_cols is None 
            else start.col + num_cols
        )
        if (last_col is not None and last_col < start.col):
            return 0
        with closing(ws.iter_rows(
            min_row=start.row,
            max_row=last_row,
            min_col=start.col,
            max_col=last_col,
            values_only=True
        )) as selected_rows:
            count = 0
            for row in selected_rows:
                if is_empty_row(row):
                    break
                count += 1
            return count


__all__ = export(
    rows, 
    cell_value, 
    get_all_rows,
    row_count
)