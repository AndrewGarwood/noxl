"""
*`@file`* `engines/pyxl/worksheet/read`
"""
from __future__ import annotations
from pathlib import Path
from typing import (Any,  overload)
import lute
from openpyxl.utils import (get_column_letter)
from noxl.models import (CellAddress, CellAddressLike, HeaderIndex, normalize_cell_value)
from ..types import (PyxlWorkbook, PyxlWorksheet, PyxlWorkbookSource, Row)
from ..utils import (resolve_pyxl_wb, get_visible_sheetnames)

LogMsg = lute.LogMsg
mlog = lute.get_logger(__file__)

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
    src: PyxlWorkbookSource,
    /, 
    sheetname: str,
    origin: CellAddressLike="A1", 
    max_cols: int | None=None, 
    values_only: bool=True
) -> list[Row]: ...

@overload
def rows(
    src: PyxlWorksheet,
    *, 
    values_only: bool=True,
) -> list[Row]: ...

def rows(
    src: PyxlWorksheet | PyxlWorkbookSource,
    sheetname: str | None=None,  
    origin: CellAddressLike | None=None, 
    max_cols: int | None=None, 
    values_only: bool=True
) -> list[Row]:
    wb: PyxlWorkbook | None=None
    try:
        ws: PyxlWorksheet | None=None
        if (isinstance(src, PyxlWorksheet)):
            ws = src
        else:
            if (not sheetname): 
                raise ValueError(LogMsg.bad_arg(
                    sheetname, "sheetname", 
                    "sheetname must be defined when src is PyxlWorkbookSource", 
                    str
                ))
            wb = resolve_pyxl_wb(src)
            ws = wb[sheetname]
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
            src=lute.annotate(src),
            sheetname=sheetname,
            wb="<unknown>" if wb is None else wb.path,
            origin=str(origin),
            max_cols=max_cols,
            values_only=values_only
        )) from e
    finally:
        if (wb is not None): wb.close()