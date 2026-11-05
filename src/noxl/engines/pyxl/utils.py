"""
*`@file`* `engines/pyxl/utils`
"""

from typing import Any, Callable
from pyqi import Array
from lute import annotate, LogMsg
from openpyxl import load_workbook
from .types import (PyxlWorkbook, PyxlWorkbookSource, PyxlWorksheet, PyxlSheetInfo)

#TODO have resolve_pyxl_wb return tuple where second val is bool indicating 
# whether module owns the wb or not, then only close workbook in `finally` 
# statements when wb is owned...

def resolve_pyxl_wb(
    wb_source: PyxlWorkbookSource,
    /, 
    read_only: bool = False, 
    data_only: bool = True
) -> PyxlWorkbook:
    return (
        wb_source if isinstance(wb_source, PyxlWorkbook)
        else load_workbook(wb_source, read_only=read_only, data_only=data_only)
    )

def is_hidden_worksheet(ws: PyxlWorksheet, /) -> bool:
    return "hidden" in ws.sheet_state.lower()

def get_visible_sheetnames(wb_source: PyxlWorkbookSource, /) -> list[str]:
    wb = resolve_pyxl_wb(wb_source)
    return Array.filter(wb.sheetnames, lambda k: not is_hidden_worksheet(wb[k]))

def get_sheetinfo(wb_source: PyxlWorkbookSource, /) -> dict[str, PyxlSheetInfo]:
    wb = resolve_pyxl_wb(wb_source, True, True)
    result: dict[str, PyxlSheetInfo] = dict()
    for i, s in enumerate(wb.sheetnames):
        result[s] = PyxlSheetInfo(wb[s].sheet_state, i)
    if not isinstance(wb_source, PyxlWorkbook):
        wb.close()
    return result


def find_sheetname(
    wb_source: PyxlWorkbookSource,
    filter_fn: Callable[[str], bool],
    /,
) -> str:
    wb: PyxlWorkbook | None = None
    try:
        wb = resolve_pyxl_wb(wb_source)
        valid_sheetnames = wb.sheetnames
        return Array.find(valid_sheetnames, filter_fn, strict=True)
    except Exception as e:
        raise ValueError(LogMsg.err(e,
            f"Unable to find sheetname matching filter criteria",
            wb_source=annotate(wb_source),
            filter_fn=annotate(filter_fn)
        ))
    finally:
        if wb is not None: wb.close()

