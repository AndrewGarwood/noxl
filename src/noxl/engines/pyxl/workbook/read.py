"""
*`@file`* `engines/pyxl/workbook/read`
"""
from ..types import (
    PyxlWorkbook, 
    PyxlWorkbookSource, 
    PyxlWorksheet, 
    PyxlSheetInfo,
    )
from ..utils import (resolve_pyxl_wb, get_sheetinfo)

# resolve = resolve_pyxl_wb # not colored correctly... maybe it's a 3.11 thing
def resolve(
    wb_source: PyxlWorkbookSource, 
    /, 
    read_only: bool=False, 
    data_only: bool=True
) -> PyxlWorkbook:
    return resolve_pyxl_wb(wb_source, read_only=read_only, data_only=data_only)

# ""
def sheetinfo(wb_source: PyxlWorkbookSource,/):
    return get_sheetinfo(wb_source)
