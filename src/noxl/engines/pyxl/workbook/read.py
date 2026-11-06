"""
*`@file`* `engines/pyxl/workbook/read`
"""
from typing import Any, Callable, OrderedDict, Sequence
from pyqi import (Array, String, Function, require)
from openpyxl.utils import (get_column_letter)
from noxl.models import (CellAddress, CellAddressLike, HeaderIndex, normalize_cell_value)
from ..types import (PyxlWorkbook, PyxlWorksheet, PyxlWorkbookSource, Row)
from ..utils import (resolve_pyxl_wb, get_visible_sheetnames, get_sheetinfo)
from lute import (LogMsg, annotate)

# resolve = resolve_pyxl_wb # not colored correctly... maybe it's a 3.11 thing
def resolve(
    wb_source: PyxlWorkbookSource, 
    /, 
    read_only: bool=False, 
    data_only: bool=True
) -> PyxlWorkbook:
    return resolve_pyxl_wb(wb_source, read_only=read_only, data_only=data_only)

def sheetinfo(wb_source: PyxlWorkbookSource,/):
    return get_sheetinfo(wb_source)

