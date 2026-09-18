"""
*`@module`* `engines/xlcom/__init__`
"""

from .types import *
from .session import *
from .payload import *
from . import (workbook, worksheet, pivot, legacy)

def refresh_all(wb_source: ComWorkbookSource, /):
    """ `wb.RefreshAll()` """
    try:
        with ExcelSession() as xl:
            wb = xl.resolve_wb_source(wb_source)
            wb.RefreshAll()
            xl.wait_for_async_query_complete()
    except Exception as e:
        raise RuntimeError(LogMsg.err(e))




