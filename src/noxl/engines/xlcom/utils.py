"""
*`@file`* `engines/xlcom/utils`
"""

from lute import LogMsg, annotate
from .types import (ComWorkbookSource)
from .session import (ExcelSession)



def refresh_all(wb_source: ComWorkbookSource, /):
    """ `wb.RefreshAll()` """
    try:
        with ExcelSession() as xl:
            wb = xl.resolve_wb_source(wb_source)
            wb.RefreshAll()
            xl.wait_for_async_query_complete()
    except Exception as e:
        raise RuntimeError(LogMsg.err(e, wb_source=annotate(wb_source)))


