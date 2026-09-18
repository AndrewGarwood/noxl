"""
*`@file`* `engines/xlcom/types`
"""

from dataclasses import dataclass
from typing import TypeAlias
from win32com.client import CDispatch
from ...models._workbook import WorkbookSource

ComWorkbookSource: TypeAlias = WorkbookSource[CDispatch]

@dataclass
class TableReference:
    wb_source: ComWorkbookSource
    sheetname: str
    tablename: str
