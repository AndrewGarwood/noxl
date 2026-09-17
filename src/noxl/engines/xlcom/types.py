"""
*`@file`* `engines/xlcom/types`
"""

from typing import TypeAlias
from win32com.client import CDispatch
from ...models._workbook import WorkbookSource

ComWorkbookSource: TypeAlias = WorkbookSource[CDispatch]
