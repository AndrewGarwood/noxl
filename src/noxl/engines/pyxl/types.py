"""
*`@file`* `engines/pyxl/types`
"""

from typing import TypeAlias
from openpyxl.workbook import Workbook as PyxlWorkbook
from ...models._workbook import WorkbookSource

ComWorkbookSource: TypeAlias = WorkbookSource[PyxlWorkbook]
