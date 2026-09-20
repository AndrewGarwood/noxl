"""
*`@file`* `engines/pyxl/types`
"""

from typing import TypeAlias
from dataclasses import dataclass
from openpyxl.worksheet.table import (Table as PyxlTable)
from openpyxl.worksheet.worksheet import (Worksheet as PyxlWorksheet)
from openpyxl.workbook import (Workbook as PyxlWorkbook)
from ...models._workbook import WorkbookSource
PyxlWorkbookSource: TypeAlias = WorkbookSource[PyxlWorkbook]


@dataclass(frozen=True)
class PyxlSheetInfo: # TODO see if any other useful props...
    """ `(frozen=True)` """
    state: str
    index: int
    
    @property
    def is_hidden(self): 
        return "hidden" in self.state.lower()