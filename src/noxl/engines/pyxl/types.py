"""
*`@file`* `engines/pyxl/types`
"""

from pathlib import Path
from typing import (Any, Sequence, TypeAlias, TypeGuard)
from dataclasses import dataclass
from openpyxl.worksheet.table import (Table as PyxlTable)
from openpyxl.worksheet._read_only import (ReadOnlyWorksheet)
from openpyxl.worksheet._write_only import (WriteOnlyWorksheet)
from openpyxl.worksheet.worksheet import (Worksheet)
from openpyxl.workbook import (Workbook as PyxlWorkbook)
from ...models._workbook import WorkbookSource

PyxlWorksheet: TypeAlias = Worksheet | ReadOnlyWorksheet

PyxlWorkbookSource: TypeAlias = WorkbookSource[PyxlWorkbook]


def is_pyxl_workbook_source(value: Any,/) -> TypeGuard[PyxlWorkbookSource]:
    return isinstance(value, (str, Path, PyxlWorkbook))

@dataclass(frozen=True)
class PyxlSheetInfo: # TODO see if any other useful props...
    """ `(frozen=True)` """
    state: str
    index: int
    
    @property
    def is_hidden(self): 
        return "hidden" in self.state.lower()


Row: TypeAlias = Sequence[Any]