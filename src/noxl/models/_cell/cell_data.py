"""
*`@file`* `models/cell/cell_data`
- classes using `CellValue` and `CellAddress`
"""
from __future__ import annotations
from typing import (
    Any, Mapping, OrderedDict, Self, TypeVar, Generic,
)
from ..types import (RowKey)
from .address import (CellAddress)
from .value import (CellValue)
from dataclasses import (dataclass, field)
from pyqi import Json, Array

KT = TypeVar('KT', bound=str | RowKey)
VT = TypeVar('VT')

@dataclass
class HeaderIndex():
    """
    ```
    row: int
    columns: OrderedDict[str, int] = field(default_factory=OrderedDict)

    - @property def headers(self) -> tuple[str, ...]
    - @property def header_cells(self) -> tuple[CellValue[str, str], ...]
    ```
    """
    row: int
    columns: OrderedDict[str, int] = field(default_factory=OrderedDict)
    @property
    def headers(self):
        return tuple(self.columns.keys())

    @property
    def header_cells(self):
        return tuple(Array.map(
            self.headers, 
            lambda h: CellValue(
                key=h, 
                value=h, 
                addr=CellAddress(self.row, self.columns[h])
            )
        ))
    
    
@dataclass
class MappedCellData(Json.Serial, Generic[KT, VT]):
    """
    ```
    sheetname: str
    header: HeaderIndex
    data: Mapping[KT, CellValue[KT, VT]]
    ```
    """
    sheetname: str
    header: HeaderIndex
    data: Mapping[RowKey, CellValue[KT, VT]]