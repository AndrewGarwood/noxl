"""
*`@file`* `models/cell/cell_data`
- classes using `CellValue` and `CellAddress`
"""
from __future__ import annotations
from typing import (
    Any, Mapping, OrderedDict, Self, TypeVar, Generic
)
from typing_extensions import (TypedDict)
from ..types import (RowKey)
from .address import (CellAddress)
from .value import (CellValue)
from dataclasses import (dataclass, field)
from pyqi import Json, Array

KT = TypeVar('KT', bound=str | RowKey)
VT = TypeVar('VT')

class HeaderIndexDict(TypedDict):
    row: int
    columns: dict[str, int]

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
                addr=CellAddress(row=self.row, col=self.columns[h])
            )
        ))

    @property
    def origin(self) -> CellAddress:
        if (not self.columns):
            raise RuntimeError("Unable to build origin address, HeaderIndex.columns is empty")
        return CellAddress(row=self.row, col=tuple(self.columns.values())[0])

    def get_offset_origin(self, row_offset: int=1, col_offset: int=0) -> CellAddress:
        """ i.e. `get_data_origin` """
        return self.origin.offset(rows=row_offset, cols=col_offset)

class MappedCellDataDict(TypedDict):
    sheetname: str
    header: HeaderIndexDict
    data: dict[Any, Any]

@dataclass
class MappedCellData(Json.Serial[MappedCellDataDict], Generic[KT, VT]):
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