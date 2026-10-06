"""
*`@file`* `models/cell/value`
"""
from __future__ import annotations
from typing import (
    Any, Sequence, 
    Optional, TypeVar, Generic, 
    NotRequired, Iterable
)
from typing_extensions import (TypedDict)
from ..types import (RowKey)
from .address import (CellAddress, CellAddressLike, CellAddressDict)
from dataclasses import (dataclass)
from pyqi import (Json) # , Array
from lute import (LogMsg)


KT = TypeVar('KT', bound=str | RowKey)
VT = TypeVar('VT')

class CellValueDict(TypedDict):
    key: NotRequired[str | list[str]]
    value: NotRequired[Any]
    addr: NotRequired[CellAddressDict]


@dataclass()
class CellValue(Json.Serial[CellValueDict], Generic[KT, VT]):
    """
    ```
    key: KT
    value: VT
    addr: Optional[CellAddress]=None
    ```
    """
    key: Optional[KT]=None
    value: Optional[VT]=None
    addr: Optional[CellAddress]=None
    def __init__(
        self, 
        key: Optional[KT], 
        value: Optional[VT]=None, 
        addr: Optional[CellAddressLike]=None
    ):
        self.key = key or ""
        self.value = value
        self.addr = CellAddress.from_any(addr) if addr else None

    def to_dict(self):
        key = (self.key if isinstance(self.key, str) else list(self.key) if isinstance(self.key, Sequence) else "")
        value = (
            self.value if isinstance(self.value, (int, float, type(None))) 
            else list(self.value) if isinstance(self.value, Sequence) 
            else str(self.value)
        )
        result = {
            "key": key,
            "value": value,
            "addr": None if self.addr is None else self.addr.to_dict()
        }
        return result
    @classmethod
    def from_dict(cls, data: Json.JsonDict[CellValueDict], /) -> CellValue:
        try:
            addr = data.get("addr", None)
            if (addr is not None):
                addr = CellAddress.from_dict(addr)
            key = data.get("key", None)
            if (isinstance(key, list)): key = tuple(key)
            return cls(
                key=key, # type: ignore
                value=data.get("value", None),
                addr=addr
            )
        except Exception as e:
            raise RuntimeError(LogMsg.err(e, received=str(data))) from e
        
    # @staticmethod
    # def is_empty(
    #     value: Any, 
    #     /, 
    #     empty_values: Iterable[Any]=[None, ""]
    # ) -> bool:
    #     return is_empty_cell_value(value, empty_values)
    
    # @staticmethod
    # def normalize():
    #     return NotImplemented

DEFAULT_NULL_VALUES = (None, "", "null", "none")

def is_empty_cell_value(
    value: Any, 
    /, 
    empty_values: Iterable[Any]=DEFAULT_NULL_VALUES
) -> bool:
    return normalize_cell_value(value) in empty_values

def normalize_cell_value(
    value: Any, 
    /,
    case_sensitive: bool=False,
    null_to_empty: bool=True,
    null_values: Iterable[Any]=DEFAULT_NULL_VALUES
) -> str:
    if value is None: return ""
    _value = str(value).strip()
    _lower = _value.casefold()
    if (_lower in null_values and null_to_empty): 
        return ""
    return _lower if not case_sensitive else _value