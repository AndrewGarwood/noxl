"""
*`@file`* `models/cell/value`
"""
from __future__ import annotations
from typing import (
    Any, Mapping, OrderedDict, Sequence, 
    Optional, TypeVar, Generic, Union ,TypeAlias, 
    TypedDict, NotRequired
)
from ..types import (RowKey)
from .address import (CellAddress, CellAddressLike, CellAddressDict)
from dataclasses import (dataclass, field)
from pyqi import (Json)
from lute import (LogMsg)


KT = TypeVar('KT', bound=str | RowKey)
VT = TypeVar('VT')

class CellValueDict(TypedDict):
    key: NotRequired[str | list[str]]
    value: NotRequired[Any]
    addr: NotRequired[CellAddressDict]


@dataclass()
class CellValue(Json.Serial, Generic[KT, VT]):
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
            raise e

