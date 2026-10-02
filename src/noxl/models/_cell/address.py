"""
*`@file`* `models/cell/address`
"""
from __future__ import annotations
from typing import (
    Any, Mapping, OrderedDict, Sequence, 
    Optional, TypeVar, Generic, Union ,TypeAlias, 
    TypeGuard
)
from typing_extensions import (TypedDict)
from dataclasses import (dataclass, field)
import re
from openpyxl.utils import (get_column_letter, column_index_from_string, coordinate_to_tuple)
from lute import LogMsg
from pyqi import (Json, require)



EXCEL_RC_PAT = re.compile(r"^R(?P<row>\d+)C(?P<col>\d+)$", re.IGNORECASE)
_LETTER_PAT = re.compile(r"^(?P<col>[A-Z]+)$", re.IGNORECASE)

class CellAddressDict(TypedDict):
    row: str | int
    col: str | int

CellAddressLike: TypeAlias = Union[
    "CellAddress", 
    CellAddressDict, 
    str, 
    Sequence[str | int]
]
""" 
- `CellAddress`
- `CellAddressDict = TypedDict({ row: str | int; col: str | int })`
- `Excel RowCol Pattern (str)` `R{row_index}C{col_index}`
- `Excel Address Pattern (str)` `{column_letter}{row_index}` → `'{col}{row}'`
- `tuple[int, int]` → assume `(row, col)`
- `tuple[str, int]` → assume `(col, row)`
- `tuple[str, str] and (t[0] matches [A-Z]+) and t[1] matches \\d+` → assume `(col, row)`
- `tuple[str, str] and t[0] matches \\d+ and t[1] matches \\d+` → assume `(row, col)`
"""
@dataclass()
class CellAddress(Json.Serial[CellAddressDict]):
    """ 1-based indexing for (row, col) """
    row: int
    col: int
    @classmethod
    def from_str(cls, addr: str, /) -> CellAddress:
        addr = addr.strip()
        if (len(addr) < 2):
            raise ValueError(LogMsg(f"Invalid cell address string: '{addr}'"))
        rc_match = EXCEL_RC_PAT.match(addr)
        if (rc_match):
            r, c = (int(x) for x in rc_match.groups())
            return cls(row=r, col=c)
        if (_LETTER_PAT.match(addr[0]) and addr[1:].isdigit()):
            r, c = coordinate_to_tuple(addr)
            return cls(row=r, col=c)
        raise ValueError(LogMsg(
            f"Invalid cell address string",
            expected="'R{row_index}C{col_index}' or '{column_letter}{row_index}'",
            received=addr
        ))
    
    @classmethod
    def from_dict(cls, addr: Json.JsonDict[CellAddressDict], /) -> CellAddress:
        row = addr["row"]
        col = addr["col"]
        if (is_valid_row_index(row) and is_valid_col_index(col)):
            return cls(
                row=int(row),
                col=(
                    col if isinstance(col, int)
                    else int(col) if col.isdigit() 
                    else column_index_from_string(col.upper())
                ),
            )
        raise ValueError(LogMsg(f"Invalid CellAddressDict", received=str(addr)))

    @classmethod
    def from_seq(cls, addr: Sequence[str | int], /) -> CellAddress:
        first, second = addr[:2]
        if isinstance(first, int):
            if isinstance(second, int): # (row, col)
                return cls(row=first, col=second)
            elif isinstance(second, str):
                if second.isdigit():
                    return cls(row=first, col=int(second))
                elif _LETTER_PAT.match(second):
                    return cls(row=first, col=column_index_from_string(second))
        if isinstance(first, str) and isinstance(second, int): # (col, row)
            return cls(
                row=second,
                col=column_index_from_string(first.upper()),
            )
        
        if ( # (col, row) where both are strings, e.g. ("A", "1")
            isinstance(first, str)
            and isinstance(second, str)
            and _LETTER_PAT.fullmatch(first)
            and second.isdigit()
        ):
            return cls(
                row=int(second),
                col=column_index_from_string(first.upper()),
            )
        if ( # (row, col) where both are strings, e.g. ("1", "2")
            isinstance(first, str)
            and isinstance(second, str)
            and first.isdigit()
            and second.isdigit()
        ):
            return cls(row=int(first), col=int(second))
        raise ValueError(LogMsg(
            f"Invalid cell address Sequence value", 
            expected=f"addr: Sequence[str | int]<2>",
            received=str(addr),
        ))

    @classmethod
    def from_any(cls, addr: CellAddressLike, /) -> CellAddress:
        """
        *`@param`* `addr (CellAddressLike)`
        - `CellAddress`
        - `CellAddressDict = { row: str | int; col: str | int}`
        - `Excel RowCol Pattern (str)` `R{row_index}C{col_index}`
        - `Excel Address Pattern (str)` `{column_letter}{row_index}` → `'{col}{row}'`
        - `tuple[int, str]` → assume `(row, col)`
        - `tuple[int, int]` → assume `(row, col)`
        - `tuple[str, int]` → assume `(col, row)`
        - `tuple[str, str] and (t[0] matches [A-Z]+) and t[1].isdigit()` → assume `(col, row)`
        - `tuple[str, str] and t[0].isdigit() and t[1].isdigit()` → assume `(row, col)`
        """
        if isinstance(addr, CellAddress):
            return addr
        if isinstance(addr, str): # Format: "A1" or "R1C1"
            return cls.from_str(addr)
        if isinstance(addr, dict):
            return cls.from_dict(addr)
        if isinstance(addr, (tuple, list)) and len(addr) == 2:
            return cls.from_seq(addr)
        
        raise ValueError(LogMsg(
            f"Unsupported CellAddressLike format.",
            expected=f"CellAddressLike = Union[CellAddress, CellAddressDict, str, Sequence[str | int]<2>]",
            received=f"{type(addr)=}, {addr=}",
        ))

    
    @property
    def excel_address(self) -> str:
        """ `{column_letter}{row_index}` """
        return f"{get_column_letter(self.col)}{self.row}"

    def to_dict(self):
        return CellAddressDict(row=self.row, col=self.col)

    def offset(self, rows: int = 0, cols: int = 0) -> CellAddress:
        """ 
        *`@returns`* `CellAddress(row=self.row + rows, col=self.col + cols)`
        
        *`@throws`* `Exception` if offset results in an `index < 1`
        """
        return CellAddress(
            row=require.positive_integer(self.row+rows, "new_row = self.row + rows"), 
            col=require.positive_integer(self.col+cols, "new_col = self.col + cols")
        )



RowIndex: TypeAlias = str | int
ColumnIndex: TypeAlias = str | int
# PrimitiveIndex: TypeAlias = str | int
# def is_primitive_index(value: Any,/) -> TypeGuard[PrimitiveIndex]:
#     return isinstance(value, (str, int))


#TODO: should probably check index >= 1

def is_valid_row_index(value: Any, /) -> TypeGuard[RowIndex]:
    return (
        isinstance(value, int) 
        or (isinstance(value, str) and value.isdigit())
    )

def is_valid_col_index(value: Any, /) -> TypeGuard[ColumnIndex]:
    return (isinstance(value, int)
        or (isinstance(value, str) and value.isdigit())
        or (isinstance(value, str) and bool(_LETTER_PAT.fullmatch(value)))
    )