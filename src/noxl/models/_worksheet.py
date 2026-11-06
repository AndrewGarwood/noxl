"""
*`@file`* `models/_worksheet`
"""
from __future__ import annotations
from pathlib import Path
from dataclasses import dataclass, field
from typing import (
    Any, Callable, Mapping,
    TypeVar, Generic,TypeAlias, 
    NotRequired
)
from typing_extensions import (TypedDict)
from pyqi import (Json, String, fs, Array)
from lute import LogMsg, callsite
from ._cell import *

T = TypeVar('T')
KT = TypeVar('KT', bound=str | RowKey)
VT = TypeVar('VT')


class SheetRangeDict(TypedDict):
    sheetname: str
    origin: NotRequired[str | list[int] | list[str | int]]
    num_cols: NotRequired[int]

@dataclass
class SheetRange(Json.Serial[SheetRangeDict]):
    """
    - *`@property`* `sheetname` `(str)`
    - *`@property`* `origin` `(CellAddressLike) default = 'A2'`
    - *`@property`* `num_cols` `(int | None) default = None`
    """
    sheetname: str = "Sheet1"
    origin: CellAddressLike = "A2"
    num_cols: int | None = None
    @classmethod
    def from_dict(cls, data: Json.JsonDict[SheetRangeDict], /) -> SheetRange:
        try:
            _origin = data.get("origin", "A2")
            if isinstance(_origin, list):
                _origin = (_origin[0], _origin[1])
            return cls(
                sheetname=data["sheetname"],
                origin=_origin,
                num_cols=data.get("num_cols", None)
            )
        except (KeyError, ValueError): raise
        except Exception as e:
            raise RuntimeError(LogMsg.err(error=e, data=data)) from e

    @classmethod
    def range_dict(cls, data: dict[str, SheetRangeDict]) -> SheetRangeDictionary:
        result: SheetRangeDictionary = {}
        for key, val in data.items():
            result[key] = cls.from_dict(val)
        return result

SheetRangeDictionary: TypeAlias = dict[str, SheetRange]
""" e.g. `{ [txnId: string]: SheetRangeSpec }` """


class SheetCopyEndpointDict(TypedDict):
    path: str
    sheetname: str
    origin: NotRequired[str | list[str | int]]
    num_cols: NotRequired[int]

@dataclass
class SheetCopyEndpoint(Json.Serial[SheetCopyEndpointDict]):
    """
    - *`@property`* `path` `(Path)`
    - *`@property`* `sheetname` `(str)`
    - *`@property`* `origin` `(CellAddressLike) default = 'A2'`
    - *`@property`* `num_cols` `(int | None) default = None`
    """
    sheetname: str
    origin: CellAddressLike
    path: Path = field(default_factory=Path)
    num_cols: int | None = None

    @classmethod
    def from_range_spec(
        cls, 
        path: fs.StrPath, 
        spec: SheetRange, 
        require_absolute: bool=True
    ) -> SheetCopyEndpoint:
        path = fs.to_path(path)
        if (require_absolute and not path.is_absolute()):
            raise ValueError(LogMsg(
                f"Invalid arg 'path'",
                expected="absolute path",
                received=f"{str(path)=}"
            ))
        return cls(
            path=path,
            sheetname=spec.sheetname,
            origin=spec.origin,
            num_cols=spec.num_cols
        )
    @classmethod
    def from_dict(cls, data: SheetCopyEndpointDict | Mapping[str, Any], /):
        try: # could use SheetRangeSpec.from_dict, then cls.from_range_spec...
            return cls(
                path=fs.to_path(data['path']),
                sheetname=data['sheetname'],
                origin=data.get('origin', "A2"),
                num_cols=data.get('num_cols', None)
            )
        except (KeyError, ValueError): raise
        except Exception as e:
            raise RuntimeError(LogMsg.err(error=e, data=Json.serialize(data))) from e


class SheetCopyRequestDict(TypedDict):
    src: SheetCopyEndpointDict
    dst: SheetCopyEndpointDict

@dataclass
class SheetCopyRequest(Json.Serial[SheetCopyRequestDict]):
    src: SheetCopyEndpoint
    dst: SheetCopyEndpoint

    @classmethod
    def from_dict(cls, data: SheetCopyRequestDict | dict[str, Any]) -> Self:
        return cls(
            src=SheetCopyEndpoint.from_dict(data['src']),
            dst=SheetCopyEndpoint.from_dict(data['dst'])
        )


@dataclass
class SheetCopyResult:
    """
    ```
    src_path: Path
    src_sheetname: str
    dst_path: Path
    dst_sheetname: str
    curr_rows: list[list[Any]]
    prev_row_count: int
    - property def curr_row_count(self) -> int: = len(self.curr_rows)
    - property def num_cols(self) -> int: = len(self.curr_rows[0]) if self.curr_rows else 0
    - property def num_rows_cleared(self) -> int: = max(self.prev_num_rows, self.curr_num_rows)
    ```
    """
    src_path: Path
    src_sheetname: str
    dst_path: Path
    dst_sheetname: str
    curr_rows: list[list[Any]]
    prev_row_count: int

    @property
    def curr_row_count(self) -> int:
        """ number of rows after write src rows to dst sheet """
        return len(self.curr_rows)

    @property
    def num_cols(self) -> int:
        curr_ncols = len(self.curr_rows[0]) if self.curr_rows else 0
        return curr_ncols

    @property
    def num_rows_cleared(self) -> int:
        return max(self.prev_row_count, self.curr_row_count)


@dataclass
class SheetOptions:
    """
    ```
    path: Path
    sheetname: str
    ```
    """
    path: Path # could make this generic for PyxlWorkbookSource or ComWorkbookSource
    sheetname: str

class SheetLookupOptionsDict(TypedDict):
    """ ***deprecated*** """
    path: str
    sheetname: str
    key_map: Mapping[str, Any]
    """ `{ [col: string]: value }` """
    val_cols: list[str]

@dataclass
class SheetLookupOptions(SheetOptions, Json.Serial[SheetLookupOptionsDict]):
    """
    ***deprecated***
    
    *`@extends`* `(SheetOptions, Json.Serial[SheetLookupOptionsDict])`
    - lookup values for single key constructed from key_map.keys()
    ```
    path: Path
    sheetname: str
    key_map: Mapping[str, Any]
    val_cols: list[str]
    ```
    """
    key_map: Mapping[str, Any]
    val_cols: list[str]
    @classmethod
    def from_dict(cls, data: Json.JsonDict[SheetLookupOptionsDict], /):
        try:
            key_map = data["key_map"]
            if not isinstance(key_map, Mapping):
                raise ValueError(LogMsg(
                    f"Invalid value for key 'key_map'",
                    f"expected: Mapping[str, Any]",
                    f"received: {type(key_map)=} = {key_map}"
                ))
            val_cols=data["val_cols"]
            if not isinstance(val_cols, (list, tuple)):
                raise ValueError(LogMsg(
                    f"Invalid value for key 'val_cols'",
                    f"expected: Sequence[str]",
                    f"received: {type(val_cols)=} = {val_cols}"
                ))
            return cls(
                path=fs.to_path(data["path"]),
                sheetname=data["sheetname"],
                key_map=key_map,
                val_cols=list(set(val_cols))
            )
        except (KeyError, ValueError): raise
        except Exception as e:
            raise RuntimeError(LogMsg.err(
                e,
                received=str(data)
            )) from e

    def key_str(self, sep: str="|"):
        """ *`@returns`* `sep.join(self.key_map.keys())` """
        return sep.join(self.key_map.keys())

    def key_tup(self):
        """ *`@returns`* `stuple(self.key_map.keys())` """
        return tuple(self.key_map.keys())


@dataclass
class SheetLookupResult(Generic[KT, VT]):
    header: HeaderIndex
    key: KT
    values: Mapping[str, CellValue[KT, VT]]


class MappedDataSearchOptionsDict(TypedDict):
    path: str
    sheetname: NotRequired[str]
    key_cols: list[str]
    val_cols: NotRequired[list[str]]
    case_sensitive: NotRequired[bool]
    col_offset: NotRequired[int]
    sheetname_selector: NotRequired[String.StringConditionOptionsDict]
    column_selector: NotRequired[String.StringConditionOptionsDict]

@dataclass
class MappedDataSearchOptions(SheetOptions, Json.Serial[MappedDataSearchOptionsDict]):
    """
    *`@extends`* `SheetOptions`
    ```
    path: Path
    sheetname: str
    key_cols: list[str]
    val_cols: list[str]
    case_sensitive: bool=False
    col_offset: int = 0
    sheetname_selector: Callable[[str], bool] | None = None
    column_selector: Callable[[str], bool] | None = None
    - property def columns(self) -> list[str]
    ```
    """
    path: Path
    sheetname: str
    key_cols: list[str]
    val_cols: list[str]
    case_sensitive: bool=False
    col_offset: int = 0
    sheetname_selector_options: String.StringConditionOptions | None = None
    column_selector_options: String.StringConditionOptions | None = None
    
    
    @property
    def sheetname_selector(self):
        if (not self.sheetname_selector_options):
            raise RuntimeError(LogMsg(f"self.sheetname_selector_options is undefined"))
        return String.compile_condition_options(self.sheetname_selector_options)

    @property
    def column_selector(self):
        print(f"column_selector called from: ")
        print(callsite(depth=1))
        print(callsite(depth=2))
        print(callsite(depth=3))
        if (not self.column_selector_options):
            raise RuntimeError(LogMsg(f"self.column_selector_options is undefined"))
        return String.compile_condition_options(self.column_selector_options)
    
    def __init__(
        self, 
        path: fs.StrPath, 
        sheetname: str, 
        key_cols: String.StrInput, 
        val_cols: String.StrInput | None = None,
        *,
        case_sensitive: bool=False,
        col_offset: int = 0,
        sheetname_selector_options: String.StringConditionOptions | None = None,
        column_selector_options: String.StringConditionOptions | None = None
    ):
        self.path = fs.to_path(path)
        self.sheetname = sheetname
        self.key_cols = list(String.resolve_str_input(key_cols))
        self.val_cols = list(String.resolve_str_input(val_cols)) if val_cols is not None else []
        self.case_sensitive = case_sensitive
        self.col_offset = col_offset
        self.sheetname_selector_options = sheetname_selector_options
        self.column_selector_options = column_selector_options

    @classmethod
    def from_dict(cls, data: Json.JsonDict[MappedDataSearchOptionsDict], /):
        try:
            _sheetname_options = data.get('sheetname_selector_options', None)
            _column_options = data.get('column_selector_options', None)
            return cls(
                path=fs.to_path(data['path']),
                sheetname=data.get('sheetname', ''), # 'Sheet1'
                key_cols=data['key_cols'],
                val_cols=data.get('val_cols', []),
                case_sensitive=data.get('case_sensitive', False),
                col_offset=data.get('col_offset', 0),
                sheetname_selector_options = (
                    String.StringConditionOptions.from_dict(_sheetname_options) 
                    if _sheetname_options is not None 
                    else None
                ),
                column_selector_options = (
                    String.StringConditionOptions.from_dict(_column_options) 
                    if _column_options is not None
                    else None
                ),
            )
        except (KeyError, ValueError): raise
        except Exception as e:
            raise RuntimeError(LogMsg.err(
                error=e, 
                received=str(data)
            )) from e
        
    @property
    def columns(self):
        """ *`@returns`* `[*self.key_cols, *self.val_cols]` """
        return [*self.key_cols, *self.val_cols]

    def get_column_selection(self, candidates: Iterable[str], /) -> list[str]:
        """ 
        - does not modify `self.val_cols` 
        - *`@raises`* `RuntimeError` if selector is undefined
        """
        if (not self.column_selector_options):
            raise RuntimeError(LogMsg(f"self.colu,m_selector_options is undefined"))
        return Array.filter(candidates, self.column_selector)

    def get_sheetname_selection(self, candidates: Iterable[str]) -> list[str]:
        """ 
        - does not modify `self.sheetname` 
        - *`@raises`* `RuntimeError` if selector is undefined
        """
        if (not self.sheetname_selector_options):
            raise RuntimeError(LogMsg(f"self.sheetname_selector_options is undefined"))
        return Array.filter(candidates, self.sheetname_selector)

@dataclass
class MappedData(): # i.e. MappedDataSearchResult
    """
    ```
    rows: list[list[Any]]
    columns: list[str] # = [*MappedDataSearchOptions.columns, '__row_index__']
    header: HeaderIndex # Header detected in MappedDataSearchOptions path[sheetname]
    ```
    """
    rows: list[list[Any]]
    columns: list[str]
    header: HeaderIndex # from the sheet
    """ 
    use `header.columns[self.columns[i]]` and `rows[i][indexOf('__row_index__')]` 
    to reconstruct `CellAddress` in original sheet
    - where `columns[i] != '__row_index__'` and `columns[i]` in `MappedDataSearchOptions.columns`
    """
