"""
*`@file`* `engines/xlcom/pivot/field`
"""
from __future__ import annotations
from dataclasses import (dataclass, field as _field, replace)
from typing import (Any, Iterable, Literal, Mapping, Optional, Sequence, TypeAlias, TypeGuard, Generic)
# from collections.abc import Mapping
from pyqi.models import BaseIntEnum
from noxl.constants import vba
from noxl.types import (ExcelAggregation, ExcelSortOptions, LayoutForm)
from pyqi import Array, Arbitrary, resolve_arbitrary
from lute import LogMsg, bad_arg


class FieldType(BaseIntEnum):
    """
    ```
    ROW = xlRowField
    COLUMN = xlColumnField
    FILTER = xlPageField
    VALUE = xlDataField
    ```
    """
    ROW = vba.xlRowField
    """ `xlRowField: int = 1` """
    COLUMN = vba.xlColumnField
    """ `xlColumnField: int = 2` """
    FILTER = vba.xlPageField
    """ `xlPageField: int = 3` """
    VALUE = vba.xlDataField
    """ `xlDataField: int = 4` """



@dataclass()
class DataFieldOptions:
    """
    ```
    field: str # source column name/field
    pivot_fn: ExcelAggregation
    display: Optional[str] = None # display name
    ```
    """
    field: str
    pivot_fn: ExcelAggregation
    display_name: Optional[str]=None
    @property
    def display(self):
        return self.display_name or self.pivot_fn.field_text(self.field)


AxisFieldType: TypeAlias = Literal[FieldType.ROW, FieldType.COLUMN]

@dataclass()
class AxisFieldOptions():
    """
    ```
    field: str
    subtotals: list[ExcelAggregation]
    layout: LayoutForm = LayoutForm.OUTLINE
    repeat_labels: bool=False
    item_positions: dict[str, int]
    ```
    """
    field: str
    subtotals: list[ExcelAggregation] = _field(default_factory=list)
    """ what agg functions to have calculate for self.field """
    layout: LayoutForm = LayoutForm.OUTLINE
    # type: Literal[FieldType.ROW, FieldType.COLUMN]=FieldType.ROW
    item_positions: dict[str, int] = _field(default_factory=dict)
    """
    `{ [fieldItem: string]: int }`, (`1-based` indexing)
    - `ActiveSheet.PivotTables(<table_name>).PivotFields(<field_name>).PivotItems(<field_item>).Position = <position_index>`
    """
    repeat_labels: bool=False
    include_new_items_in_filter: bool = False
    def __post_init__(self):
        # if (self.type == FieldType.VALUE):
        #     raise ValueError(LogMsg(
        #         f"FieldOptions cannot use FieldType.VALUE; use DataFieldOptions"
        #     ))
        if (Array.some(self.item_positions.values(), lambda pos: pos < 1)):
            raise ValueError(LogMsg(
                "Invalid FieldOptions.item_positions;",
                "Pivot item positions use 1-based indexing;",
                received=str(self.item_positions)
            ))
@dataclass 
class FilterFieldOptions():
    field: str
    selection: list[str]=_field(default_factory=list)
    enable_multiple_page_items: bool=True

AxisFieldOptionsInput: TypeAlias = (
    Arbitrary[str | AxisFieldOptions]
)
FilterFieldsInput: TypeAlias = (
    Arbitrary[str | FilterFieldOptions]
)
DataFieldsInput: TypeAlias = (
    Arbitrary[str | DataFieldOptions]
    | Mapping[ExcelAggregation, Arbitrary[str]]
)

def _normalize_axis_field_options(
    value: Optional[AxisFieldOptionsInput], 
    /, 
    ftype: Literal[FieldType.ROW, FieldType.COLUMN]
) -> list[AxisFieldOptions]:
    if value is None: return []
    def acc_fields(acc: list[AxisFieldOptions], curr: str | AxisFieldOptions | None, i: int):
        if (curr is None): return acc
        if isinstance(curr, str):
            acc.append(AxisFieldOptions(curr))
        elif (isinstance(curr, AxisFieldOptions)):
            acc.append(curr)
        else:
            raise TypeError(LogMsg(
                f"Invalid value @ index {i} for field_type={ftype}",
                expected="str | FieldOptions",
                received=f"{type(curr)}"
            ))
        return acc
    return Array.reduce(resolve_arbitrary(value), acc_fields, [])

def is_aggregation_mapping(
        value: Any, 
        /
    ) -> TypeGuard[Mapping[ExcelAggregation, Arbitrary[str]]]:
    return (
        isinstance(value, Mapping)
        and Array.every(
            value.keys(), 
            lambda k: (
                isinstance(k, ExcelAggregation) 
                and isinstance(value[k], (str, Iterable))
            )
        )
    )
def _normalize_data_field_options(
    value: Optional[DataFieldsInput],
    /,
    default_pivot_fn: ExcelAggregation,
) -> list[DataFieldOptions]:
    if value is None: return []
    if (isinstance(value, str)):
        return [DataFieldOptions(field=value, pivot_fn=default_pivot_fn)]
    if (isinstance(value, DataFieldOptions)):
        return [value]
    if (isinstance(value, Iterable) and not isinstance(value, Mapping)):
        return Array.filter_map(
            resolve_arbitrary(value),
            lambda f: (
                DataFieldOptions(f, default_pivot_fn) 
                if isinstance(f, str) 
                else f
            )
        )
    if (is_aggregation_mapping(value)):
        def data_fields_from_map(acc: list[DataFieldOptions], curr: ExcelAggregation, _: int):
            acc.extend(Array.map(
                resolve_arbitrary(value[curr]), 
                lambda f: DataFieldOptions(f, curr)
            ))
            return acc
        return Array.reduce(value.keys(), data_fields_from_map, [])
    raise TypeError(*bad_arg(
        value, "value", 
        "Invalid value for argument 'value'",
        "Arbitrary[str | DataFieldOptions] | Mapping[ExcelAggregation, Arbitrary[str]]"
    ))


@dataclass
class Fields:
    """
    ```
    rows: list[PivotFieldOptions]
    columns: list[PivotFieldOptions]
    values: list[DataFieldOptions]
    filters: list[FilterFieldOptions]
    default_pivot_fn: ExcelAggregation=ExcelAggregation.SUM,
    ```
    """
    rows: list[AxisFieldOptions]
    columns: list[AxisFieldOptions]
    values: list[DataFieldOptions]
    filters: list[FilterFieldOptions]

    def __init__(
        self, 
        rows: Optional[AxisFieldOptionsInput]=None, 
        columns: Optional[AxisFieldOptionsInput]=None, 
        values: DataFieldsInput | Mapping[ExcelAggregation, Arbitrary[str]] | None=None,
        filters: Optional[FilterFieldsInput] = None,
        default_pivot_fn: ExcelAggregation=ExcelAggregation.SUM,
    ):
        """
        *`@note`* param `rows`, `columns`, `filters`, `values` can be str | Iterable[str] ; 
        - strings are interpreted as column/field names
        - they are normalized to list of their respective dataclass

        *`@param`* `values`
        - if is `Arbitrary[str]` - `DataFieldOptions` with `pivot_fn=default_aggregation` created for each field specified
        - if is `Mapping[ExcelAggregation, Arbitrary[str]]`
            - a `DataFieldOption` with `pivot_fn=key` is created the fields specified in `values[key]`
        """
        self.rows = _normalize_axis_field_options(rows, FieldType.ROW)
        self.columns = _normalize_axis_field_options(columns, FieldType.COLUMN)
        self.filters = Array.map(
            resolve_arbitrary(filters), 
            lambda f: (
                FilterFieldOptions(field=f) 
                if isinstance(f, str) 
                else f
            )
        ) if filters else []
        self.values = _normalize_data_field_options(values, default_pivot_fn)