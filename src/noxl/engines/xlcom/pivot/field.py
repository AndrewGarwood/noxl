"""
*`@file`* `engines/xlcom/pivot/slicer`
"""

from dataclasses import (dataclass, field as _field)
from typing import Literal, Mapping, Optional, Sequence

from pyqi.models import BaseIntEnum
from ....constants.vba import *
from ....types import (ExcelAggregation, ExcelSortOptions, LayoutForm)
from pyqi import Array, Arbitrary, resolve_arbitrary

class PivotFieldType(BaseIntEnum):
    """
    ```
    ROW = xlRowField
    COLUMN = xlColumnField
    FILTER = xlPageField
    VALUE = xlDataField
    ```
    """
    ROW = xlRowField
    """ `xlRowField: int = 1` """
    COLUMN = xlColumnField
    """ `xlColumnField: int = 2` """
    FILTER = xlPageField
    """ `xlPageField: int = 3` """
    VALUE = xlDataField
    """ `xlDataField: int = 4` """

@dataclass
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
    _display: Optional[str]=None

    @property
    def display(self):
        return self._display or self.pivot_fn.field_text(self.field)


@dataclass
class PivotFieldOptions:
    """
    ```
    field: str
    subtotals: list[ExcelAggregation]
    layout: LayoutForm = LayoutForm.OUTLINE
    repeat_labels: bool=False
    type: Optional[PivotFieldType]=None # all types except for `xlDataField` 
    item_positions: dict[str, int]
    ```
    """
    field: str
    subtotals: list[ExcelAggregation] = _field(default_factory=list)
    """ what agg functions to have calculate for self.field """
    layout: LayoutForm = LayoutForm.OUTLINE
    repeat_labels: bool=False
    type: Optional[Literal[PivotFieldType.ROW, PivotFieldType.COLUMN, PivotFieldType.FILTER]]=None
    item_positions: dict[str, int] = _field(default_factory=dict)
    """
    `{ [fieldItem: string]: int }`, (`1-based` indexing)
    - `ActiveSheet.PivotTables(<table_name>).PivotFields(<field_name>).PivotItems(<field_item>).Position = <position_index>`
    """
    include_new_items_in_filter: bool = False


@dataclass 
class FilterFieldOptions():
    field: str
    selection: list[str]=_field(default_factory=list)
    enable_multiple_page_items: bool=True

@dataclass
class PivotTableFields:
    """
    ```
    rows: list[PivotFieldOptions]
    columns: list[PivotFieldOptions]
    values: list[DataFieldOptions]
    filters: list[str] = field(default_factory=list) # TODO mk dataclass to allow initial selection of filter values...
    ```
    """
    rows: list[PivotFieldOptions] = _field(default_factory=list)
    columns: list[PivotFieldOptions] = _field(default_factory=list)
    values: list[DataFieldOptions] = _field(default_factory=list)
    filters: list[FilterFieldOptions] = _field(default_factory=list)

    def __init__(
        self, 
        rows: Sequence[str | PivotFieldOptions]=[], 
        columns: Sequence[str | PivotFieldOptions]=[], 
        values: Sequence[str | DataFieldOptions] | Mapping[ExcelAggregation, Arbitrary[str]]=[],
        filters: Sequence[str | FilterFieldOptions] = [],
        default_aggregation: ExcelAggregation=ExcelAggregation.SUM,
    ):
        """
        *`@param`* `values`
        """
        self.rows = Array.map(
            rows, 
            lambda f: (
                PivotFieldOptions(field=f, type=PivotFieldType.ROW) 
                if isinstance(f, str) 
                else f
            )
        )
        self.columns = Array.map(
            columns, 
            lambda f: (
                PivotFieldOptions(field=f, type=PivotFieldType.COLUMN) 
                if isinstance(f, str) 
                else f
            )
        )
        if (isinstance(values, Sequence)):
            self.values = (
                Array.map(
                    values, 
                    lambda f: (
                        DataFieldOptions(field=f, pivot_fn=default_aggregation) 
                        if isinstance(f, str) 
                        else f
                    )
                ) 
            )
        else: # is Mapping[pivot_fn, fields]
            def _acc_data_fields(acc: list[DataFieldOptions], curr: ExcelAggregation, _: int):
                x= resolve_arbitrary(values[curr])
                acc.extend(Array.map(
                    resolve_arbitrary(values[curr]), 
                    lambda f: DataFieldOptions(f, curr)
                ))
                return acc
            self.values = Array.reduce(values.keys(), _acc_data_fields, [])
        self.filters = Array.map(
            filters, 
            lambda f: (
                FilterFieldOptions(field=f) 
                if isinstance(f, str) 
                else f
            )
        )

