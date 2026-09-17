"""
*`@file`* `noxl/types.py`
"""
from __future__ import annotations
from .constants.vba import *
from dataclasses import dataclass
from typing import Any
from pyqi import (BaseIntEnum, MetaEnum, SortOptions)


class CalculationState(BaseIntEnum): # from vba Eenum XlCalculationState
    DONE = xlDone
    CALCULATING = xlCalculating
    PENDING = xlPending


class LayoutForm(BaseIntEnum):
    TABULAR = xlTabular
    OUTLINE = xlOutline

class ExcelSortOrder(BaseIntEnum):
    """
    ```
    ASC = xlAscending # 1
    DSC = xlDescending # 2
    ```
    """
    ASC = xlAscending
    DSC = xlDescending


class ExcelSortOptions(SortOptions[ExcelSortOrder]):
    ...

@dataclass(frozen=True)
class FunctionMeta:
    display_name: str
    subtotals_payload_index: int=0

# pylance error: ExcelAggregation(MetaEnum[int, FunctionMeta])
class ExcelAggregation(MetaEnum):
    """ `MetaEnum[int, FunctionMeta]` """
    _value_: int
    meta: FunctionMeta

    SUM = (xlSum, FunctionMeta("Sum", 1))
    COUNT = (xlCount, FunctionMeta("Count", 2))
    AVERAGE = (xlAverage, FunctionMeta("Average", 3))
    MAX = (xlMax, FunctionMeta("Max", 4))
    MIN = (xlMin, FunctionMeta("Min", 5))
    PRODUCT = (xlProduct, FunctionMeta("Product", 6))
    COUNT_NUMS = (xlCountNums, FunctionMeta("Count Numbers", 7))
    STDEV = (xlStDev, FunctionMeta("StdDev", 8))
    STDEVP = (xlStDevP, FunctionMeta("StdDevP", 9))
    VAR = (xlVar, FunctionMeta("Var", 10))
    VARP = (xlVarP, FunctionMeta("VarP", 11))


    @property
    def code(self) -> int:
        if not isinstance(self._value_, int): 
            raise RuntimeError(f"@ self.code, ExcelAggregation._value_ is not int")
        return self._value_

    @property
    def subtotals_index(self):
        return self.meta.subtotals_payload_index

    def field_text(self, field: str, /):
        """ *`@returns`* `{self.name} of {field}` """
        return f"{self.meta.display_name} of {field}"

    @classmethod
    def from_code(cls, code: int, /) -> ExcelAggregation:
        return cls.from_value(code)


class ExcelSaveTimeoutError(TimeoutError):
    """Raised when an Excel workbook save exceeds the configured timeout."""
    ...

class ExcelAutomationError(RuntimeError):
    ...


EXCEL_CVERR_VALUES: set[int] = {
    -2146826281,  # #DIV/0!
    -2146826246,  # #N/A
    -2146826259,  # #NAME?
    -2146826288,  # #NULL!
    -2146826252,  # #NUM!
    -2146826265,  # #REF!
    -2146826273,  # #VALUE!
}

def is_excel_error_value(value: Any, /) -> bool:
    return (
        isinstance(value, int)
        and not isinstance(value, bool)
        and (
            value in EXCEL_CVERR_VALUES
            or -2146827000 < value < -2146826000
        )
    )
