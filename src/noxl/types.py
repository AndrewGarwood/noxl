"""
*`@file`* `noxl/types.py`
"""
from __future__ import annotations
from .constants import vba
from .models import *
from dataclasses import dataclass
from typing import Any
from pyqi import (BaseIntEnum, MetaEnum, SortOptions)


class CalculationState(BaseIntEnum): # from vba Eenum XlCalculationState
    DONE = vba.xlDone
    CALCULATING = vba.xlCalculating
    PENDING = vba.xlPending


class LayoutForm(BaseIntEnum):
    TABULAR = vba.xlTabular
    OUTLINE = vba.xlOutline

class ExcelSortOrder(BaseIntEnum):
    """
    ```
    ASC = vba.xlAscending # 1
    DSC = vba.xlDescending # 2
    ```
    """
    ASC = vba.xlAscending
    DSC = vba.xlDescending


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

    SUM = (vba.xlSum, FunctionMeta("Sum", 1))
    COUNT = (vba.xlCount, FunctionMeta("Count", 2))
    AVERAGE = (vba.xlAverage, FunctionMeta("Average", 3))
    MAX = (vba.xlMax, FunctionMeta("Max", 4))
    MIN = (vba.xlMin, FunctionMeta("Min", 5))
    PRODUCT = (vba.xlProduct, FunctionMeta("Product", 6))
    COUNT_NUMS = (vba.xlCountNums, FunctionMeta("Count Numbers", 7))
    STDEV = (vba.xlStDev, FunctionMeta("StdDev", 8))
    STDEVP = (vba.xlStDevP, FunctionMeta("StdDevP", 9))
    VAR = (vba.xlVar, FunctionMeta("Var", 10))
    VARP = (vba.xlVarP, FunctionMeta("VarP", 11))


    @property
    def code(self) -> int:
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

    @staticmethod
    def to_subtotals_payload(*fn: ExcelAggregation) -> list[bool]:
        """
        *`@returns`* `bool[]<12>`

        *`@example`*
        ```
        >>> ExcelAggregation.to_subtotals_payload(ExcelAggregation.SUM, ExcelAggregation.COUNT)
        [False, True, True, False, False, False, False, False, False, False, False, False]
        >>>
        ```
        """
        payload: list[bool] = [False] * 12
        for f in fn:
            payload[f.subtotals_index] = True
        return payload

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
