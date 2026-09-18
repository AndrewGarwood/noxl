"""
*`@file`* `engines/xlcom/payload`
"""

from datetime import time as dt_time
from typing import (Sequence, Any)
from pyqi import Array


def to_com_payload(rows: Sequence[Sequence[Any]],/) -> tuple[tuple[Any, ...], ...]:
    """ COM sanitization b/c can't handle types like datetime.time """
    return tuple(
        tuple(Array.map(r, _to_com_safe))
        for r in rows
    )

def _to_com_safe(value: Any, /) -> Any: # incomplete
    if value is None:
        return None
    if isinstance(value, dt_time):
        return value.strftime("%I:%M:%S %p")
    return value

def is_empty_cell_value(
    value: Any, 
    /, 
    empty_values: list[Any]=[None, ""]
) -> bool:
    """ *`@returns`* `Array.some(empty_values, lambda v: v == value)` """
    return Array.some(empty_values, lambda v: v == value)
