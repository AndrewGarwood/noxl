"""
*`@file`* `enginges/xlcom/worksheet/write`
"""
from __future__ import annotations
from typing import (Any, Optional)
import win32com.client as win32
from lute import LogMsg, annotate
from pyqi import Array
from ....models import *

from .read import (range_by_dimensions, get_range)
from ..payload import (to_com_payload)


def clear_range(
    ws: win32.CDispatch, 
    /,
    origin: CellAddressLike,
    num_rows: Optional[int]=None, 
    num_cols: Optional[int]=None 
):
    try:
        _range = (
            get_range(ws, origin) if (num_rows is None or num_cols is None)
            else range_by_dimensions(ws, origin, num_rows, num_cols)
        )
        _range.ClearContents()
    except Exception as e:
        raise Exception(LogMsg.err(
            error=e,
            origin=str(origin),
            num_rows=num_rows,
            num_cols=num_cols
        )) from e

def write_rows(
    ws: win32.CDispatch,
    /,
    rows: Sequence[Sequence[Any]],
    origin: CellAddressLike = "A2"
) -> None:
    if not rows:
        return
    try:
        M = len(rows)
        N = len(rows[0])
        _range = range_by_dimensions(ws, origin, num_rows=M, num_cols=N)
        _range.Value = to_com_payload(Array.matrix.standardize_row_length(rows))
    except Exception as e:
        raise Exception(LogMsg(
            f"Failed to write rows to Range in Excel Worksheet",
            error=e,
            origin=str(origin),
            rows=annotate(rows)
        )) from e
    pass
