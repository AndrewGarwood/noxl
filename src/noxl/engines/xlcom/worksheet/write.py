"""
*`@file`* `enginges/xlcom/worksheet/write`
"""
from __future__ import annotations
from typing import (Any)
import win32com.client as win32
from lute import LogMsg
from pyqi import Array
from ....models import *

from .read import (range_by_dimensions)
from ..payload import (to_com_payload)


def clear_range_values(
    ws: win32.CDispatch, 
    /,
    origin: CellAddressLike,
    num_rows: int, 
    num_cols: int 
):
    try:
        _range = range_by_dimensions(ws, origin, num_rows, num_cols)
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
            rows=f"{type(rows)=}, {len(rows)=}"
        )) from e
    pass
