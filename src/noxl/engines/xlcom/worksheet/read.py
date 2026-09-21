"""
*`@file`* `engines/xlcom/worksheet/read`
"""
from __future__ import annotations
from typing import (Any, Optional)
from  win32com.client import CDispatch
from lute import LogMsg, annotate
from ....models import (CellAddress, CellAddressLike)

def range_by_dimensions(
    ws: CDispatch, 
    origin: CellAddressLike,
    /,
    num_rows: int, 
    num_cols: int 
) -> CDispatch:
    if num_rows <= 0:
        raise ValueError(f"num_rows must be positive. Got {num_rows!r}")
    if num_cols <= 0:
        raise ValueError(f"num_cols must be positive. Got {num_cols!r}")
    
    start = CellAddress.from_any(origin)
    end = start.offset(rows=num_rows-1, cols=num_cols-1)
    _range: CDispatch = ws.Range(
        ws.Cells(start.row, start.col), 
        ws.Cells(end.row, end.col)
    )
    return _range


def get_value(ws: CDispatch, addr: CellAddressLike, /) -> Any:
    """ read value from cell """
    try:
        return get_cell(ws, addr).Value
    except Exception as e:
        raise Exception(LogMsg.err(
            error=e,
            addr=str(addr),
            ws=f"{type(ws)=}"
        )) from e

def get_cell(ws: CDispatch, addr: CellAddressLike="A2", /) -> CDispatch:
    addr = CellAddress.from_any(addr)
    return ws.Cells(addr.row, addr.col)


def get_range(
    ws: CDispatch, 
    origin: CellAddressLike="A2", 
    end: Optional[CellAddressLike]=None,
    /,
    use_current_region: bool=False
) -> CDispatch:
    start_cell = get_cell(ws, origin)
    if (end is not None):
        return ws.Range(start_cell, get_cell(ws, end))
    return (
        ws.Range(start_cell).CurrentRegion if use_current_region 
        else ws.Range(start_cell)
    )


def get_current_region(
    ws: CDispatch, 
    addr: CellAddressLike="A2", 
    /,
    include_header: bool=True
) -> CDispatch:
    """ `typeof range.CurrentRegion == Range` """
    region = ws.Range(get_cell(ws, addr)).CurrentRegion
    return (
        region if include_header 
        else region.Offset(1, 0).Resize(region.Rows.Count - 1, region.Columns.Count)
    )

def get_current_region_dimensions(region: CDispatch, /) -> tuple[int, int]: 
    """ 
    - *`@param`* `region (CDispatch)` the return value of `get_current_region(...)`
    - *`@returns`* `dimensions (tuple[int, int]) = (row_count, col_count)` 
    """
    try:
        row_count = region.Rows.Count
        if (not isinstance(row_count, int)):
            raise RuntimeError(LogMsg(
                "region.Rows.Count = rowcount is not an integer",
                row_count=annotate(row_count)
            ))
        col_count = region.Columns.Count
        if (not isinstance(row_count, int)):
            raise RuntimeError(LogMsg(
                "region.Columns.Count = col_count is not an integer",
                col_count=annotate(col_count)
            ))
        return (row_count, col_count)
    except Exception as e:
        raise RuntimeError(LogMsg.err(
            e,
            region=annotate(region)
        ))
