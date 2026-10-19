"""
*`@file`* `engines/xlcom/workbook/_sheets`
"""
from __future__ import annotations
from win32com.client import CDispatch
from lute import LogMsg
from typing import Optional
from pyqi import export, safe, Array

def add_worksheet(
    wb: CDispatch, 
    sheetname: Optional[str], 
    /,
    after: Optional[int]=None
    # TODO: change after to dst_anchor_key: str | int | None=None
    # and add param dst_anchor_how: "Before" | "After"="After"
) -> CDispatch:
    after = wb.Worksheets.Count if after is None else after % wb.Worksheets.Count
    ws: CDispatch = wb.Worksheets.Add(After=after)
    if sheetname is not None: ws.Name = sheetname
    return ws

def resolve_dst_sheet(wb: CDispatch, base: str, use_existing: bool=False) -> CDispatch:
    """
    Retrieves the target worksheet. If base_name exists, searches for the first 
    available indexed name (e.g., Pivot0, Pivot1) and creates it.
    """
    existing_sheets = get_sheetnames(wb)
    # mlog.debug(f"{base in existing_sheets=}")
    # mlog.debug(f"{use_existing=}")
    if base not in existing_sheets:
        return add_worksheet(wb, base)
    if use_existing:
        return wb.Worksheets(base)
    idx: int = 0
    while f"{base}{idx}" in existing_sheets:
        idx += 1
    return add_worksheet(wb, f"{base}{idx}")

def get_worksheet(wb: CDispatch, key: str | int, /) -> CDispatch:
    """
    i.e. `get_existing_worksheet`
    - *`@returns`* `ws (CDispatch)`
    - *`@throws`* `RuntimeError` if error from `wb.Worksheets(key)`
    """
    try:
        return wb.Worksheets(key)
    except Exception as e:
        raise RuntimeError(LogMsg.err(
            error=e,
            key=key,
            wb_FullName=safe.readattr(wb, "FullName")
        )) from e

def get_sheetnames(wb: CDispatch, /) -> list[str]:
    try:
        return Array.map(wb.Worksheets, lambda ws: str(ws.Name))
    except Exception as e:
        raise RuntimeError(LogMsg.err(
            e, 
            "Failed to read sheetnames of Excel workbook com object",
            wb_name=safe.readattr(wb, "FullName")
        )) from e

__all__ = export(add_worksheet, resolve_dst_sheet, get_worksheet, get_sheetnames)