"""
*`@file`* `engines/xlcom/workbook/_sheets`
"""

from __future__ import annotations
from win32com.client import CDispatch
from lute import LogMsg
from typing import Optional


def add_worksheet(
    wb: CDispatch, 
    sheetname: Optional[str], 
    /,
    after: Optional[int]=None
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
    existing_sheets: list[str] = [sheet.Name for sheet in wb.Worksheets]
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

def get_worksheet(wb: CDispatch, name: str, /) -> CDispatch:
    """
    i.e. `get_existing_worksheet`
    - *`@returns`* `ws (CDispatch)`
    - *`@throws`* `RuntimeError` if error from `wb.Worksheets(name)`
    """
    try:
        return wb.Worksheets(name)
    except Exception as e:
        raise RuntimeError(LogMsg.err(
            error=e,
            name=name,
            wb_FullName=wb.FullName
        )) from e