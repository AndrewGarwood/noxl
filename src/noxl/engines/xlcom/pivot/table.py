"""
*`@file`* `engines/xlcom/pivot/table`
"""
from __future__ import annotations
import time
from pathlib import Path
from typing import (Sequence)
from lute import LogMsg, get_logger
from pyqi import Array
from noxl.types import ExcelAggregation
from .options import *
from .field import (
    FieldType, 
    FilterFieldOptions, 
    AxisFieldOptions, 
    AxisFieldType
)
from .slicer import (SlicerType)
from noxl.constants.vba import *
from noxl.models._cell import is_empty_cell_value
from noxl.engines.xlcom.types import (TableReference)
from .. import workbook
from ..session import (ExcelSession)

mlog = get_logger(Path(__file__).parent.parent.stem)

def refresh(pt: CDispatch | TableReference) -> bool:
    if (isinstance(pt, TableReference)):
        with ExcelSession() as xl:
            wb = xl.resolve_wb_source(pt.wb_source)
            table = wb.Worksheets(pt.sheetname).PivotTables(pt.tablename)
            return bool(table.RefreshTable())
    else:
        return bool(pt.RefreshTable())
    return False

# TODO maybe add param 'strict' or '_raise' so can decide if want to throw error when there is
# something wrong
def create(options: CreateOptions, /): 
    try:
        with ExcelSession() as xl:
            wb = xl.resolve_wb_source(options.wb_source)
            src_data = options.src.resolve(wb)
            dst_ws = workbook.resolve_dst_sheet(
                wb, 
                options.dst.sheetname, 
                options.dst.use_existing_sheet
            )
            dst_addr = CellAddress.from_any(options.dst.addr)
            dst_cell: CDispatch = dst_ws.Cells(dst_addr.row, dst_addr.col)
            dst_cell_value = dst_cell.Value
            if (not is_empty_cell_value(dst_cell_value)):
                raise RuntimeError(LogMsg(
                    f"Destination cell is not empty or None",
                    f"@ sheet {options.dst.sheetname}",
                    f"@  cell {dst_addr.row=}, {dst_addr.col=}",
                    f"current value = {dst_cell_value}"
                ))
            dst_tablename: str=f"Table_R{dst_addr.row}C{dst_addr.col}"
            if options.dst.tablename is not None:
                existing_pivots: list[str] = [pt.Name for pt in dst_ws.PivotTables()]
                if options.dst.tablename in existing_pivots:
                    raise ValueError(
                        f"PivotTable '{options.dst.tablename}' already exists on sheet '{dst_ws.Name}'"
                    )
                dst_tablename = options.dst.tablename
            else:
                src_label: str
                if isinstance(src_data, str):
                    src_label = src_data
                else:
                    clean_addr: str = (
                        str(src_data.Address(False, False)) 
                        if not isinstance(src_data.Address, str) 
                        else src_data.Address
                        .replace(':', '_')
                        .replace(' ', '')
                    )
                    src_label = f"{src_data.Worksheet.Name}_{clean_addr}"
                dst_tablename = f"PivotTable_{src_label}_{dst_ws.PivotTables().Count + 1}"
            pivot_cache: CDispatch = wb.PivotCaches().Create(
                SourceType=xlDatabase, 
                SourceData=src_data,
                Version=8
            )
            pt: CDispatch = pivot_cache.CreatePivotTable(
                TableDestination=dst_cell,
                TableName=dst_tablename,
                DefaultVersion=8
            )
            pt.ColumnGrand = options.col_grand
            pt.RowGrand = options.row_grand
            pt.DisplayErrorString = False
            pt.DisplayNullString = True
            pt.NullString = ""
            pt.PageFieldOrder = 2
            pt.RowAxisLayout(xlCompactRow)
            time.sleep(1)
            _handle_fields(pt, options.fields)
            time.sleep(1)
            if (options.sort):
                pt.PivotFields(options.sort.on).AutoSort(
                    options.sort.order,
                    options.sort.by
                )
            if (options.slicers):
                _handle_slicers(wb, dst_ws, pt, options.slicers)
            xl.save_workbook(wb)
            time.sleep(1)
            xl.close_workbook(wb)
    except Exception as e:
        raise RuntimeError(LogMsg.err(
            e, 
            options=str(options)
        )) from e

def _handle_fields(pt: CDispatch, fields: Fields, /) -> None:
    try:
        if (fields.rows): 
            _handle_axis_fields(pt, fields.rows, FieldType.ROW)
        if (fields.columns): 
            _handle_axis_fields(pt, fields.columns, FieldType.COLUMN)
        if (fields.filters):
            _handle_filter_fields(pt, fields.filters)
        if (fields.values): 
            Array.forEach(
                fields.values, 
                lambda f: pt.AddDataField(
                    pt.PivotFields(f.field),
                    f.display,
                    f.pivot_fn.code
                )
            )
    except Exception as e:
        raise RuntimeError(LogMsg.err(e, fields=str(fields)))

def _handle_axis_fields(
    pt: CDispatch, 
    fields: Sequence[AxisFieldOptions],
    /, 
    axis: AxisFieldType
) -> None:
    for i, f in enumerate(fields, start=1):
        pf = pt.PivotFields(f.field)
        pf.Orientation = axis
        pf.Position = i
        pf.LayoutForm = f.layout
        pf.RepeatLabels = f.repeat_labels
        if (f.subtotals):
            pf.Subtotals = ExcelAggregation.to_subtotals_payload(*f.subtotals)
        if (not f.item_positions): continue
        pivot_items = set(Array.map(pf.PivotItems(), lambda pi: str(pi.Value).casefold()))
        for item, pos in f.item_positions.items():
            invalid_entry = (
                not isinstance(pos, int) 
                or str(item).casefold() not in pivot_items
            )
            if (invalid_entry): continue
            if (pos < 1): pos = 1
            if (pos > len(pivot_items)): pos = len(pivot_items)
            pf.PivotItems(item).Position = pos
    return None

def _handle_filter_fields(
    pt: CDispatch, 
    fields: Sequence[FilterFieldOptions], 
    /
) -> None:
    for i, f in enumerate(fields, start=1):
        pf = pt.PivotFields(f.field)
        pf.Orientation = FieldType.FILTER
        pf.Position = i
        if (f.enable_multiple_page_items is not None):
            pf.EnableMultiplePageItems = f.enable_multiple_page_items
        if (f.selection):
            hidden_count: int = 0
            pf.EnableMultiplePageItems = True
            selection = set(Array.map(f.selection, lambda s: s.casefold()))
            N = len(selection)
            for pi in pf.PivotItems():
                if (hidden_count >= N - 1):
                    mlog.debug(LogMsg(
                        f"Reached limit for number of hidden PivotItems",
                        field=f.field,
                        selection=f"string[]<{len(f.selection)}> = {f.selection}",
                        num_pivot_items=N,
                        hidden_count=hidden_count
                    ))
                    break
                if (str(pi.Value).casefold() in selection): 
                    continue # default Visible=True
                pi.Visible = False
                hidden_count += 1
    return None

def _handle_slicers(
    wb: CDispatch,
    ws: CDispatch,
    pt: CDispatch,
    slicers: Sequence[SlicerOptions],
    /
) -> None:
    current_top: float = 20.0
    default_left: float = 500.0
    default_width: float = 180.0
    default_height: float = 200.0
    spacing: float = 15.0
    for i, slc in enumerate(slicers, start=1):
        try:
            slc_cache: CDispatch = wb.SlicerCaches.Add2(
                pt, 
                slc.field,
                SlicerCacheType=slc.type.value
            )
            pos = slc.position
            top = pos.top if pos else current_top
            left = pos.left if pos else default_left
            width = pos.width if pos else default_width
            height = pos.height if pos else (
                100.0 if slc.type == SlicerType.TIMELINE 
                else default_height
            )
            slc_cache.Slicers.Add(
                SlicerDestination=ws,
                Name=f"Slicer_{pt.Name}_{i}_{slc.field}",
                Caption=str(slc.caption or slc.field),
                Top=top,
                Left=left,
                Width=width,
                Height=height
            )
            if (slc.selection):
                keys = Array.map(slc.selection, lambda s: s.casefold())
                items = list(slc_cache.SlicerItems)
                for item in items:
                    item_key = str(item.Name).casefold()
                    slc_cache.SlicerItems(item_key).Selected = item_key in keys
            current_top += height + spacing
        except Exception as e:
            raise RuntimeError(LogMsg.err(
                e,
                f"Error handling SlicerOptions ({i} / {len(slicers)})",
                slicer_options=str(slc),
                wb=str(getattr(wb, "FullName", "<name unretrievable>")),
                ws=str(getattr(ws, "Name", "<name unretrievable>"))
            ))



__all__ = [ 
    *[x.__name__ for x in (refresh, create)], # type: ignore
]