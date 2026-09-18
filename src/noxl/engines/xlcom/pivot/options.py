"""
*`@file`* `engines/xlcom/pivot/options`
"""
from __future__ import annotations
from dataclasses import (dataclass, field as _field)
from typing import (Optional)
from win32com.client import (CDispatch)
from lute import (LogMsg)
from .slicer import (SlicerOptions)
from .field import (
    Fields
)
from ..types import (ComWorkbookSource)
from ....types import (SheetRange, CellAddressLike, CellAddress, ExcelSortOptions)

# TODO: maybe mixin Json.Serial if eventually want to read from config files...

@dataclass
class TableSource:
    """
    ```
    sheetname: Optional[str] = None
    tablename: Optional[str] = None
    origin: Optional[CellAddressLike] = None
    num_rows: Optional[int] = None
    num_cols: Optional[int] = None
    range_obj: Optional[CDispatch] = None
    range_spec: Optional[SheetRange] = None

    resolve(self, wb: CDispatch, /) -> str | CDispatch: # get SourceData for `PivotCaches().Create(SourceData=...)`
    ```
    """
    sheetname: Optional[str] = None
    tablename: Optional[str] = None
    origin: Optional[CellAddressLike] = None
    num_rows: Optional[int] = None
    num_cols: Optional[int] = None
    range_obj: Optional[CDispatch] = None
    range_spec: Optional[SheetRange] = None
    
    def __post_init__(self) -> None:
        if self.range_spec is not None:
            self.sheetname = self.range_spec.sheetname
            self.origin = self.range_spec.origin
            self.num_cols = self.range_spec.num_cols
    
    def resolve(self, wb: CDispatch, /) -> str | CDispatch:
        """
        Returns the exact COM object or string required by `PivotCaches().Create(SourceData=...)`.
        """
        try:
            # Path 1: Direct COM Range object
            if self.range_obj is not None:
                return self.range_obj
            # Path 2: Structured Excel Table (ListObject)
            if self.tablename is not None:
                if self.sheetname is not None:
                    ws: CDispatch = wb.Worksheets(self.sheetname)
                    # Return the COM Range of the table to guarantee exact grid binding
                    return ws.ListObjects(self.tablename).Range
                return self.tablename
            # Path 3: Coordinate origin expansion
            if (self.sheetname is None or self.origin is None):
                raise ValueError(
                    f"PivotTableSource requires either range_obj, tablename, or (sheetname + origin)."
                )
            ws: CDispatch = wb.Worksheets(self.sheetname)
            addr = CellAddress.from_any(self.origin)
            start_cell: CDispatch = ws.Cells(addr.row, addr.col)
            # If dimensions are explicit, use the excel_range utility
            if (self.num_rows is not None and self.num_cols is not None):
                end = addr.offset(rows=self.num_rows-1, cols=self.num_cols-1)
                return ws.Range(
                    start_cell,
                    ws.Cells(end.row, end.col)
                )
            return start_cell.CurrentRegion
        except Exception as e:
            raise RuntimeError(LogMsg.err(
                e, 
                self=str(self), 
                wb_name=getattr(wb, "Name", None)
            )) from e


@dataclass
class TableDestination():
    """
    ```
    sheetname: str = "Pivot"
    address: CellAddressLike = "A3"
    tablename: Optional[str] = None
    use_existing_sheet: bool = False
    ```
    """
    sheetname: str = "Pivot"
    addr: CellAddressLike = "A3"
    tablename: Optional[str] = None
    use_existing_sheet: bool = False

@dataclass
class CreateOptions:
    """
    ```
    wb_source: ComWorkbookSource
    src: PivotTableSource
    dst: PivotTableDestination = field(default_factory=PivotTableDestination)
    fields: TableFieldOptions = field(default_factory=TableFieldOptions)
    slicers: list[SlicerOptions] = _field(default_factory=list)
    sort: Optional[SortOptions] = None
    save_on_close: bool = True
    row_grand: bool = True
    col_grand: bool = True
    ```
    """
    wb_source: ComWorkbookSource
    #TODO, see if ExcelSession can/should acquire external com objects and store in ExcelSession._workbooks
    src: TableSource
    dst: TableDestination = _field(default_factory=TableDestination)
    fields: Fields = _field(default_factory=Fields)
    slicers: list[SlicerOptions] = _field(default_factory=list)
    sort: Optional[ExcelSortOptions] = None
    # repeat_labels: bool = True
    save_on_close: bool = True
    row_grand: bool = True
    col_grand: bool = True

