"""
*`@file`* `engines/xlcom/pivot/slicer`
"""

from dataclasses import (dataclass, field as _field)
from typing import Optional
from pyqi.models import BaseIntEnum
from ....constants.vba import (xlSlicer, xlTimeline)



class SlicerType(BaseIntEnum):
    SLICER = xlSlicer
    TIMELINE = xlTimeline

@dataclass
class SlicerPosition:
    """
    Layout options for Slicers in Points (1 point = 1/72 inch). 
    (b/c excel uses this unit for some reason)
    ```
    top: float = 100.0
    left: float = 400.0
    width: float = 150.0
    height: float = 200.0
    ```
    """
    top: float = 100.0
    left: float = 400.0
    width: float = 150.0
    height: float = 200.0
@dataclass
class SlicerOptions:
    """
    ```
    field: str
    caption: Optional[str] = None
    type: SlicerType = SlicerType.SLICER
    position: Optional[SlicerPosition] = None
    selection: list[str]=None
    ```

    *`@note`* `slicerCache = Workbook.SlicerCaches.Add2(Source, SourceField, [Name], [SlicerCacheType])`
    - `Source (Required):` The data source object, such as a PivotTable, a ListObject (Excel Table), or a connection string.
    - `SourceField (Required):` The name of the field or column to filter by
    - `Name (Optional):` The unique internal name for the slicer cache. If omitted, Excel auto-generates one.
    - `SlicerCacheType (Optional):` Specifies the type of slicer cache. Use xlSlicer (default) or xlTimeline for date timelines.
    """
    field: str
    caption: Optional[str] = None
    type: SlicerType = SlicerType.SLICER
    position: Optional[SlicerPosition] = None
    selection: list[str]=_field(default_factory=list)

    def __post_init__(self) -> None:
        if self.caption is None:
            self.caption = self.field