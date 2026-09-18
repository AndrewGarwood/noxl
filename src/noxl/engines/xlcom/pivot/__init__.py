"""
*`@module`* `engines/xlcom/pivot/__init__`
"""

from .field import *
from .options import *
from .table import *
from .slicer import (SlicerOptions, SlicerPosition, SlicerType)


# temp alias for backwards compatibility
PivotTableFields = Fields
PivotFieldOptions = AxisFieldOptions
PivotFieldType = FieldType

PivotTableOptions = CreateOptions
PivotTableSource = TableSource
PivotTableDestination = TableDestination