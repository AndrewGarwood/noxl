"""
*`@module`* `constants/__init__`
"""

from .vba import *
from .formatting import *
from datetime import datetime

EXCEL_EPOCH = datetime(1899, 12, 30)
""" Excel's default Windows date system is: `1899-12-30` """
"""
because Excel intentionally preserves a 
historical Lotus 1-2-3 bug where Excel incorrectly believes 1900 was a leap year.
"""

DEFAULT_EXE = "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE"