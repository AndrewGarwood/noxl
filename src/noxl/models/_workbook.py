"""
*`@file`* `models/_workbook`
"""
from typing import (TypeAlias, TypeVar)
from pyqi.fs._path import StrPath

T = TypeVar('T')
WorkbookSource: TypeAlias = T | StrPath