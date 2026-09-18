"""
*`@file`* `engines/xlcom/workbook/_names`
"""
from typing import (Any)
from win32com.client import CDispatch


def entry(wb: CDispatch, name: str, /) -> CDispatch:
    """ 
    *`@returns`* `wb.Names(name)`
    """
    obj = wb.Names(name)
    # slog.debug(f"{lsource(__file__, get_named_obj)}, obj = wb.Names(name) {type(obj)=}")
    return obj

def value(wb: CDispatch, name: str, /) -> Any:
    """ 
    *`@returns`* `wb.Names(name).Value`
    - use when name is formula-defined workbook name whose formula evaluates directly to a scalar value. 
    """
    return wb.Names(name).Value

def ref(wb: CDispatch, name: str, /) -> Any:
    """ 
    *`@returns`* `wb.Names(name).RefersTo`
    - use when name is formula-defined workbook name whose formula evaluates directly to a scalar value. 
    """
    return wb.Names(name).RefersTo

def ref_range(wb: CDispatch, name: str, /) -> Any:
    """ 
    *`@returns`* `wb.Names(name).RefersToRange`
    - use when name is formula-defined workbook name whose formula evaluates directly to a scalar value. 
    """
    return wb.Names(name).RefersToRange

def ref_range_value(wb: CDispatch, name: str, /) -> Any:
    """ 
    *`@returns`* `wb.Names(name).RefersToRange.Value`
    - use when name is formula-defined workbook name whose formula evaluates directly to a scalar value. 
    """
    return wb.Names(name).RefersToRange.Value