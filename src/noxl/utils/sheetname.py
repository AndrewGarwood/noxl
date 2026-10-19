"""
*`@file`* `utils/sheetname`
"""

from pyqi import (String, export)
from lute import LogMsg, annotate
from collections.abc import Iterable

_MAX_ITER = 500
def get_resolved_sheetname(
    base_sheetname: str, 
    existing_sheetnames: Iterable[str], 
    /,
    copy_suffix_fmt: str = "({})",
    start_index: int = 1
):
    existing = set(existing_sheetnames)
    if (base_sheetname not in existing):
        return base_sheetname
    if (not copy_suffix_fmt or not String.contains(copy_suffix_fmt, "{}")):
        raise ValueError(LogMsg.bad_arg(
            copy_suffix_fmt, 
            'copy_suffix_fmt', 
            "Missing curly braces for formatting copy index"
        ))
    def propose_name(base: str, copy_fmt: str, index: int, /):
        return f"{base}" + copy_fmt.format(index)
    index: int = start_index
    sheetname = propose_name(base_sheetname, copy_suffix_fmt, index)
    _MAX_INDEX = _MAX_ITER + start_index
    while sheetname in existing:
        if (index > (_MAX_INDEX)):
            raise RuntimeError(LogMsg(
                f"Unable to resolve sheetname after {_MAX_ITER} iterations",
                base_sheetname=base_sheetname,
                last_propsed_sheetname=sheetname,
                existing_sheetnames=annotate(existing),
                copy_suffix_fmt=copy_suffix_fmt,
                start_index=start_index,
                last_index=index
            ))
        index += 1
        sheetname = propose_name(base_sheetname, copy_suffix_fmt, index)

    return sheetname

__all__ = export(get_resolved_sheetname)