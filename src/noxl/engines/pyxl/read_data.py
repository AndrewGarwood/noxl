"""
*`@file`* `utils/read_mapped_data`
"""
from __future__ import annotations
from typing import Any, Callable, OrderedDict
from ...models import (MappedData, MappedDataSearchOptions)
from pyqi import (Array, export)
from lute import (LogMsg, annotate)
from .types import (PyxlWorkbook)

from .utils import (resolve_pyxl_wb, find_header)
from .constants import ROW_INDEX_COLUMN



def read_mapped_data(
    options: MappedDataSearchOptions,
    /,
    data_only: bool = True, 
    scan_rows: int = 10,
    scan_cols: int = 50,
    key_normalizer: Callable[[Any], str] = lambda x: str(x).strip(),
    skip_incomplete_keys: bool = True,
    close_wb: bool = True
) -> MappedData:
    """
    *`@returns`* `MappedData` `{ rows: list[list[Any]]; columns: list[str]; header: HeaderIndex }`
    - where `rows[i]` = `{ *RowKey, *(values from val_cols), original_row_index }`
    - `→ len(rows[i])` = `len(options.key_cols) + len(options.val_cols + selector cols) + 1`
    """
    
    wb: PyxlWorkbook | None = None
    try:
        resolved_sheetname: str | None = None
        wb = resolve_pyxl_wb(options.path, read_only=True, data_only=data_only)
        _sheetnames = wb.sheetnames
        if (options.sheetname):
            resolved_sheetname = options.sheetname
        elif (options.sheetname_selector_options):
            resolved_sheetname = Array.find(_sheetnames, options.sheetname_selector, strict=True)
        elif (len(wb.sheetnames) == 1):
            resolved_sheetname = _sheetnames[0]
        if (not resolved_sheetname):
            raise RuntimeError(LogMsg(
                f"Unable to resolve sheetname for workbook",
                candidates=_sheetnames,
                sheetname=options.sheetname,
                sheetname_selector_options=annotate(options.sheetname_selector_options)
            ))
        ws = wb[resolved_sheetname]
        header = find_header(
            ws, 
            required_headers=options.columns,
            scan_rows=scan_rows, 
            scan_cols=scan_cols, 
            col_offset=options.col_offset, 
            close_wb=False
        )
        key_indices = OrderedDict({
            key: header.columns[key]
            for key in options.key_cols
        })
        _val_cols = [*options.val_cols]
        if (options.column_selector_options is not None):
            _val_cols.extend(Array.filter(header.headers, options.column_selector))
        val_indices = OrderedDict({
            key: header.columns[key]
            for key in _val_cols
        })
        columns = [
            *options.key_cols,
            *_val_cols,
            ROW_INDEX_COLUMN
        ]
        rows: list[list[Any]] = []
        keys_seen: set[tuple[str,...]] = set()
        data_start_row = header.row + 1
        for row_idx, r in enumerate(
            ws.iter_rows(data_start_row, ws.max_row+1, values_only=True), 
            start=data_start_row
        ):
            key = tuple(key_normalizer(r[col_idx-1]) for col_idx in key_indices.values())
            if _key_invalid(key, skip_incomplete_keys): 
                continue
            if (key in keys_seen):
                raise RuntimeError(LogMsg(
                    f"Encountered duplicate '{key=} @ {row_idx=}"
                ))
            values = tuple(r[col_idx-1] for col_idx in val_indices.values())
            rows.append([*key, *values, row_idx])
        return MappedData(rows, columns, header)
    except Exception as e:
        raise RuntimeError(LogMsg.err(
            e, 
            f"Failed to read data from worksheet",
            options=str(options)
        )) from e
    finally:
        if (wb is not None and close_wb): 
            wb.close()

def _key_invalid(
    key: tuple[str, ...],
    /,
    blank_is_invalid: bool=True
):
    return (
        not all(key) 
        or (
            blank_is_invalid 
            and any(part == "None" for part in key) # part == "" or 
        )
    )

__all__ = export(read_mapped_data)