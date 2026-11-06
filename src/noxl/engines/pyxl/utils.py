"""
*`@file`* `engines/pyxl/utils`
"""

from pathlib import Path
from typing import Any, Callable, Literal, OrderedDict, Sequence, overload
from pyqi import (Array, Function, require)
from lute import annotate, LogMsg
from openpyxl import load_workbook

from openpyxl.worksheet._read_only import (ReadOnlyWorksheet)
from openpyxl.worksheet.worksheet import (Worksheet)
from noxl.models import (HeaderIndex, normalize_cell_value)
from .types import (PyxlWorkbook, PyxlWorkbookSource, PyxlWorksheet, PyxlSheetInfo)

#TODO have resolve_pyxl_wb return tuple where second val is bool indicating 
# whether module owns the wb or not, then only close workbook in `finally` 
# statements when wb is owned...

def resolve_pyxl_wb(
    wb_source: PyxlWorkbookSource,
    /, 
    read_only: bool = False, 
    data_only: bool = True
) -> PyxlWorkbook:
    return (
        wb_source if isinstance(wb_source, PyxlWorkbook)
        else load_workbook(wb_source, read_only=read_only, data_only=data_only)
    )

def is_hidden_worksheet(ws: PyxlWorksheet, /) -> bool:
    return "hidden" in ws.sheet_state.lower()

def get_visible_sheetnames(wb_source: PyxlWorkbookSource, /) -> list[str]:
    wb = resolve_pyxl_wb(wb_source)
    return Array.filter(wb.sheetnames, lambda k: not is_hidden_worksheet(wb[k]))

def get_sheetinfo(wb_source: PyxlWorkbookSource, /) -> dict[str, PyxlSheetInfo]:
    wb = resolve_pyxl_wb(wb_source, True, True)
    result: dict[str, PyxlSheetInfo] = dict()
    for i, s in enumerate(wb.sheetnames):
        result[s] = PyxlSheetInfo(wb[s].sheet_state, i)
    if not isinstance(wb_source, PyxlWorkbook):
        wb.close()
    return result


def find_sheetname(
    wb_source: PyxlWorkbookSource,
    filter_fn: Callable[[str], bool],
    /,
    *,
    close_wb: bool=True
) -> str:
    wb: PyxlWorkbook | None = None
    try:
        wb = resolve_pyxl_wb(wb_source)
        valid_sheetnames = wb.sheetnames
        return Array.find(valid_sheetnames, filter_fn, strict=True)
    except Exception as e:
        raise ValueError(LogMsg.err(e,
            f"Unable to find sheetname matching filter criteria",
            wb_source=annotate(wb_source),
            filter_fn=annotate(filter_fn)
        )) from e
    finally:
        if wb is not None and close_wb: wb.close()


@overload
def find_header(
    src: PyxlWorkbookSource,
    *,
    sheetname: str | Callable[[str], bool],
    required_headers: Sequence[str] | None=None,
    min_header_len: int = 2,
    scan_rows: int=25,
    scan_cols: int=50,
    unique_sheetname: bool=False,
    normalize_header_value: Callable[..., str] | None=None,
    col_offset: int | None=None,
    close_wb: bool=True
) -> HeaderIndex: ...

@overload
def find_header(
    src: PyxlWorksheet,
    *,
    required_headers: Sequence[str] | None=None,
    min_header_len: int = 2,
    scan_rows: int=25,
    scan_cols: int=50,
    unique_sheetname: bool=False,
    normalize_header_value: Callable[..., str] | None=None,
    col_offset: int | None=None,
    close_wb: Literal[False]=False
) -> HeaderIndex: ...


def find_header(
    src: PyxlWorkbookSource | PyxlWorksheet,
    *,
    sheetname: str | Callable[[str], bool] | None=None,
    required_headers: Sequence[str] | None=None,
    min_header_len: int = 2,
    scan_rows: int=25,
    scan_cols: int=50,
    unique_sheetname: bool=False,
    normalize_header_value: Callable[..., str] | None=None,
    col_offset: int | None=None,
    close_wb: bool=True
) -> HeaderIndex:
    """
    *`@returns`* `HeaderIndex (dataclass) = { row: int; columns: dict[str, int]; headers: tuple[str, ...] }`
    """
    wb: PyxlWorkbook | None = None
    resolved_sheetname: str | None = None
    try:
        require.positive_integer(scan_cols, "scan_cols")
        require.positive_integer(scan_rows, "scan_rows")
        if (not required_headers and min_header_len < 1):
            raise ValueError(
                "min_header_len must be positive integer when required_headers not provided"
            )
        if (col_offset is not None and (not isinstance(col_offset, int) or col_offset < 0)):
            raise ValueError(LogMsg.bad_arg(
                col_offset, "col_offset", "Nonnegative integer required", int
            ))
        if (col_offset is None): col_offset = 0
        resolved_worksheet: PyxlWorksheet | None = None
        if (isinstance(src, (ReadOnlyWorksheet, Worksheet))):
            resolved_worksheet = src
            resolved_sheetname = src.title
        elif (
            isinstance(src, (PyxlWorkbook, str, Path)) 
            and sheetname is not None
        ):
            wb = resolve_pyxl_wb(src)
            valid_sheetnames = wb.sheetnames
            if (isinstance(sheetname, str)):
                if (sheetname not in valid_sheetnames):
                    raise KeyError(LogMsg(
                        f"sheetname not found in workbook",
                        workbook=wb.path,
                        sheetname=sheetname,
                        valid_sheetnames=valid_sheetnames
                    ))
                resolved_sheetname = sheetname
            elif isinstance(sheetname, Callable):
                resolved_sheetname = Array.find(
                    valid_sheetnames, 
                    sheetname, 
                    unique=unique_sheetname, 
                    strict=True
                )
            else:
                resolved_sheetname = get_visible_sheetnames(wb)[0]
            resolved_worksheet = wb[resolved_sheetname]
        if (resolved_worksheet is None):
            raise RuntimeError(LogMsg(f"Unable to resolve sheetname"))
        ws = resolved_worksheet
        max_scan_row = min(scan_rows, ws.max_row or 1)
        max_scan_col = min(scan_cols, ws.max_column or 1)
        _rows = list(ws.iter_rows(values_only=True))
        def acc_rowdict(acc: OrderedDict[str, int], curr: Any, i: int):
            if (not curr): return acc
            value = str(curr).strip()
            if (not value): return acc
            acc[value] = i
            return acc
        if (normalize_header_value is None):
            normalize_header_value = Function.Partial(
                normalize_cell_value, 
                case_sensitive=False
            )
        header_found: bool = False
        for row_idx, r in enumerate(_rows[:max_scan_row], start=1):
            if row_idx > scan_rows:
                break
            row = r[col_offset:max_scan_col]
            # print(f"@ {row_idx=}, {r=}")
            rowdict = Array.reduce(row, acc_rowdict, OrderedDict(), 1+col_offset)
            norm_row = set(Array.map(rowdict.keys(), normalize_header_value))
            if (required_headers):
                header_found = Array.every(
                    Array.map(required_headers, normalize_header_value),
                    lambda h: h in norm_row
                )
            else:
                header_found = len(norm_row) >= min_header_len
            
            if (header_found):
                return HeaderIndex(row_idx, columns=rowdict)
        raise RuntimeError(LogMsg(
            f"Unable to find header in target worksheet",
            max_scan_row=max_scan_row,
            max_scan_col=max_scan_col
        ))
    except Exception as e:
        raise RuntimeError(LogMsg.err(
            e,
            scan_rows=scan_rows,
            scan_cols=scan_cols,
            wb_source=annotate(src),
            resolved_sheetname=resolved_sheetname,
            required_headers=required_headers,
            min_header_len=min_header_len
        )) from e
    finally:
        if wb is not None and close_wb: wb.close()

