"""
*`@file`* `engines/xlcom/legacy/import_wizard`
"""
from pathlib import Path
import time
from typing import (Any, Literal)
from win32com.client import CDispatch
from pyqi import Array, Csv, String, fs
from pyqi.types import (Arbitrary, resolve_arbitrary)
from lute import (LogMsg, get_logger)
from noxl.constants import vba
from ..session import ExcelSession
from .types import (ExcelDelimiterAttributesDict, TextImportAttributesDict)

mlog = get_logger(Path(__file__).parent.parent.stem)

class ConnectionSource(String.BaseEnum):
    TEXT = 'TEXT'

def text_import(
    src: fs.StrPath,
    /,
    *outpath: fs.StrPath,
    delimiter: Literal["\t", ";", ",", " "] = "\t",
    qualifier: int = vba.xlTextQualifierNone,
    encoding: int = vba.xlUTF8,
    overwrite: bool=True
):
    _outpaths = Array.map(outpath, fs.to_path)
    written: list[Path] = []
    try:
        if (not overwrite):
            for p in _outpaths:
                if (not p.is_file()): continue
                raise FileExistsError(LogMsg(
                    f"File already exists at outpath {p}"
                ))
        src = fs.to_path(src)
        with ExcelSession() as xl:
            column_count = Csv.get_header_length(src, delimiter=delimiter)
            if (not column_count):
                raise RuntimeError(LogMsg(
                    f"Unable to determine positve number of coumns"
                ))
            wb = xl.excel.Workbooks.Add()
            ws = wb.Worksheets(1)
            column_types = tuple([vba.xlGeneralFormat] * column_count)
            qt = ws.QueryTables.Add(
                Connection=f"{ConnectionSource.TEXT};{str(src)}",
                Destination=ws.Range("A1")
                )
            attrs = get_text_import_attrs(
                column_types, 
                delimiter=delimiter, 
                text_qualifier=qualifier, 
                encoding=encoding
            )
            for k, v in attrs.items():
                safe_setattr(qt, k, v)
            qt.Refresh(False)
            while qt.Refreshing:
                time.sleep(0.1)
            qt.Delete()
            for p in _outpaths:
                fs.ensure_dir_exists(p.parent)
                if (not fs.dir_is_writable(p)): 
                    raise RuntimeError(LogMsg(
                        f"parent directory not writable for {p}"
                    ))
                p.unlink(missing_ok=True)
                wb.SaveAs(Filename=str(p), FileFormat=vba.xlOpenXMLWorkbook)
                written.append(p)
                time.sleep(1)
            wb.Close(SaveChanges=False)
        return written
    except Exception as e:
        raise RuntimeError(LogMsg.err(e,
            message=f"Failed to execute text import using legacy wizard",
            src=str(src),
            outpath=str(outpath),
            qualifer=qualifier,
            encoding=encoding,
            overwrite=overwrite,
        )) from e

def get_text_import_attrs(
    column_types: tuple[int, ...],
    field_names: bool = True,
    row_numbers: bool = False,
    refresh_style: int = vba.xlInsertDeleteCells,
    encoding: int = vba.xlUTF8,
    parse_type: int = vba.xlDelimited,
    text_qualifier: int = vba.xlTextQualifierNone,
    delimiter: Literal["\t", ";", ",", " "] = "\t",
) -> TextImportAttributesDict:
    return {
        "FieldNames": field_names,
        "RowNumbers": row_numbers,
        "FillAdjacentFormulas": False,
        "PreserveFormatting": True,
        "RefreshOnFileOpen": False,
        "RefreshStyle": refresh_style,
        "SavePassword": False,
        "SaveData": True,
        "AdjustColumnWidth": True,
        "RefreshPeriod": 0,
        "TextFilePromptOnRefresh": False,
        "TextFilePlatform": encoding,
        "TextFileStartRow": 1,
        "TextFileParseType": parse_type,
        "TextFileTextQualifier": text_qualifier,
        "TextFileTrailingMinusNumbers": True,
        "TextFileColumnDataTypes": column_types,
        **delimiter_attrs(delimiter)
    }

def delimiter_attrs(
    delimiter: Literal["\t", ";", ",", " "] = "\t",
    /,
) -> ExcelDelimiterAttributesDict:
    result = ExcelDelimiterAttributesDict(
        TextFileConsecutiveDelimiter=False,
        TextFileTabDelimiter=False,
        TextFileSemicolonDelimiter=False,
        TextFileCommaDelimiter=False,
        TextFileSpaceDelimiter=False
    )
    match delimiter:
        case "\t":
            result["TextFileTabDelimiter"] = True
        case ";":
            result["TextFileSemicolonDelimiter"] = True
        case ",":
            result["TextFileCommaDelimiter"] = True
        case " ":
            result["TextFileSpaceDelimiter"] = True
        case _:
            raise ValueError(f"Unsupported delimiter: {delimiter!r}")
    return result


def safe_setattr(obj: CDispatch, key: str, value: Any, /) -> None:
    try:
        setattr(obj, key, value)
    except Exception as e:
        raise RuntimeError(f"Failed setting Object.{key} = {value!r}") from e