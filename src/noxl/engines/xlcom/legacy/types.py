"""
*`@file`* `engines/xlcom/legacy/types.py`
"""




from typing import TypedDict


class TextImportAttributesDict(TypedDict):
    FieldNames: bool
    RowNumbers: bool
    FillAdjacentFormulas: bool
    PreserveFormatting: bool
    RefreshOnFileOpen: bool
    RefreshStyle: int
    SavePassword: bool
    SaveData: bool
    AdjustColumnWidth: bool
    RefreshPeriod: int
    TextFilePromptOnRefresh: bool
    TextFilePlatform: int
    TextFileStartRow: int
    TextFileParseType: int
    TextFileTextQualifier: int
    TextFileTrailingMinusNumbers: bool
    TextFileColumnDataTypes: tuple[int, ...]
    # ======================================
    # ExcelDelimiterAttributesDict
    # ======================================
    TextFileConsecutiveDelimiter: bool
    TextFileTabDelimiter: bool
    TextFileSemicolonDelimiter: bool
    TextFileCommaDelimiter: bool
    TextFileSpaceDelimiter: bool

class ExcelDelimiterAttributesDict(TypedDict):
    TextFileConsecutiveDelimiter: bool
    TextFileTabDelimiter: bool
    TextFileSemicolonDelimiter: bool
    TextFileCommaDelimiter: bool
    TextFileSpaceDelimiter: bool
