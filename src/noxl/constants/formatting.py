"""
*`@file`* `constants/formatting`
"""
from pyqi import String


class CellFormat(String.BaseEnum):
    ACCOUNTING = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
    GENERAL = "General"
    INTEGER = "0"
    TWO_DECIMALS = "0.00"
    COMMA_SEP_TWO_DECIMALS = "#,##0.00"
    COMMA_SEP_INTEGER = "#,##0"
    FLOAT = "0.###############"
    """ up to 15 decimal places"""
    MDY = "m/d/yyyy"

ACCOUNTING_FORMAT_PROPERTIES = {
    'num_format': CellFormat.ACCOUNTING.value
}

DEFAULT_HEADER_FORMAT = {
    "bold": True,
    "align": "center",
    "bottom": 1,
    "text_wrap": True,
}