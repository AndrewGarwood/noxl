"""
*`@file`* `constants/vba`
- [VBA constants](https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/using-constants)
- [constant-enumerations-for-Excel](https://www.experts-exchange.com/questions/27613392/List-all-constant-enumerations-for-Excel.html)
"""

from typing import Final

xlUTF8: Final[int] = 65001
xlInsertDeleteCells: Final[int] = 1
xlDelimited: Final[int] = 1
xlTextQualifierNone: Final[int] = -4142
xlGeneralFormat: Final[int] = 1
xlOpenXMLWorkbook: Final[int] = 51
"""  `.xlsx` """
xlDatabase: Final[int] = 1


xlRowField: Final[int] = 1
xlColumnField: Final[int] = 2
xlPageField: Final[int] = 3
""" Filter field """
xlDataField: Final[int] = 4

xlSum: Final[int] = -4157
xlCount: Final[int] = -4112
xlCountNums: Final[int] = -4113
xlAverage: Final[int] = -4106
xlMax: Final[int] = -4136
xlMin: Final[int] = -4139
xlProduct: Final[int] = -4149
xlStDev: Final[int] = -4155
xlStDevP: Final[int] = -4156
xlVar: Final[int] = -4164
xlVarP: Final[int] = -4165
xlUnknown: Final[int] = 1000

xlAscending: Final[int] = 1
xlDescending: Final[int] = 2

xlDone: Final[int] = 0
xlCalculating: Final[int] = 1
xlPending: Final[int] = 2


xlRepeatLabels: Final[int] = 2
xlCompactRow: Final[int] = 0
xlMissingItemsDefault: Final[int] = -1

xlSlicer: Final[int] = 1
xlTimeline: Final[int] = 2

xlTabular: Final[int] = 0
xlOutline: Final[int] = 1

