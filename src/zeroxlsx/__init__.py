"""zeroxlsx — write .xlsx files with zero dependencies.

>>> from zeroxlsx import Workbook, Style
>>> book = Workbook()
>>> sheet = book.add_sheet("Report")
>>> sheet.append(["item", "qty"], style=Style(bold=True, fill="#DDDDDD"))
>>> sheet.append(["widget", 12])
>>> book.save("report.xlsx")
"""

from ._styles import Style
from .workbook import Sheet, Workbook

__all__ = ["Workbook", "Sheet", "Style", "__version__"]
__version__ = "0.1.0"
