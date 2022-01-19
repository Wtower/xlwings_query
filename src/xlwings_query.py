"""
Defines the main class
"""
from pathlib import Path
import xlwings as xw

class Query:
    """
    Import, connect and transform Excel data
    """
    def __init__(self, filename) -> None:
        # Replace extension with .xlsx if __file__ has been provided
        self.filename = Path(filename).with_suffix('.xlsx')

        # Check if xl is not open, then open it
        if not xw.apps:
            self.app = xw.App(visible=True)

        # Get or open book
        self.book = xw.books.open(self.filename)
        print(self.book.sheets(1).range('A1').value)
