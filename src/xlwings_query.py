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

        # Check book is open
        # https://github.com/Wtower/xlwings_query/issues/3
        self.book = next((book for book in xw.books if book.name == Path(self.filename).name), None)
        if self.book is None:
            self.book = xw.books.open(self.filename)
        print(self.book.sheets(1).range('A1').value)
