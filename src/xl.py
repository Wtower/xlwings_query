"""
Defines classes to encapsulate xlwings
"""
from __future__ import annotations
from pathlib import Path
import xlwings as xw

class App():
    """
    Encapsulate the App class that corresponds to an Excel instance
    Check if app is not open, then open it
    https://docs.xlwings.org/en/latest/api.html#app
    """
    def __init__(self) -> None:
        self._app: xw.App = xw.App(visible=True) if not xw.apps else xw.apps.active

class Book():
    """
    Encapsulate the Book class
    https://docs.xlwings.org/en/latest/api.html#book
    """
    def __init__(self, filename: str) -> None:
        """
        Check that a book is open or open it
        https://github.com/Wtower/xlwings_query/issues/3
        https://stackoverflow.com/questions/2491819/how-to-return-a-value-from-init-in-python
        https://stackoverflow.com/q/33533148/940098
        """
        book: xw.Book = next((book for book in xw.books if book.name == Path(filename).name), None)
        self._book = xw.books.open(filename) if book is None else book

    def __getattr__(self, __name: str):
        """
        Return the encapsulated object properties
        https://stackoverflow.com/a/14182553/940098
        https://stackoverflow.com/a/3464154/940098 (not working)
        """
        return getattr(self._book, __name)

    def get_sheet(self: Book, sheet_name: str) -> Sheet:
        """
        Get a sheet
        """
        return Sheet(self._book.sheets[sheet_name])

    def get_or_create_sheet(self: Book, sheet_name: str) -> Sheet:
        """
        Get an existing sheet or create it
        """
        sheet: xw.Sheet = next((s for s in self._book.sheets if s.name == sheet_name), None)
        return Sheet(self._book.sheets.add(name=sheet_name) if sheet is None else sheet)

class Sheet():
    """
    Encapsulate the Sheet class
    https://docs.xlwings.org/en/latest/api.html#sheet
    """
    def __init__(self, sheet: xw.Sheet) -> None:
        self._sheet = sheet

    def __getattr__(self, __name: str):
        return getattr(self._sheet, __name)

    def __getitem__(self, items):
        return self._sheet[items]

    def get_or_create_table(self: Sheet, table_name: str):
        """
        Get an existing table or create it
        """
        table = next((table for table in self._sheet.tables if table.name == table_name), None)
        return self._sheet.tables.add(source=self._sheet['A1'], name=table_name) if table is None else table
