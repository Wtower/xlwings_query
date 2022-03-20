"""
Defines classes to encapsulate xlwings
"""
from __future__ import annotations
import os
from pathlib import Path
import xlwings as xw

class App(): # pylint: disable=too-few-public-methods
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
    def __init__(self, filename: str, fuzzy: bool = False) -> None:
        """
        Check that a book is open or open it.
        Initializes Excel if not open.
        ## Parameters
        filename: str
            The pathfilename to open.
        fuzzy: bool, default False
            If defined, match the closest filename.
            Useful for cross-platform execute where extended charset filenames may slightly differ.
        ## Links
        https://github.com/Wtower/xlwings_query/issues/3
        https://stackoverflow.com/q/33533148/940098
        """
        App()
        if fuzzy:
            from thefuzz import process # pylint: disable=import-outside-toplevel
            filename = str(Path(
                Path(filename).parent,
                process.extractOne(Path(filename).name, os.listdir())[0]))
        if Path(filename).name in [b.name for b in xw.books]:
            self._book: xw.Book = xw.books[Path(filename).name]
        else:
            self._book = xw.books.open(filename)

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
        if sheet_name in [s.name for s in self._book.sheets]:
            return Sheet(self._book.sheets[sheet_name])
        return Sheet(self._book.sheets.add(name=sheet_name))

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
        if table_name in [table for table in self._sheet.tables]:
            return self._sheet.tables[table_name]
        return self._sheet.tables.add(source=self._sheet['A1'], name=table_name)

class Table(): # pylint: disable=too-few-public-methods
    """
    Encapsulate the Table class
    https://docs.xlwings.org/en/latest/api.html#table
    """
    def __init__(self, table=None): #, file_name=None, sheet_name=None, table_name=None) -> None:
        self._table = table
