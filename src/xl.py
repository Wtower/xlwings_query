"""
Defines classes to encapsulate xlwings
"""
from __future__ import annotations

import unicodedata
from pathlib import Path
from typing import Optional, Union

import pandas as pd
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
    def __init__(self, filename: str) -> None:
        """
        Check that a book is open or open it.
        Normalizes filename for cross-platform support (#6).
        Initializes Excel if not open.
        ## Parameters
        filename: str
            The pathfilename to open.
        ## Links
        https://github.com/Wtower/xlwings_query/issues/3
        https://stackoverflow.com/q/33533148/940098
        """
        App()
        filename = unicodedata.normalize('NFC', filename)
        if Path(filename).name in [b.name for b in xw.books]:
            self._book: xw.Book = xw.books[Path(filename).name]
        else:
            self._book = xw.books.open(filename)

    @staticmethod
    def read( # pylint: disable=too-many-arguments
        filename: str,
        sheet_name: Optional[Union[int, str]] = 0,
        table_name: Optional[Union[int, str]] = None,
        index_col: Optional[int] = None,
        header: Optional[int] = 0
    ) -> pd.DataFrame:
        """
        Read an Excel file into Pandas Dataframe.
        If the file is open, use xlwings, otherwise pandas.
        ## Parameters
        filename: str
            The pathfilename to get or open.
        sheet_name: int or str, default 0
            The sheet index or name.
        table_name: int or str, default 0
            The table index or name in sheet (only for xlwings).
            If not provided, then `used_range` is used.
        index_col: int, default None
            Column (0-indexed) to use as row labels (as with Pandas).
            Note: For xlwings, None is 0, so converted to 1-indexed index.
        header: int, default 0
            Row (0-indexed) to use as column labels.
            Note: For xlwings, None is 0, so converted to 1-indexed index.
            Not used when `table_name` is specified.
        """
        filename = unicodedata.normalize('NFC', filename)
        if xw.apps and Path(filename).name in [b.name for b in xw.books]:
            if index_col is None:
                index_col = -1
            if header is None:
                header = -1
            if table_name is None:
                return xw.books[Path(filename).name].sheets[sheet_name].used_range \
                    .options(pd.DataFrame, index=index_col + 1, header=header + 1).value
            return xw.books[Path(filename).name].sheets[sheet_name].tables[table_name].range \
                .options(pd.DataFrame, index=index_col + 1).value
        return pd.read_excel(filename,
            sheet_name=sheet_name, header=header, index_col=index_col)

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
    def __init__(self, table=None) -> None:
        #, file_name=None, sheet_name=None, table_name=None) -> None:
        self._table = table
