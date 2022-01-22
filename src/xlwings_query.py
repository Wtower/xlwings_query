"""
Defines the main class
"""
from email import header
from operator import index
from pathlib import Path
from statistics import mode
import xlwings as xw
import pandas as pd

class Query:
    """
    Import, connect and transform Excel data
    """
    def __init__(self, filename: str, query_name: str) -> None:
        # Replace extension with .xlsx if __file__ has been provided
        self.filename = Path(filename).with_suffix('.xlsx')

        # The query name to be used as the xported table and sheet name
        self.query_name = query_name

        # The current source object
        self.source = None

        # The query data to be exported
        self.df = None

        # Check if xl app is not open, then open it
        self.app = xw.App(visible=True) if not xw.apps else None

        # The target object
        self.book = self.__get_excel_workbook(self.filename)
        # print(self.book.sheets(1).range('A1').value)

    def __enter__(self) -> None:
        """
        Context manager enter
        """
        return self

    def __exit__(self, *exc) -> None:
        """
        Context manager exit: save transformed data
        """
        sheet = next((sheet for sheet in self.book.sheets if sheet.name == self.query_name), None)
        sheet = self.book.sheets.add(name=self.query_name) if sheet is None else sheet
        table_name = 'tbl' + self.query_name
        table = next((table for table in sheet.tables if table.name == table_name), None)
        table = sheet.tables.add(source=sheet['A1'], name=table_name) if table is None else table
        table.update(self.df, index=False)

    @staticmethod
    def __get_excel_workbook(filename: str) -> xw.Book:
        """
        Check that a book is open or open it
        https://github.com/Wtower/xlwings_query/issues/3
        """
        book = next((book for book in xw.books if book.name == Path(filename).name), None)
        return xw.books.open(filename) if book is None else book

    def source_excel_workbook(self, filename: str) -> None:
        """
        Append an Excel workbook to the query
        """
        self.source = self.__get_excel_workbook(Path(filename).with_suffix('.xlsx'))
        # Get a list of book sheets. Tables in xlwings belong in xw.sheet, not book.
        data = [(sheet.name, 'Sheet') for sheet in self.source.sheets]
        self.df = pd.DataFrame(data, columns=('Name', 'Kind'))

    def navigate(self, sheet_name: str, table_name=None) -> None:
        """
        Navigate to the selected item (sheet/table) and append to the query
        https://gist.github.com/Elijas/2430813d3ad71aebcc0c83dd1f130e33
        """
        sheet = self.source.sheets[sheet_name]
        if table_name is None:
            size = (sheet.api.UsedRange.Rows.Count, sheet.api.UsedRange.Columns.Count)
            range = sheet.range((1, 1), size)
        else:
            range = sheet.tables[table_name].data_body_range
        self.df = range.options(pd.DataFrame, index=False, header=False).value
