"""
Defines the main class
"""
from operator import index
from pathlib import Path
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

        # The query data to be exported
        self.df = None

        # Check if xl app is not open, then open it
        if not xw.apps:
            self.app = xw.App(visible=True)

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
        # table.update(self.df[['Name', 'Data', 'Kind']], index=False)
        # print(type(self.df.iloc[0]['Data']))

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
        source = self.__get_excel_workbook(Path(filename).with_suffix('.xlsx'))
        # Get a list of book sheets. Tables in xlwings belong in xw.sheet, not book.
        data = [(sheet.name, sheet, 'Sheet') for sheet in source.sheets]
        self.df = pd.DataFrame(data, columns=('Name', 'Data', 'Kind'))

    def navigate(self, item: str) -> None:
        """
        Navigate to the selected item (sheet/table) and append to the query
        """
        # if isinstance(self.data, xw.Book):
        #    print('chk')
        # TODO: handle unexpected data
