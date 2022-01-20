"""
Defines the main class
"""
from pathlib import Path
import xlwings as xw

class Query:
    """
    Import, connect and transform Excel data
    """
    def __init__(self, filename: str, query_name: str) -> None:
        # Replace extension with .xlsx if __file__ has been provided
        self.filename = Path(filename).with_suffix('.xlsx')

        # The query name to be used as the xported table and sheet name
        self.query_name = query_name

        # The query object to be exported
        self.query = None

        # Check if xl app is not open, then open it
        if not xw.apps:
            self.app = xw.App(visible=True)

        # The target object
        self.book = self.__get_excel_workbook(self.filename)
        # print(self.book.sheets(1).range('A1').value)
        # Eventually, a context manager will be needed to cleanup and export

    @staticmethod
    def __get_excel_workbook(filename: str) -> xw.Book:
        """
        Check that a book is open or open it
        https://github.com/Wtower/xlwings_query/issues/3
        """
        book = next((book for book in xw.books if book.name == Path(filename).name), None)
        if book is None:
            book = xw.books.open(filename)
        return book

    def origin_excel_workbook(self, filename: str) -> None:
        """
        Append an Excel workbook to the query
        """
        self.query = self.__get_excel_workbook(Path(filename).with_suffix('.xlsx'))
