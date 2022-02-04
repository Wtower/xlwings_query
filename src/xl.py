"""
Defines inherited classes from xlwings
"""
from __future__ import annotations
from pathlib import Path
import xlwings as xw

class App(xw.App):
    """
    Override the App class that corresponds to an Excel instance
    Check if xl app is not open, then open it
    https://docs.xlwings.org/en/latest/api.html#app
    """
    def __init__(self, visible: bool=None, spec=None, add_book=True, impl=None) -> None:
        if not xw.apps:
            visible=True if visible is None else visible
            super().__init__(visible, spec, add_book, impl)

class Book(xw.Book):
    """
    Override the Book class
    https://docs.xlwings.org/en/latest/api.html#book
    """
    def __new__(cls: type[Book], filename: str) -> Book:
        """
        Check that a book is open or open it
        https://github.com/Wtower/xlwings_query/issues/3
        https://stackoverflow.com/questions/2491819/how-to-return-a-value-from-init-in-python
        https://stackoverflow.com/questions/33533148/how-do-i-type-hint-a-method-with-the-type-of-the-enclosing-class
        """
        book: xw.Book = next((book for book in xw.books if book.name == Path(filename).name), None)
        book = xw.books.open(filename) if book is None else book
        # book.__class__ = Book
        return book

    def get_or_create_sheet(self: Book, sheet_name: str) -> Sheet:
        """
        Get an existing sheet or create it
        """
        sheet: xw.Sheet = next((s for s in self.sheets if s.name == sheet_name), None)
        return self.sheets.add(name=sheet_name) if sheet is None else sheet

class Sheet(xw.Sheet):
    """
    Override the Sheet class
    https://docs.xlwings.org/en/latest/api.html#sheet
    """
