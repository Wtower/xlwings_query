from pathlib import Path
import xlwings as xw

class Query:
    def __init__(self, filename) -> None:
        # Replace extension with .xlsx if __file__ has been provided
        self.filename = Path(filename).with_suffix('.xlsx')
        self.app = None

        # Check if xl is not open, then open it
        if not xw.apps:
            self.app = xw.App(visible=False)
        
        # Get or open book
        self.book = xw.books.open(self.filename)
        print(self.book.sheets(1).range('A1').value)

    def __enter__(self) -> None:
        # If we opened excel, provide context manager
        if self.app is not None:
            self.app.__enter__()
        return self

    def __exit__(self, type, value, traceback) -> None:
        # If we opened excel, then close it in context manager
        if self.app is not None:
            self.app.__exit__(type, value, traceback)