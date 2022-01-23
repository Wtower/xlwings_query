"""
Defines the main class
"""
from pathlib import Path
import xlwings as xw
import pandas as pd

class Query:
    """
    Import, connect and transform Excel data
    """
    def __init__(self, filename: str, query_name: str) -> None:
        # Replace extension with .xlsx if __file__ has been provided
        self.filename: Path = Path(filename).with_suffix('.xlsx')

        # The query name to be used as the xported table and sheet name
        self.query_name: str = query_name

        # The current source object
        self.source: xw.Book = None

        # The query data to be exported
        self.df: pd.DataFrame = None # pylint: disable=invalid-name

        # Check if xl app is not open, then open it
        self.app: xw.App = xw.App(visible=True) if not xw.apps else None

        # The target object
        self.book: xw.Book = self.__get_excel_workbook(self.filename)

    def __enter__(self):
        """
        Context manager enter
        """
        return self

    def __exit__(self, *exc) -> None:
        """
        Context manager exit: save transformed data
        """
        sheet: xw.Sheet = next((s for s in self.book.sheets if s.name == self.query_name), None)
        sheet = self.book.sheets.add(name=self.query_name) if sheet is None else sheet
        table_name: str = 'tbl' + self.query_name
        table = next((table for table in sheet.tables if table.name == table_name), None)
        table = sheet.tables.add(source=sheet['A1'], name=table_name) if table is None else table
        #TODO: merge existing columns in target not present in source (xlw replaces them now)
        table.update(self.df, index=False)

    @staticmethod
    def __get_excel_workbook(filename: Path) -> xw.Book:
        """
        Check that a book is open or open it
        https://github.com/Wtower/xlwings_query/issues/3
        """
        book: xw.Book = next((book for book in xw.books if book.name == Path(filename).name), None)
        return xw.books.open(filename) if book is None else book

    def source_excel_workbook(self, filename: str) -> None:
        """
        Append an Excel workbook to the query
        """
        self.source = self.__get_excel_workbook(Path(filename).with_suffix('.xlsx'))
        # Get a list of book sheets. Tables in xlwings belong in xw.sheet, not book.
        data: list[tuple[str, str]] = [(sheet.name, 'Sheet') for sheet in self.source.sheets]
        self.df = pd.DataFrame(data, columns=('Name', 'Kind'))

    def navigate(self, sheet_name: str, table_name: str = None) -> None:
        """
        Navigate to the selected item (sheet/table) and append to the query
        https://gist.github.com/Elijas/2430813d3ad71aebcc0c83dd1f130e33
        """
        sheet: xw.Sheet = self.source.sheets[sheet_name]
        if table_name is None:
            size: tuple[int, int] = (
                sheet.api.UsedRange.Rows.Count,
                sheet.api.UsedRange.Columns.Count
            )
            xl_range: xw.Range = sheet.range((1, 1), size)
        else:
            xl_range = sheet.tables[table_name].data_body_range
        self.df = xl_range.options(pd.DataFrame, index=False, header=False).value

    def remove_first_rows(self, rows: int) -> None:
        """
        Remove first rows from table
        """
        self.df = self.df.iloc[rows:]

    def remove_last_rows(self, rows: int) -> None:
        """
        Remove last rows from table
        """
        self.df = self.df.iloc[:-rows]

    def promote_headers(self) -> None:
        """
        Promotes the first row of values as the new column headers.
        """
        self.df.columns = [name if name else i for i, name in enumerate(self.df.iloc[0])]
        self.remove_first_rows(1)

    def fillna(self, method: str, columns: list[str] = None) -> None:
        """
        The value of the previous or next cell is propagated to the null-value cells
        """
        columns = columns if columns else self.df.columns
        self.df[columns] = self.df[columns].fillna(method=method)

    def split_text_column(self, column: str, pat: str = None, columns: list[str] = None) -> None:
        """
        Split a column around a given delimiter or regex
        """
        columns = columns if columns else [column + '.1', column + '.2']
        self.df[columns] = self.df[column].str.split(pat, len(columns) - 1, expand=True)

    def extract_text_column(self, column: str, pat: str, columns: list[str]) -> None:
        """
        Extract regex capture groups as columns
        """
        self.df[columns] = self.df[column].str.extract(pat, expand=True)

    def replace_value_text_column(self, column: str, pat: str, repl: str) -> None:
        """
        Replace each occurence of pattern in the column
        """
        self.df[column] = self.df[column].str.replace(pat, repl)

    def drop_columns_idx(self, idx: list[int]) -> None:
        """
        Remove columns by index
        """
        self.df.drop(self.df.columns[idx], axis=1, inplace=True)
