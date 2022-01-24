xlwings_query
=============

Import, connect and transform data into Excel.

Description
-----------

The concept is to apply data transformations to a main query object.
When the data is ready, export it to an Excel table.
It is inspired by MS Power Query based in Python.
The target is to use the power of Pandas and overcome platform issues with Excel.

Methods
-------

See `sample.py`.

### xlwings_query.Query(filename: str, query_name: str)
Specify the target excel filename and the query name.

### source_excel_workbook(filename: Path)
Append an Excel workbook to the query.

### navigate(self, sheet_name: str, table_name: str = None)
Navigate to the selected item (sheet/table) and append to the query

### remove_first_rows(rows: int)
Remove first rows from table

### remove_last_rows(rows: int)
Remove last rows from table

### promote_headers()
Promotes the first row of values as the new column headers.

### fillna(method: str, columns: list[str] = None)
The value of the previous or next cell is propagated to the null-value cells

### split_text_column(column: str, pat: str = None, columns: list[str] = None)
Split a column around a given delimiter or regex

### extract_text_column(column: str, pat: str, columns: list[str])
Extract regex capture groups as columns

### replace_value_text_column(column: str, pat: str, repl: str)
Replace each occurence of pattern in the column

### drop_columns_idx(idx: list[int])
Remove columns by index
