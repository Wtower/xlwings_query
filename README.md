xlwings_query
=============

Import, connect and transform data into Excel.

Description
-----------

The concept is to apply data transformations to a main query object.
When the data is ready, export it to an Excel table.
It is inspired by MS Power Query based in Python.
The target is to use the power of Pandas and overcome platform issues with Excel.

Sample usage
------------

Specify the same filename as the xlsx in same folder.

```
"""
Transform data
"""
import xlwings_query as xwq

def main():
    """
    Main function
    """
    q = xwq.Query(__file__, 'Target sheet')
    q.excel_workbook('My source workbook.xlsx')
    # ...

main()
```

Methods
-------

### Query initialisation

Specify the target excel filename and the query name.

### excel_workbook

Append an Excel workbook to the query.
