xlwings_query
=============

Import, connect and transform data into Excel.

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
    with xwq.Query(__file__) as q:
        pass

main()
```
