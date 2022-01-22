"""
Transform data.
Name this file with the same filename as the target xlsx in same folder.
"""
import xlwings_query as xwq

def main():
    """
    Main function
    """
    with xwq.Query(__file__, 'Target sheet') as query:
        query.source_excel_workbook('My source workbook.xlsx')
        query.navigate('Sheet1')
        # ...

main()
