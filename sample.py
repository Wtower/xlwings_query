"""
Transform data.
Name this file with the same filename as the target xlsx in same folder.
"""
import xlwings_query as xwq

def main():
    """
    Main function
    """
    query = xwq.Query(__file__, 'Target sheet')
    query.origin_excel_workbook('My source workbook.xlsx')
    # ...

main()
