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
        query.remove_first_rows(4)
        query.fillna(columns=list(range(3)), method='ffill')
        query.fillna(columns=list(range(4, 26)), method='bfill')
        query.promote_headers()
        query.df.rename(columns={4: 'ID'}, inplace=True)
        query.df.query('`A/A` != "A/A"', inplace=True)
        print(query.df.info())

main()
