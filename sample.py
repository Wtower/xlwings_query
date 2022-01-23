# pylint: disable=invalid-name, non-ascii-name
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
        query.remove_last_rows(2)
        query.fillna(columns=list(range(3)), method='ffill')
        query.fillna(columns=list(range(4, 26)), method='bfill')
        query.promote_headers()
        query.df.rename(columns={4: 'ID'}, inplace=True)
        query.df.query('`Column X` != "A" and `Column Y` == 1', inplace=True)
        query.split_text_column('Column Z', 'x|y')
        regx = r'(\d+)\s?(?:pcs|pieces|)\s{0,3}[xX\*]\s{0,3}(\d+[.,]\d+)[m]?'
        query.extract_text_column('Comment', regx, ['Pieces', 'Length'])
        query.replace_value_text_column('Length', ',', '.')
        query.df.dropna(subset=['Length'], inplace=True)
        query.drop_columns_idx(list(range(3, 27)))
        print(query.df.info())

main()
