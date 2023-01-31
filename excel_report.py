import pandas as pd

def merged_indices(df,col):
    '''
    df: a single level indexed pd.DataFrame
    col: a list of columns of interest
    merged_row_indices_df: a pd.DataFrame with 
        col (unique grouping of values from cols), 
        first (start row/col index), 
        last (end row/col index),
        idx (the index of the col/row index)
    '''
    merged_indices_df = (
        df
        .assign(index=range(0,len(df))) #creates an index column
        .groupby(col)['index']
        .agg(['first', 'last'])
        .assign(idx=list(df.columns).index(col[-1]))
        .reset_index() # moves row values from index to columns
    ) 
    return merged_indices_df 

def single_level_df_to_excel(df, filename = 'pandas_simple.xlsx', sort_df = True):
    '''
    This function takes a single level column and and single level index pd.DataFrame 
    and merges each column's row values based on the previous column's groupings.

    '''
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    sheet_name = 'Sheet1'
    row_padding = df.columns.nlevels + 1 if df.columns.nlevels >1 else 1 # accounts for column headers
    df.to_excel(writer, sheet_name = sheet_name, index=False, header=False)

    # Create an new Excel file and add a worksheet.
    workbook  = writer.book
    worksheet = writer.sheets[sheet_name]

    header_format = workbook.add_format({
        'bg_color': '#4F81BD',
        'bold': True,
        'font_color': '#FFFFFF',
        'border_color': '#FFFFFF',
        'border': 2
    })
    row_format1 = workbook.add_format({
        'bg_color': '#B8CCE4',
        'border_color': '#FFFFFF',
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'
    })
    row_format2 = workbook.add_format({
        'bg_color': '#DCE6F1',
        'border_color': '#FFFFFF',
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'
    })

    # write the column headers
    for idx, col_name in enumerate(df.columns):
        val = col_name
        start_row = end_row = 0
        start_col = end_col = idx
        worksheet.write(start_row, start_col, val, header_format)

    # write the data rows
    cols = list(df.columns)
    for counter in range(1,len(cols)+1):
        col = cols[0:counter]
        # print(f'col: {col}')
        if sort_df:
            df = df.sort_values(by=col)
        for idx, row_value in enumerate(merged_indices(df,col).values):
            val = row_value[-4]
            start_row = row_value[-3] + row_padding
            start_col = end_col = row_value[-1]
            end_row = row_value[-2] + row_padding
            # print(f'val: {val}, start_row: {start_row}, start_col: {start_col}, end_row: {end_row}')
            if start_row == end_row:
                worksheet.write(start_row, start_col, val, row_format1 if idx%2==0 else row_format2)
            else:
                worksheet.merge_range(start_row, start_col, end_row, end_col, val, row_format1 if idx%2==0 else row_format2)
            
    workbook.close()