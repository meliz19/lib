import pandas as pd
import os

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
        .assign(index=range(len(df))) #creates an index column
        .groupby(col)['index']
        .agg(['first', 'last'])
        .assign(idx=list(df.columns).index(col[-1]))
        .reset_index() # moves row values from index to columns
    ) 
    return merged_indices_df 

def create_formatting(workbook):
    header_format = workbook.add_format({
            'bg_color': '#4F81BD',
            'bold': True,
            'font_color': '#FFFFFF',
            'border_color': '#FFFFFF',
            'border': 2})
    
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
    return header_format, row_format1, row_format2

def write_column_headers(df, worksheet, header_format):
    # write the column headers
    for idx, col_name in enumerate(df.columns):
        # start_row = 0; start_col = idx; cell_value = col_name; cell_format = header_format
        worksheet.write(0, idx, col_name, header_format)

def write_merged_row_data(df, sortby, worksheet, row_format1, row_format2, writer, sheet_name):
    # write the data rows
    cols = sortby if sortby is not None else list(df.columns)
    row_padding = df.columns.nlevels + 1 if df.columns.nlevels >1 else 1 # accounts for column headers
    for counter in range(1,len(cols)+1):
        col = cols[0:counter]
        df = df.sort_values(by=col)
        for idx, row_value in enumerate(merged_indices(df,col).values):
            val = row_value[-4]
            start_row = int(row_value[-3] + row_padding)
            start_col = end_col = int(row_value[-1])
            end_row = int(row_value[-2] + row_padding)
            # print(f'val: {val}, start_row: {start_row}, start_col: {start_col}, end_row: {end_row}')
            if start_row == end_row:
                worksheet.write(start_row, start_col, val, row_format1 if idx%2==0 else row_format2)
            else:
                # print(f'start_row: {start_row}\t{type({start_row})}\nstart_col: {start_col}\t{type(start_col)}\nend_row: {end_row}\t{type(start_col)}\nend_col: {end_col}\t{type(end_col)}\nval: {val}\t{type(val)}')
                worksheet.merge_range(start_row, start_col, end_row, end_col, val, row_format1 if idx%2==0 else row_format2)
    if sortby:
        df_remaining = df.reset_index().drop(columns=['index'] + cols)
        df_remaining.to_excel(writer, 
            sheet_name=sheet_name, 
            index=False, 
            header=False, 
            merge_cells=False,
            startrow=row_padding,
            startcol=len(sortby))
        for idx in range(row_padding, df_remaining.shape[0]+row_padding):
            start_row = end_row = idx
            start_col = 0
            end_col = len(df.columns)-1
            worksheet.conditional_format(start_row, start_col,end_row, end_col, 
                {'type': 'no_errors', 'format': row_format1 if idx%2==0 else row_format2})

def write_grpby_color_row_data(df, sortby, worksheet, row_format1, row_format2, writer, sheet_name):
    # write the data rows
    cols = sortby if sortby is not None else list(df.columns)
    row_padding = df.columns.nlevels + 1 if df.columns.nlevels >1 else 1 # accounts for column headers
    for counter in range(1,len(cols)+1):
        col = cols[0:counter]
        df = df.sort_values(by=col)
        for idx, row_value in enumerate(merged_indices(df,col).values):
            val = row_value[-4]
            start_row = int(row_value[-3] + row_padding)
            col_idx = int(row_value[-1])
            end_row = int(row_value[-2] + row_padding)
            for row_idx in range(start_row, end_row+1):
                worksheet.write(row_idx, col_idx, val, row_format1 if idx%2==0 else row_format2)
    if sortby:
        df_remaining = df.reset_index().drop(columns=['index'] + cols)
        df_remaining.to_excel(writer, 
            sheet_name=sheet_name, 
            index=False, 
            header=False, 
            merge_cells=False,
            startrow=row_padding,
            startcol=len(sortby))
        for idx in range(row_padding, df_remaining.shape[0]+row_padding):
            start_row = end_row = idx
            start_col = 0
            end_col = len(df.columns)-1
            worksheet.conditional_format(start_row, start_col,end_row, end_col, 
                {'type': 'no_errors', 'format': row_format1 if idx%2==0 else row_format2})

def df_to_excel_sheet(df, writer, workbook, sheet_name, sortby, merge_cells):
    '''
    This function takes a single level column and and single level index pd.DataFrame 
    and merges each column's row values based on the previous column's groupings.

    '''
    df.to_excel(writer, 
        sheet_name=sheet_name, 
        index=False, 
        header=False, 
        merge_cells=False)

    # Create an Excel worksheet object
    worksheet = writer.sheets[sheet_name]

    header_format, row_format1, row_format2 = create_formatting(workbook)

    # # write the column headers
    write_column_headers(df, worksheet, header_format)

    # write the data rows
    if merge_cells:
        write_merged_row_data(df, sortby, worksheet, row_format1, row_format2, writer, sheet_name)
    else:
        write_grpby_color_row_data(df, sortby, worksheet, row_format1, row_format2, writer, sheet_name)

    print(f'Sheet ({sheet_name}) created.')

def single_df_to_excel_book(df, filename, sheet_name, sortby=None, merge_cells=True):
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        # Create the Excel workbook (aka file) object
        workbook  = writer.book
        df_to_excel_sheet(df, writer, workbook, sheet_name, sortby=sortby, merge_cells=merge_cells)
        
    print(f'File ({os.path.join(os.getcwd(),filename)}) has been saved!')

def multiple_df_to_excel_book(filename, func):
    ''' Inputs: 
        filename
        func: problem specific code (func)
    Notes:
        The func function must include the df_to_excel_sheet(df,writer,workbook,sheet_name,sort_df=True)
        function to create the individual sheets.
    Example:

    bool = True
    filename=f'multiple_df_to_excel_book_merge_cells_{bool}.xlsx'
    df_init = pd.DataFrame({
       "lev1": [1, 1, 1, 1, 2, 2, 2, 1, 1],
       "lev2": [1, 1, 2, 1, 1, 1, 2, 3, 1],
       "lev3": [1, 2, 1, 2, 1, 1, 2, 1, 1],
       "lev4": [1, 1, 3, 1, 5, 6, 7, 1, 2],
       "values": [0, 1, 2, 3, 4, 5, 6, 1, 7]})
       
    def func(writer, workbook): 
        #problem specific specific code
            lev1 = df_init.lev1.unique()
            for lev in lev1:
                df_lev = df_init.query('lev1==@lev').fillna('')
                df_to_excel_sheet(
                    df=df_lev, 
                    writer=writer, 
                    workbook=workbook, 
                    sheet_name=str(lev), 
                    sortby=None, #list(df_lev.columns)[0:2],
                    merge_cells=bool)

    multiple_df_to_excel_book(filename=filename, func=func)
    
    ''' 
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        # Create the Excel workbook (aka file) object
        workbook  = writer.book

        # func must include the df_to_excel_sheet function to create the sheets of the workbook.
        # Its inputs must include writer and workbook.
        func(writer,workbook)
        
    print(f'File ({os.path.join(os.getcwd(),filename)}) has been saved!')
                