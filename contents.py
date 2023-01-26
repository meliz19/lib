import pandas as pd

# Summary info about table.
def contents(df, output_filename):
    '''
    Overview: Create Excel writer.
    
    Variables:
    df: the dataframe to analyze
    output_filename: the name of the output file

    '''
    output_filename = output_filename + '.xlsx'
    with pd.ExcelWriter(output_filename) as writer:

        # Table-Level Details
        table_contents = pd.DataFrame(data=None)
        table_contents.insert(0, 'no_rows', [df.shape[0]])
        table_contents.insert(1, 'no_col', [df.shape[1]])
        table_contents.insert(2, 'dup_rows', df.duplicated().sum())
        table_contents.insert(3, 'all_nan_rows', df.isna().all().sum())
        table_contents.to_excel(writer,
                                sheet_name='table_contents',
                                index=False)

        # Column-Level Details
        col_contents = pd.DataFrame(data=None)
        col_contents.insert(0, 'variable', df.columns)
        col_contents.insert(1, 'dtype', list(df.dtypes))
        col_contents.insert(2, 'count_excluding_nan', list(df.count()))
        col_contents.insert(3, 'count_nan', list(df.isna().sum()))
        col_contents.insert(4, 'pct_missing', col_contents.iloc[:,3]/(col_contents.iloc[:,2] + col_contents.iloc[:,3]) * 100)
        col_contents.to_excel(writer, sheet_name='col_contents')

        # Object-Level Details
        var_objects = df.select_dtypes(include=object)

        obj_contents = pd.DataFrame(data=None)
        obj_contents.insert(0, 'variable', var_objects.columns)
        obj_contents.insert(1, 'no_levels', list(var_objects.nunique()))
        obj_contents.insert(2, 'no_levels_all', list(var_objects.nunique(
            dropna=False)))

        var_level_le10 = list(obj_contents.where(
            obj_contents.no_levels_all <= 10).dropna(how='any')['variable'])

        obj_level_contents = pd.DataFrame(data=None, index=var_level_le10)
        obj_level_contents.insert(0, 'levels', None)
        for var in var_objects:
            levels = list(var_objects[var].unique())

            if len(levels) > 10:
                obj_level_contents.loc[var, 'levels'] = levels[:10]
            else:
                obj_level_contents.loc[var, 'levels'] = levels

        obj_contents = obj_contents.set_index('variable').join(
            obj_level_contents)

        obj_contents.to_excel(writer, sheet_name='obj_contents')

        df.describe().transpose().to_excel(writer, sheet_name='num_stats')

    print(f'The output_filename ({output_filename}) has been saved!')
    