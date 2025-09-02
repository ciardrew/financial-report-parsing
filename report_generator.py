import pandas as pd
import xlsxwriter as xw

def xlsx_create(df):
    """Creates an Excel file from a DataFrame."""
    with pd.ExcelWriter('output.xlsx', engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='output', index=False)
        working_sheet = writer.sheets['output']

        # set column widths based on contents 
        for index, col in enumerate(df):
            column_len = max(df[col].astype(str).map(len).max(), len(col))
            working_sheet.set_column(index, index, column_len)