import pandas as pd
import xlsxwriter as xw

def xlsx_create(df_dict):
    """Creates an Excel file from a DataFrame."""
    with pd.ExcelWriter('output.xlsx', engine='xlsxwriter') as writer:
        
        sheet_name = 'Combined Data'
        row_offset = 0
        
        for branch, df in df_dict.items():
            df.to_excel(writer, sheet_name=sheet_name, startrow=row_offset, index=False)
            working_sheet = writer.sheets[sheet_name]
            
            working_sheet.write(row_offset, 0, branch)

            for col_num, column_name in enumerate(df.columns):
                column_len = max(df[column_name].astype(str).map(len).max(), len(column_name))
                working_sheet.set_column(col_num, col_num, column_len)

            row_offset += len(df) + 4  # Add 3 rows for a gap