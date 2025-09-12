import pandas as pd
import xlsxwriter as xw
import graph_creation as gc

section_headers = ["Income", "Expenditure", "Operating Surplus / Deficit before Contributions and Depn", "Intra-Regional Contribution", "Surplus / Deficit"]

def xlsx_create(df_dict, date, complete_path):
    """Creates an Excel file from a DataFrame."""
    with pd.ExcelWriter(complete_path, engine='xlsxwriter') as writer:
        sheet_name = f"{date}"
        row_offset = 0

        working_sheet = writer.book.add_worksheet(sheet_name)
        orange_format = writer.book.add_format({'bg_color': '#FFC000', 'bold': True, 'border': 1})
        bold_format = writer.book.add_format({'bold': True, 'border': 1})

        working_sheet.write(row_offset, 0, "Order of St John", orange_format)
        working_sheet.write(row_offset + 1, 0, "Income and Expenditure by Cost Centre", orange_format)
        working_sheet.write(row_offset + 2, 0, f"{date}", orange_format)
        row_offset += 5

        for branch, df in df_dict.items():
            df_out = df.drop(columns=['is_bold'], inplace=False)
            # headers with orange format
            for col_num, col_name in enumerate(df_out.columns):
                if col_num == 0:
                    working_sheet.write(row_offset, 0, branch, orange_format)
                else:
                    working_sheet.write(row_offset, col_num, col_name, orange_format)
            row_offset += 1

            # descriptions
            for _, row in df.iterrows():
                # Prepare the output row without 'is_bold'
                row_out = [value for key, value in row.items() if key != 'is_bold']
                if row['description'] in section_headers:
                    for col_num, value in enumerate(row_out):
                        working_sheet.write(row_offset, col_num, value, bold_format)
                    row_offset += 1
                else:
                    for col_num, value in enumerate(row_out):
                        if not row['is_bold']:
                            working_sheet.write(row_offset, 0, " " * 12 + row['description'])
                        working_sheet.write(row_offset, col_num, value)
                
                row_offset += 1

            # adjust width of columns based on contents
            for col_num, column_name in enumerate(df.columns):
                column_len = max(df[column_name].astype(str).map(len).max(), len(column_name))
                working_sheet.set_column(col_num, col_num, column_len)

            row_offset += 4 # Add 3 rows for a gap before the next table

    