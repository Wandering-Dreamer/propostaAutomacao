import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

df = pd.read_excel("SA38 Test.xlsx")
unique_values = df.iloc[:, 16].drop_duplicates().index
print(unique_values)
merge_df = [2]

writer = pd.ExcelWriter("pandas_to_excel.xlsx") 
for i in unique_values:
    if i == 1:
        x = 2
        fl = df.iloc[i:(unique_values[x]), 16]
        fl.to_excel(writer, sheet_name='sheetName', index=False, startrow=1, startcol=0, header=False)
        df2 = df.iloc[i:(unique_values[x] - 1), [8, 9, 11, 2]]
        df2.columns = ['Part Number', 'Hardware', 'Qty', 'Serial Number']
        df2.to_excel(writer, sheet_name='sheetName', index=False, startrow=2, startcol=0)
        for column in df2:
            column_length = max(df2[column].astype(str).map(len).max(), len(column))
            col_idx = df2.columns.get_loc(column)
            writer.sheets['sheetName'].set_column(col_idx, col_idx, column_length)
        lenght = len(df2)

    if i > 1 and i != unique_values[-1]:
        x += 1
        fl = df.iloc[i:(unique_values[x]), 16]
        fl.to_excel(writer, sheet_name='sheetName', index=False, startrow=1+lenght+3, startcol=0, header=False)
        y = 1 + lenght + 4
        merge_df.append(y)
        print(merge_df)
        df2 = df.iloc[i:(unique_values[x] - 1), [8, 9, 11, 2]]
        df2.columns = ['Part Number', 'Hardware', 'Qty', 'Serial Number']
        df2.to_excel(writer, sheet_name='sheetName', index=False, startrow=1+lenght+4, startcol=0)
        for column in df2:
            column_length = max(df2[column].astype(str).map(len).max(), len(column))
            col_idx = df2.columns.get_loc(column)
            writer.sheets['sheetName'].set_column(col_idx, col_idx, column_length)
        lenght = lenght + len(df2) + 4
            
    if i == unique_values[-1]:
        fl = df.iloc[i:, 16]
        fl.to_excel(writer, sheet_name='sheetName', index=False, startrow=1+lenght+3, startcol=0, header=False)
        y = 1 + lenght + 4
        merge_df.append(y)
        print(merge_df)
        df2 = df.iloc[i:, [8, 9, 11, 2]]
        df2.columns = ['Part Number', 'Hardware', 'Qty', 'Serial Number']
        df2.to_excel(writer, sheet_name='sheetName', index=False, startrow=1+lenght+4, startcol=0)
        for column in df2:
            column_length = max(df2[column].astype(str).map(len).max(), len(column))
            col_idx = df2.columns.get_loc(column)
            writer.sheets['sheetName'].set_column(col_idx, col_idx, column_length) 

writer.close()

x = 0
for i in merge_df:
    wb = load_workbook("pandas_to_excel.xlsx")
    ws = wb.active
    start_row = merge_df[x]
    print(start_row)
    end_row = merge_df[x]
    start_column = 1
    end_column = 4
    merge_range = f"{ws.cell(row=start_row, column=start_column).coordinate}:{ws.cell(row=end_row, column=end_column).coordinate}"
    ws.merge_cells(merge_range)
    currentCell = ws.cell(start_row, start_column)
    currentCell.alignment = Alignment(horizontal='center')
    wb.save("pandas_to_excel.xlsx")
    x += 1