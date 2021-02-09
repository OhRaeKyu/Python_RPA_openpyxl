from openpyxl import load_workbook

wb = load_workbook("sample.xlsx")
ws = wb.active

ws.delete_rows(8, 3)  # 8번 째 줄부터 세줄 삭제
wb.save("sample_delete_rows.xlsx")

ws.delete_cols(2, 2) # 2번 째 열부터 두열 삭제
wb.save("sample_delete_cols.xlsx")