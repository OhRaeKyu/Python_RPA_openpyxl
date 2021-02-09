from openpyxl import load_workbook

wb = load_workbook("sample.xlsx")
ws = wb.active

# ws.insert_rows(8, 5)  # '8번' 째 줄 위치에 '다섯'줄 추가
# wb.save("sample_insert_rows.xlsx")

ws.insert_cols(2, 3)  #'2번' 째 열 위치에 '세'열 추가
wb.save("sample_insert_cols.xlsx")