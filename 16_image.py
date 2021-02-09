from openpyxl import Workbook
from openpyxl.drawing.image import Image

wb = Workbook()
ws = wb.active

img = Image("img.png")

ws.add_image(img, "C3") # 'C3' 위치에 이미지 삽입

wb.save("sample_img.xlsx")