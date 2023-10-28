from openpyxl import Workbook
from openpyxl.styles import Color, PatternFill, Font, Border
from PIL import Image

redFill = PatternFill(start_color='FFFF0000',
                   end_color='FFFF0000',
                   fill_type='solid')

wb = Workbook()
ws = wb.active

# ws['A1'] = 2
# ws['A1'].fill = redFill


img = Image.open("img.jpg")
im = img.load()
for x in range(img.size[0]):
    for y in range(img.size[1]):
        col = im[x,y]

        hexCol = 'FF'
        for i in range(3):
            hexCol += hex(col[0])[2:]
        print(hexCol)
        excelCol = redFill = PatternFill(start_color=hexCol,
                   end_color=hexCol,
                   fill_type='solid')

wb.save("test.xlsx")