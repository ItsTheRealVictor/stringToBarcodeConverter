from barcode import Code128
from barcode.writer import ImageWriter
import openpyxl
import fitz

wb = openpyxl.Workbook()
ws = wb.worksheets[0]

numberOfBarcodes = int(input("How many strings do you want to make into barcodes? "))
stringList = ['PlaceHolder'] * numberOfBarcodes
for index, string in enumerate(stringList):
    ask = input(f'Enter string #{index + 1}: ')
    stringList[index] = ask

def convertToBarcode():
    for string in stringList:
        converted = Code128(string, writer=ImageWriter())
        converted.save(f"C:\\Users\\VD102541\\Desktop\\Barcodes\\{string}")

convertToBarcode()


# adds images to an excel file. This doesn't work well so I'm commenting it out for now.

# for index, string in enumerate(stringList):
#     ws[f'A{index + 1}'] = string
#     image = openpyxl.drawing.image.Image(f"C:\\Users\\VD102541\\Desktop\\Barcodes\\{string}.png")
#     image.anchor = f'E{index}'
#     ws.add_image(image)

# wb.save(f"C:\\Users\\VD102541\\Desktop\\Barcodes\\myCustomBarcodes.xlsx")





