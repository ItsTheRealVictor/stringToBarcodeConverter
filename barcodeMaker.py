from barcode import Code128
from barcode.writer import ImageWriter
import openpyxl

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

# ws['A1'] = stringOne
# imageOne = openpyxl.drawing.image.Image(r"C:\Users\VD102541\Desktop\PDX_Archive_Barcode.png")
# imageOne.anchor = 'E1'
# ws.add_image(imageOne)

# ws['A20'] = stringTwo
# imageTwo = openpyxl.drawing.image.Image(r'C:\Users\VD102541\Desktop\PDX_DeArchive_Barcode.png')
# imageTwo.anchor = 'E20'
# ws.add_image(imageTwo)

# ws['A40'] = stringThree
# imageThree = openpyxl.drawing.image.Image(r"C:\Users\VD102541\Desktop\PDX_ArchiveRecycle_Barcode.png")
# imageThree.anchor = 'E40'
# ws.add_image(imageThree)

# wb.save(r"C:\Users\VD102541\Desktop\barcodes.xlsx")
