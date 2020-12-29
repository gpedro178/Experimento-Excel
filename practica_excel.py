import openpyxl
excelWorkbook = openpyxl.load_workbook("Experimento-Excel/MuestraTarjetaX2019.xlsx")
excelSheet = excelWorkbook.active

print("Total Filas:", excelSheet.max_row)
print("Total Columnas:", excelSheet.max_column)
print()

clientesProvincia = excelSheet["B"]
