import openpyxl
excelWorkbook = openpyxl.load_workbook("Experimento-Excel/MuestraTarjetaX2019.xlsx")
excelSheet = excelWorkbook.active

############
# Cantidad de clientes que viven fuera de BS AS
############

clientesProvincia = excelSheet["B"]
clientesFuera_de_BSAS = 0

for provincia in clientesProvincia[1:]:
    if provincia.value != "Buenos Aires":
        clientesFuera_de_BSAS += 1

print()
print("Cantidad de clientes que viven fuera de BS AS:", clientesFuera_de_BSAS)
print()

############
# Gasto promedio en supermercado de clientes en BSAS
############

clientesSupermercado = excelSheet["D"]

clientesGasto_total_super_BSAS = 0
clientesEn_BSAS = 0

for provincia in clientesProvincia[1:]:
    if provincia.value == "Buenos Aires":
        
        clientesEn_BSAS += 1
        
        indice = clientesProvincia.index(provincia)
        
        clientesGasto_total_super_BSAS += float(
            clientesSupermercado[indice].value)

clientesGasto_prom_super_BSAS = round( clientesGasto_total_super_BSAS / clientesEn_BSAS, 2)

print("Gasto Promedio en Supermercado de Clientes que Viven en BS AS: $", clientesGasto_prom_super_BSAS)
print()