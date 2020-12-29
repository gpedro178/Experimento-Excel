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
        
        indiceProvincia = clientesProvincia.index(provincia)
        
        clientesGasto_total_super_BSAS += float(
            clientesSupermercado[indiceProvincia].value)

clientesGasto_prom_super_BSAS = round( clientesGasto_total_super_BSAS / clientesEn_BSAS, 2)

print("Gasto Promedio en Supermercado de Clientes que Viven en BS AS: $", clientesGasto_prom_super_BSAS)
print()

############
# Quiénes gastan más en total? solteros o casados?
############

clientesEstado_civil = excelSheet["W"]
# La columna "C" o "INGRESO" es equivalente a la suma del gasto en todas las otras columnas por lo que se usa como gasto total.
clientesGasto_total = excelSheet["C"]

clientesGasto_total_solteros = 0
clientesGasto_total_casados = 0

for estado in clientesEstado_civil[1:]:
    if estado.value == "Soltero":

        indiceEstado = clientesEstado_civil.index(estado)

        clientesGasto_total_solteros += float(
            clientesGasto_total[indiceEstado].value)

    elif estado.value == "Casado":

        indiceEstado = clientesEstado_civil.index(estado)

        clientesGasto_total_casados += float(
            clientesGasto_total[indiceEstado].value)

print("Gasto Total de Solteros: $", round(
    clientesGasto_total_solteros,2))
print("Gasto Total de Casados: $", round(
    clientesGasto_total_casados,2))
print()

# Determinando qué grupo consume más

if clientesGasto_total_casados > clientesGasto_total_solteros:
    print("Los clientes casados consumen más en total que los clientes solteros")
elif clientesGasto_total_casados < clientesGasto_total_solteros:
    print("Los clientes solteros consumen más en total que los clientes casados")
else:
    print("Los clientes casados y solteros tienen el mismo nivel de consumo total")