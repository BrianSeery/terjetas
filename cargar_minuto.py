import pymongo
import xlrd
import datetime
#Coneccion con mondodb
contador = 0
myclient = pymongo.MongoClient("mongodb://localhost:27017/")
mydb = myclient["Tarjetas"]
mycol = mydb["cupones_minuto"]
#importacion de los datos del excel
loc = ("Cupones tarjetas Sinteplast2020.xls")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(2)
#bucle para lectura de los datos
for _ in range(2,6072):
    tupla = xlrd.xldate_as_tuple(sheet.cell_value(_,5), wb.datemode)
    time = datetime.datetime(tupla[0],tupla[1],tupla[2])
    print(time)
    dir = {
        'local': sheet.cell_value(_,1),
        'numero': sheet.cell_value(_,2),
        'descripcion': sheet.cell_value(_,3),
        'cuotas': sheet.cell_value(_,4),
        'fecha': time,
        'lote': sheet.cell_value(_,6),
        'importe': sheet.cell_value(_,7),
        'liquidacion': sheet.cell_value(_,8)
    }
    contador += 1
    print(dir)      #Muestro lo que cargo
    x = mycol.insert_one(dir)       #inserto el JSON en la base de datos.
print("se cargaron: {} cupones a la base de datos".format(contador))