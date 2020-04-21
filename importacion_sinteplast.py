import pymongo
import xlrd
import datetime
#Coneccion con mondodb
myclient = pymongo.MongoClient("mongodb://localhost:27017/")
mydb = myclient["Tarjetas"]
mycol = mydb["Liquidacion"]
#importacion de los datos del excel
loc = ("sinteplast.xlsx")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
#bucle para lectura de los datos
for _ in range(1,391):
    fecha = sheet.cell_value(_,6)
    fecha = fecha.split('.')
    fecha = datetime.datetime(year=int(fecha[2]),month= int(fecha[1]),day= int(fecha[0]))
    lote = sheet.cell_value(_,8)
    print (lote)
    dir = {
        'descripcion': sheet.cell_value(_,0),
        'desc': sheet.cell_value(_,1),
        'importe': sheet.cell_value(_,5),
        'fecha': fecha,
        'cuotas': sheet.cell_value(_,7),
        'lote': lote,
    }
    print(dir)    #Muestro lo que cargo
    x = mycol.insert_one(dir)       #inserto el JSON en la base de datos.