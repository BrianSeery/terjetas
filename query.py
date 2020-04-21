import pymongo
from bson.objectid import ObjectId
import xlrd
import datetime
import pprint
import xlsxwriter
contador = 0
no_contador = 0
no_match = []
liq = int(input("Cual es el numero de liquidacion?:"))
#Coneccion base de datos
myclient = pymongo.MongoClient("mongodb://localhost:27017/")
mydb = myclient["Tarjetas"]
mycol = mydb["cupones_minuto"]
sinte = mydb["Liquidacion"]
loc = ("sinteplast.xlsx")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
#abro libro pendientes
workbook = xlsxwriter.Workbook('pendientes.xlsx')
worksheet = workbook.add_worksheet()
#comienzo a recorrer
for _ in range(1,365):
    fecha = sheet.cell_value(_,6)
    fecha = fecha.split('.')
    fecha = datetime.datetime(year=int(fecha[2]),month= int(fecha[1]),day= int(fecha[0]))
    dir = {
        'descripcion': sheet.cell_value(_,0),
        'desc': sheet.cell_value(_,1),
        'importe': sheet.cell_value(_,5),
        'fecha': fecha,
        'cuotas': sheet.cell_value(_,7),
        'lote': sheet.cell_value(_,8),
        'liqui': ''
    }

    myquery = { "liquidacion": "" , "importe": dir.get('importe') , "fecha": dir.get('fecha')}
    if mycol.count_documents(myquery) == 1:
        x = mycol.find_one(myquery)
        id = x.get("_id")
        mycol.update_one({"_id" : id}, {'$set':{"liquidacion": liq}})
        dir["liqui"]=1
        contador += 1
    elif mycol.count_documents(myquery) == 0:
        no_match.append(dir)
        no_contador += 1
    else:
        mydoc = list(mycol.find(myquery))
        lon = len(mydoc)
        print(lon)
        s = 0
        print(dir)
        for _ in mydoc:
            while s < lon:
                print(mydoc[s])
                s+=1
        eleccion = int(input("Cual de desea liquidar? indicar con numero,comienza en 0:"))
        id = mydoc[eleccion].get("_id")        
        mycol.update_one({"_id" : id}, {'$set':{"liquidacion": liq}})
        dir["liqui"] = 1
        contador += 1    

    #Muestro lo que cargo
co = 1
for elements in no_match:
    fila = 'A' + str(co)
    print(elements)
    str(elements)
    print(type(str(elements)))
    worksheet.write(fila,str(elements))
    co += 1
workbook.close()
print("se cargaron: {} cupones a la base de datos".format(contador))
print("elementos que no se encontraron:", no_contador)
