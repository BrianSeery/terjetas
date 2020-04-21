import pymongo
from bson.objectid import ObjectId
import xlrd
import datetime
myclient = pymongo.MongoClient("mongodb://localhost:27017/")
mydb = myclient["Tarjetas"]
mycol = mydb["cupones_minuto"]


while True:
    print("-----------------------------------------------------------------------------------------------------------------------------------------------")
    print("opcion numero 0 -> salir, opcion numero 1 -> liquidacion, opcion numero 2 -> importe, opcion 3 -> importe y liquidacion, opcion 4 -> pendientes")
    print("-----------------------------------------------------------------------------------------------------------------------------------------------")
    tipo = int(input("que desea buscar?"))
    if tipo == 1:
        print("ah seleccionado opcion numero:", tipo)
        query = int(input("que numero de liquidacion ?"))
        mydoc = list(mycol.find({"liquidacion": query}))
        for element in mydoc:
            print(element)
        print("se han encontrado {} resultados".format(len(mydoc)))       
    elif tipo == 2:
        print("ah seleccionado opcion numero:", tipo)
        query = float(input("que importe ?"))
        mydoc = list(mycol.find({"importe": query}))
        for element in mydoc:
            print(element)
        print("se han encontrado {} resultados".format(len(mydoc)))
    elif tipo == 3:
        print("ah seleccionado opcion numero:", tipo)
        importe = float(input("que importe?"))
        liquidacion = int(input("que liquidacion"))
        mydoc = list(mycol.find({"importe": importe , "liquidacion": liquidacion}))
        for element in mydoc:
            print(element)
        print("se han encontrado {} resultados".format(len(mydoc)))
    elif tipo == 4:
        print("ah seleccionado opcion numero:", tipo)
        print("Se mostraran todos los cupones pendientes")
        mydoc = list(mycol.find({"liquidacion": ""}))
        for element in mydoc:
            print(element)
        print("se han encontrado {} resultados".format(len(mydoc)))
    else:
        break

