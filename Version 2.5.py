import csv
from numpy import random
from random import randint

print 'Escoja una de las siguientes opciones'
print 'Opcion 1 para generar los datos'
print 'Opcion 2 para resolver los datos ya puestos'

decision1 = raw_input('escriba el numero de opcion: ')

ruta_archivo = 0
if decision1 == '1':
    ruta_archivo = raw_input('escriba la ubicacion de los archivo que se modificaran '
                           '(ejemplo -> /Users/rodrigo/Desktop/logistica/instancias generadas/): ')

    # Generar datos para el centro de distribucion

    encabezado = ['DCID;DCLocationLatitude;DCLocationLongitude;DCTransportationCapacityLiters;HandlingCostPerLiter']
    # print encabezado

    Numero_centros_distribucion = int(raw_input("cantidad de centros de distribucion para esta instancias: "))
    # Numero_centros_distribucion=4

    coordenada1 = ''
    coordenada2 = ''
    capacidad_transporte = 0
    costo_manejo = 0

    data = []

    data.append(encabezado)
    for i in range(Numero_centros_distribucion):
        coordenada1 = str(randint(-9, 9)) + '.' + str(randint(0, 9)) + str(randint(0, 9)) + str(
            randint(0, 9)) + '.' + str(randint(0, 9)) + str(randint(0, 9)) + str(
            randint(0, 9)) + '.' + str(randint(0, 9)) + str(randint(0, 9)) + str(
            randint(0, 9))
        coordenada2 = str(randint(-9, 9)) + '.' + str(randint(0, 9)) + str(randint(0, 9)) + str(
            randint(0, 9)) + '.' + str(randint(0, 9)) + str(randint(0, 9)) + str(
            randint(0, 9)) + '.' + str(randint(0, 9)) + str(randint(0, 9)) + str(
            randint(0, 9))
        capacidad_transporte = randint(0, 100)
        costo_manejo = randint(0, 100)
        data.append(['DC00' + str(i + 1) + ';' + str(coordenada1) + ';' + str(coordenada2) + ';' + str(
            capacidad_transporte) + ';' + str(costo_manejo)])

    # print data

    myFile = open(str(ruta_archivo) + 'DCs.csv', 'w')
    with myFile:
        writer = csv.writer(myFile)
        writer.writerows(data)

    archivo = []

    # generar los datos del csv de plantas

    encabezado = [
        'PlantID;PlantLocationLatitude;PlantLocationLongitude;PlantProdCapacityLiters;PlantRailTransportationCapacityLiters;PlantRoadTransportationCapacityLiters']

    Numero_plantas = int(raw_input("cantidad de plantas para esta instancias (hasta 5): "))
    # Numero_plantas=3

    data = []
    coordenada1 = ''
    coordenada2 = ''
    capacidad_produccion = 0
    capacidad_transporte_tren = 0
    capacidad_transporte_camion = 0

    data.append(encabezado)
    for i in range(Numero_plantas):
        coordenada1 = str(randint(-9, 9)) + '.' + str(randint(0, 9)) + str(randint(0, 9)) + str(
            randint(0, 9)) + '.' + str(randint(0, 9)) + str(randint(0, 9)) + str(
            randint(0, 9)) + '.' + str(randint(0, 9)) + str(randint(0, 9)) + str(
            randint(0, 9))
        coordenada2 = str(randint(-9, 9)) + '.' + str(randint(0, 9)) + str(randint(0, 9)) + str(
            randint(0, 9)) + '.' + str(randint(0, 9)) + str(randint(0, 9)) + str(
            randint(0, 9)) + '.' + str(randint(0, 9)) + str(randint(0, 9)) + str(
            randint(0, 9))
        capacidad_produccion = randint(0, 1000)
        capacidad_transporte_tren = randint(0, 500)
        capacidad_transporte_camion = randint(0, 500)
        data.append(['Plant00' + str(i + 1) + ';' + str(coordenada1) + ';' + str(coordenada2) + ';' + str(
            capacidad_produccion) + ';' + str(capacidad_transporte_tren) + ';' + str(capacidad_transporte_camion)])

    myFile = open(str(ruta_archivo) + 'plants.csv', 'w')
    with myFile:
        writer = csv.writer(myFile)
        writer.writerows(data)

    # generar el csv de detergente

    encabezado = ['DetergentID;Composition;PurityPercentage;Packaging;NominalPricePerLiter']

    Numero_detergentes = int(raw_input("cantidad de detergentes para esta instancias: "))

    data = []
    data.append(encabezado)

    composicion = ['C70', 'C71', 'C72', 'C73', 'C74', 'C75', 'C76', 'C77', 'C78', 'C79', 'C80']
    porcentajes_pureza = ['90', '95', '99']
    empaquetado = ['crystal40cc', 'crysta250cc', 'crysta500cc', 'crysta1L', 'plastic']
    precio_nominal = 0
    A = 0
    composicion_escogida = ''
    B = 0
    porcentajes_pureza_escogido = ''
    C = 0
    empaquetado_escogido = ''
    lista_productos = []

    cant_max = len(composicion) * len(porcentajes_pureza) * len(empaquetado)
    i = 0
    while i < Numero_detergentes:
        if Numero_detergentes < cant_max + 1:
            A = randint(0, len(composicion)-1)
            composicion_escogida = composicion[A]
            B = randint(0, len(porcentajes_pureza)-1)
            porcentajes_pureza_escogido = porcentajes_pureza[B]
            C = randint(0, len(empaquetado)-1)
            empaquetado_escogido = empaquetado[C]
            precio_nominal = randint(500, 2000)
            if str(A) + str(B) + str(C) not in lista_productos:
                lista_productos.append(str(A) + str(B) + str(C))
                data.append(['Det00' + str(i + 1) + ';' + str(composicion_escogida) + ';' + str(
                    porcentajes_pureza_escogido) + ';' + str(empaquetado_escogido) + ';' + str(precio_nominal)])
                i += 1

        else:
            print 'numero de detergentes supera la maxima cantidad permitida, se utilizara el numero maximo de: ' + str(
                cant_max)
            Numero_detergentes = cant_max

    myFile = open(str(ruta_archivo) + 'detergents.csv', 'w')
    with myFile:
        writer = csv.writer(myFile)
        writer.writerows(data)

    # generar los datos del csv de consumidores

    encabezado = ['CustomerID;CustomerLocationLatitude;CustomerLocationLongitude']

    Numero_clientes = int(raw_input("cantidad de clientes para esta instancias : "))
    # Numero_clientes=5

    data = []
    coordenada1 = ''
    coordenada2 = ''

    data.append(encabezado)
    for i in range(Numero_clientes):
        coordenada1 = str(randint(-9, 9)) + '.' + str(randint(0, 9)) + str(randint(0, 9)) + str(
            randint(0, 9)) + '.' + str(randint(0, 9)) + str(randint(0, 9)) + str(
            randint(0, 9)) + '.' + str(randint(0, 9)) + str(randint(0, 9)) + str(
            randint(0, 9))
        coordenada2 = str(randint(-9, 9)) + '.' + str(randint(0, 9)) + str(randint(0, 9)) + str(
            randint(0, 9)) + '.' + str(randint(0, 9)) + str(randint(0, 9)) + str(
            randint(0, 9)) + '.' + str(randint(0, 9)) + str(randint(0, 9)) + str(
            randint(0, 9))
        data.append(['Customer00' + str(i + 1) + ';' + str(coordenada1) + ';' + str(coordenada2)])

    myFile = open(str(ruta_archivo) + 'customers.csv', 'w')
    with myFile:
        writer = csv.writer(myFile)
        writer.writerows(data)

    # generar los datos del csv de ordenes

    encabezado = ['OrderID;ProductID;CustomerID;QuantityLiters']

    Numero_ordenes = int(raw_input("cantidad de ordenes para esta instancias : "))
    # Numero_ordenes=7

    # abrimos el archivo de productos para sacar su sku y el archivo de customers para sacar su id

    archivo1 = []
    with open(str(ruta_archivo) + 'detergents.csv') as File:
        reader = csv.reader(File, delimiter=';')
        for row in reader:
            archivo1.append(row)
    archivo1.pop(0)
    archivo2 = []
    with open(str(ruta_archivo) + 'customers.csv') as File:
        reader = csv.reader(File, delimiter=';')
        for row in reader:
            archivo2.append(row)
    archivo2.pop(0)

    data = []
    lista_sku = []
    lista_clientes_id = []
    A = 0
    B = 0
    sku_escogido = ''
    consumidor_escogido = ''
    cantidad_litros = 0

    for i in range(len(archivo1)):
        lista_sku.append(archivo1[i][0])

    for i in range(len(archivo2)):
        lista_clientes_id.append(archivo2[i][0])

    lista_ordenes = []
    data.append(encabezado)

    cant_max = len(archivo1) * len(archivo2)
    i = 0
    while i < Numero_ordenes:
        if Numero_ordenes < cant_max + 1:
            A = randint(0, len(lista_sku)-1)
            sku_escogido = lista_sku[A]
            B = randint(0, len(lista_clientes_id)-1)
            consumidor_escogido = lista_clientes_id[B]
            cantidad_litros = randint(1000, 2000)
            if str(A) + str(B) not in lista_ordenes:
                lista_ordenes.append(str(A) + str(B))
                data.append(['Order00' + str(i + 1) + ';' + str(sku_escogido) + ';' + str(
                    consumidor_escogido) + ';' + str(cantidad_litros)])
                i += 1

        else:
            print 'numero de ordenes supera la maxima cantidad permitida, se utilizara el numero maximo de: ' + str(
                cant_max)
            Numero_ordenes = cant_max

    myFile = open(str(ruta_archivo)+'orders.csv', 'w')
    with myFile:
        writer = csv.writer(myFile)
        writer.writerows(data)

    # generar csv costo variable

    encabezado = ['PlantID;DetergentID;costPerliter']

    archivo1 = []
    with open(str(ruta_archivo) + 'plants.csv') as File:
        reader = csv.reader(File, delimiter=';')
        for row in reader:
            archivo1.append(row)
    archivo1.pop(0)
    archivo2 = []
    with open(str(ruta_archivo) + 'detergents.csv') as File:
        reader = csv.reader(File, delimiter=';')
        for row in reader:
            archivo2.append(row)
    archivo2.pop(0)

    data = []
    A = 0
    B = 0
    costo_produccion = 0
    lista_costos_variables = []
    data.append(encabezado)

    for j in range(len(archivo2)):
        costo_produccion = random.randint(1, 40)
        for i in range(len(archivo1)):
            A = archivo1[i][0]
            B = archivo2[j][0]
            if str(A) + str(B) not in lista_costos_variables:
                lista_costos_variables.append(str(A) + str(B))
                data.append([str(A) + ';' + str(B) + ';' + str(costo_produccion)])

    myFile = open(str(ruta_archivo) + 'variable_costs.csv', 'w')
    with myFile:
        writer = csv.writer(myFile)
        writer.writerows(data)

    # generar archivo csv para routes_available2

    encabezado = ['RouteID;StartRoute;MiddleRoute;EndRoute;RailUse']

    archivo1 = []
    with open(str(ruta_archivo) + 'customers.csv') as File:
        reader = csv.reader(File, delimiter=';')
        for row in reader:
            archivo1.append(row)
    archivo1.pop(0)

    lista_clientes_id = []
    for i in range(len(archivo1)):
        lista_clientes_id.append(archivo1[i][0])

    archivo2 = []
    with open(str(ruta_archivo) + 'plants.csv') as File:
        reader = csv.reader(File, delimiter=';')
        for row in reader:
            archivo2.append(row)
    archivo2.pop(0)

    lista_plantas_id = []
    for i in range(len(archivo2)):
        lista_plantas_id.append(archivo2[i][0])

    archivo3 = []
    with open(str(ruta_archivo) + 'DCs.csv') as File:
        reader = csv.reader(File, delimiter=';')
        for row in reader:
            archivo3.append(row)
    archivo3.pop(0)

    lista_DC_id = []
    for i in range(len(archivo3)):
        lista_DC_id.append(archivo3[i][0])
    lista_DC_id.append('')

    # para cada cliente debe tener minimo una ruta y un maximo de rutas determinado por la cantidad de centros + la ruta directa (DC='')

    cantidad_plantas = 0
    cantidad_centros = 0
    rutas_cliente_cantDC_cantPlants = []
    for i in range(len(lista_clientes_id)):
        rutas_cliente_cantDC_cantPlants.append([])

    for i in range(len(lista_clientes_id)):
        cantidad_centros = randint(1, len(lista_DC_id))
        rutas_cliente_cantDC_cantPlants[i].append(lista_clientes_id[i])
        rutas_cliente_cantDC_cantPlants[i].append(cantidad_centros)

    for i in range(len(lista_clientes_id)):
        cantidad_plantas = randint(1, len(lista_plantas_id))
        rutas_cliente_cantDC_cantPlants[i].append(cantidad_plantas)

    cantidad_rutas = 0
    for i in range(len(rutas_cliente_cantDC_cantPlants)):
        cantidad_rutas = cantidad_rutas + (
                    rutas_cliente_cantDC_cantPlants[i][1] * rutas_cliente_cantDC_cantPlants[i][2])

    lista_rutas = []
    planta_escogida = ''
    DC_escogido = ''
    A = 0
    B = 0
    C = 1
    data = []
    data.append(encabezado)

    while C < 2 * cantidad_rutas + 1:
        for i in range(len(lista_clientes_id)):
            for j in range(rutas_cliente_cantDC_cantPlants[i][1]):
                for k in range(rutas_cliente_cantDC_cantPlants[i][2]):
                    for l in range(2):
                        A = randint(0, rutas_cliente_cantDC_cantPlants[i][2]-1)
                        planta_escogida = lista_plantas_id[A]
                        B = randint(0, rutas_cliente_cantDC_cantPlants[i][1]-1)
                        DC_escogido = lista_DC_id[B]
                        if str(A) + str(B) + str(i) + str(l) not in lista_rutas:
                            lista_rutas.append(str(A) + str(B) + str(i) + str(l))
                            data.append(['Route00' + str(C) + ';' + str(planta_escogida) + ';' + str(
                                DC_escogido) + ';' + str(lista_clientes_id[i]) + ';' + str(l)])
                            C += 1


    myFile = open(str(ruta_archivo)+'routes_available.csv', 'w')
    with myFile:
        writer = csv.writer(myFile)
        writer.writerows(data)

    # generar el csv para routes_endtoend2

    encabezado = ['Start;End;EndToEndCostPerLiter']

    archivo1 = []
    with open(str(ruta_archivo) + 'routes_available.csv') as File:
        reader = csv.reader(File, delimiter=';')
        for row in reader:
            archivo1.append(row)
    archivo1.pop(0)

    inicio_final = []
    for i in range(len(archivo1)):
        inicio_final.append([archivo1[i][1], archivo1[i][3]])

    inicio_final1 = []
    for i in range(len(inicio_final)):
        if inicio_final[i] not in inicio_final1:
            inicio_final1.append(inicio_final[i])

    data = []
    data.append(encabezado)
    costo_inicio_final = 0
    for i in range(len(inicio_final1)):
        costo_inicio_final = randint(1, 10)
        data.append([str(inicio_final1[i][0]) + ';' + str(inicio_final1[i][1]) + ';' + str(costo_inicio_final)])

    myFile = open(str(ruta_archivo) + 'routes_endtoend.csv', 'w')
    with myFile:
        writer = csv.writer(myFile)
        writer.writerows(data)

    # generar el csv de segmentos

    encabezado = ['SegmentStart;SegmentEnd;SegmentMode;SegmentCostPerUnit']

    archivo1 = []
    with open(str(ruta_archivo) + 'routes_available.csv') as File:
        reader = csv.reader(File, delimiter=';')
        for row in reader:
            archivo1.append(row)

    inicio_medio = []
    for i in range(len(archivo1)):
        if archivo1[i][2] != '':
            inicio_medio.append([archivo1[i][1], archivo1[i][2], archivo1[i][4]])

    inicio_medio1 = []
    for i in range(len(inicio_medio)):
        if inicio_medio[i] not in inicio_medio1:
            inicio_medio1.append(inicio_medio[i])

    data = []
    data.append(encabezado)
    costo_inicio_medio = 0
    for i in range(len(inicio_medio1)):
        costo_inicio_medio = randint(1, 10)
        if inicio_medio1[i][2] == '1':
            data.append([str(inicio_medio1[i][0]) + ';' + str(inicio_medio1[i][1]) + ';' + 'Rail' + ';' + str(
                costo_inicio_medio)])
        elif inicio_medio1[i][2] == '0':
            data.append([str(inicio_medio1[i][0]) + ';' + str(inicio_medio1[i][1]) + ';' + 'Road' + ';' + str(
                costo_inicio_medio)])

    medio_final = []
    for i in range(len(archivo1)):
        if archivo1[i][2] != '':
            medio_final.append([archivo1[i][2], archivo1[i][3], archivo1[i][4]])

    medio_final1 = []
    for i in range(len(medio_final)):
        if medio_final[i] not in medio_final1:
            medio_final1.append(medio_final[i])

    costo_medio_final = 0
    for i in range(len(medio_final1)):
        costo_medio_final = randint(1, 10)
        if medio_final1[i][2] == '1':
            data.append(
                [str(medio_final1[i][0]) + ';' + str(medio_final1[i][1]) + ';' + 'Rail' + ';' + str(costo_medio_final)])
        elif medio_final1[i][2] == '0':
            data.append(
                [str(medio_final1[i][0]) + ';' + str(medio_final1[i][1]) + ';' + 'Road' + ';' + str(costo_medio_final)])

    inicio_final = []
    for i in range(len(archivo1)):
        inicio_final.append([archivo1[i][1], archivo1[i][3], archivo1[i][4]])

    inicio_final1 = []
    for i in range(len(inicio_final)):
        if inicio_final[i] not in inicio_final1:
            inicio_final1.append(inicio_final[i])

    costo_inicio_final = 0
    for i in range(len(inicio_final1)):
        costo_inicio_final = randint(1, 10)
        if inicio_final1[i][2] == '1':
            data.append([str(inicio_final1[i][0]) + ';' + str(inicio_final[i][1]) + ';' + 'Rail' + ';' + str(
                costo_inicio_final)])
        elif inicio_final1[i][2] == '0':
            data.append([str(inicio_final1[i][0]) + ';' + str(inicio_final1[i][1]) + ';' + 'Road' + ';' + str(
                costo_inicio_final)])

    myFile = open(str(ruta_archivo) + 'routes_segment.csv', 'w')
    with myFile:
        writer = csv.writer(myFile)
        writer.writerows(data)


elif decision1 == '2':
    ruta_archivo = raw_input('escriba la ubicacion de donde se leeran los archivos: '
                           '(ejemplo -> /Users/rodrigo/Desktop/logistica/instancias generadas/): ')
    import csv

    #para esta parte se buscara generar los conjuntos necesarios para el problema de optimizacion
    # cargar archivo de rutas disponibles

    archivo = []

    with open(str(ruta_archivo) + 'routes_available.csv') as File:
        reader = csv.reader(File, delimiter=';')
        for row in reader:
            archivo.append(row)

    print 'generando conjunto de rutas ...'

    lista_rutas = []
    with open(str(ruta_archivo) + 'routes_available.csv') as File:
        reader = csv.reader(File, delimiter=';', quotechar=';')
        for row in reader:
            lista_rutas.append(row[0])

    lista_rutas.pop(0)

    # Rutas

    Rutas = []
    for i in range(len(lista_rutas) / 2):
        Rutas.append(i + 1)



    # Plantas

    print 'generando conjunto de plantas ...'

    lista_plantas = []
    with open(str(ruta_archivo) + 'routes_available.csv') as File:
        reader = csv.reader(File, delimiter=';', quotechar=';')
        for row in reader:
            lista_plantas.append(row[1])
    lista_plantas.pop(0)

    plantas_id = []
    for plant in range(len(lista_plantas)):
        if lista_plantas[plant] not in plantas_id:
            plantas_id.append(lista_plantas[plant])

    Plantas = []
    for plant in range(len(plantas_id)):
        Plantas.append(plant + 1)


    # Modo = [1, 2] el modo 1 corresponde a que se utilizo el tren y modo 2 a que solo se utilizo camion

    Modo = [1, 2]

    # Clientes

    print 'generando conjunto de clientes ...'

    lista_clientes = []
    with open(str(ruta_archivo) + 'routes_available.csv') as File:
        reader = csv.reader(File, delimiter=';')
        for row in reader:
            lista_clientes.append(row[3])
    lista_clientes.pop(0)

    clientes_id = []
    for client in range(len(lista_clientes)):
        if lista_clientes[client] not in clientes_id:
            clientes_id.append(lista_clientes[client])

    Clientes = []
    for client in range(len(clientes_id)):
        Clientes.append(client + 1)

    # Centro_distribucion

    print 'generando conjunto de centro de distribucion ...'

    lista_CD = []
    with open(str(ruta_archivo) + 'routes_available.csv') as File:
        reader = csv.reader(File, delimiter=';')
        for row in reader:
            lista_CD.append(row[2])
    lista_CD.pop(0)

    # CoD: center of distribution
    CoD = 0
    while CoD < len(lista_CD):
        if lista_CD[CoD] == '':
            lista_CD.pop(CoD)
            CoD = CoD
        else:
            CoD += 1

    CD_id = []
    for CoD in range(len(lista_CD)):
        if lista_CD[CoD] not in CD_id:
            CD_id.append(lista_CD[CoD])

    Centro_distribucion = []
    for CoD in range(len(CD_id)):
        Centro_distribucion.append(CoD + 1)

    # conjunto_rutas_directas

    print 'generando conjunto de rutas directas ...'

    # diccionario de rutas

    diccionario_de_rutas = {}
    archivo.pop(0)

    lista_valores_rutas = []
    for i in range(len(archivo)):
        for j in range(len(archivo)):
            if archivo[i][1] == archivo[j][1] and archivo[i][2] == archivo[j][2] and archivo[i][3] == archivo[j][
                3] and i != j:
                lista_valores_rutas.append([archivo[i][0], archivo[j][0]])

    lista_valores_rutas1 = []
    lista_valores_rutas2 = []
    lista_valores_rutas3 = []

    for i in range(len(lista_valores_rutas)):
        for j in range(len(lista_valores_rutas)):
            if lista_valores_rutas[i][0] == lista_valores_rutas[j][1] and lista_valores_rutas[i][1] == \
                    lista_valores_rutas[j][0] and i != j and [i + j, abs(i - j)] not in lista_valores_rutas1:
                lista_valores_rutas1.append([i + j, abs(i - j)])
                lista_valores_rutas2.append([lista_valores_rutas[i][0], lista_valores_rutas[j][0]])

    for i in range(len(lista_valores_rutas2)):
        diccionario_de_rutas[i + 1] = lista_valores_rutas2[i]

    lista_rutas_directas_id = []
    for row1 in range(len(archivo)):
        if archivo[row1][2] == '':
            lista_rutas_directas_id.append(archivo[row1][0])

    conjunto_rutas_directas = []
    row1 = 1
    while row1 < len(diccionario_de_rutas) + 1:
        for row2 in range(len(lista_rutas_directas_id)):
            if lista_rutas_directas_id[row2] in diccionario_de_rutas[row1] and row1 not in conjunto_rutas_directas:
                conjunto_rutas_directas.append(row1)
        row1 += 1

    # conjunto_rutas_cliente

    print 'generando conjunto de rutas a clientes...'

    # conjuto de rutas que terminan en un cliente en especifico

    conjunto_rutas_clientes_id = {}
    for client in Clientes:
        conjunto_rutas_clientes_id[client] = []

    clientes_id.sort()  # para ordenar la lista clientes de menor a mayor
    client = 0
    while client < len(Clientes):
        row1 = 1
        while row1 < len(archivo):
            if archivo[row1][3] == clientes_id[client]:
                conjunto_rutas_clientes_id[client + 1].append(archivo[row1][0])
                row1 += 1
            else:
                row1 += 1
        client += 1

    conjunto_rutas_cliente = {}

    for client in range(len(Clientes)):
        conjunto_rutas_cliente[client + 1] = []

    for client in range(len(Clientes)):
        for row1 in range(len(diccionario_de_rutas)):
            for mode in range(2):
                if diccionario_de_rutas[row1 + 1][mode] in conjunto_rutas_clientes_id[client + 1] and row1 + 1 not in \
                        conjunto_rutas_cliente[client + 1]:
                    conjunto_rutas_cliente[client + 1].append(row1 + 1)

    # conjunto_rutas_plantas

    print 'generando conjunto de rutas desde plantas ...'

    # conjunto de rutas que empiezan con una planta en especifico

    conjunto_rutas_plantas_id = {}
    for plant in Plantas:
        conjunto_rutas_plantas_id[plant] = []

    plantas_id.sort()  # para ordenar la lista plantas de menor a mayor
    plant = 0
    while plant < len(Plantas):
        row1 = 1
        while row1 < len(archivo):
            if archivo[row1][1] == plantas_id[plant]:
                conjunto_rutas_plantas_id[plant + 1].append(archivo[row1][0])
                row1 += 1
            else:
                row1 += 1
        plant += 1

    conjunto_rutas_plantas = {}

    for plant in range(len(Plantas)):
        conjunto_rutas_plantas[plant + 1] = []

    for plant in range(len(Plantas)):
        for row1 in range(len(diccionario_de_rutas)):
            for mode in range(2):
                if diccionario_de_rutas[row1 + 1][mode] in conjunto_rutas_plantas_id[plant + 1] and row1 + 1 not in \
                        conjunto_rutas_plantas[plant + 1]:
                    conjunto_rutas_plantas[plant + 1].append(row1 + 1)

    # conjunto_rutas_intermedias

    print 'generando conjunto de rutas intermedias ...'

    # conjunto de rutas que pasan por un centro de distribucion

    conjunto_rutas_intermedias_id = {}
    for CoD in range(len(Centro_distribucion)):
        conjunto_rutas_intermedias_id[CoD + 1] = []

    for CoD in range(len(Centro_distribucion)):
        for row1 in range(len(archivo)):
            if CD_id[CoD] in archivo[row1][2]:
                conjunto_rutas_intermedias_id[CoD + 1].append(archivo[row1][0])

    conjunto_rutas_intermedias = {}

    for CoD in range(len(Centro_distribucion)):
        conjunto_rutas_intermedias[CoD + 1] = []

    for CoD in range(len(Centro_distribucion)):
        for row1 in range(len(diccionario_de_rutas)):
            for mode in range(2):
                if diccionario_de_rutas[row1 + 1][mode] in conjunto_rutas_intermedias_id[CoD + 1] and row1 + 1 not in \
                        conjunto_rutas_intermedias[CoD + 1]:
                    conjunto_rutas_intermedias[CoD + 1].append(row1 + 1)

    # ahora abriremos otro csv para seguir sacando los parametros necesarios para el modelo

    archivo = []

    with open(str(ruta_archivo) + 'plants.csv') as File:
        reader = csv.reader(File, delimiter=';')
        for row in reader:
            archivo.append(row)

    archivo.pop(0)
    archivo.sort()

    # Capp capacidad de la planta para producir

    print 'extrayendo capacidad de produccion de plantas ...'

    Capp = {}
    for row1 in range(len(archivo)):
        Capp[row1 + 1] = float(archivo[row1][3])

    # capacidad_planta_tren

    print 'extrayendo capacidad de distribucion de plantas por tren ...'

    capacidad_planta_tren = {}
    for row1 in range(len(archivo)):
        capacidad_planta_tren[row1 + 1] = float(archivo[row1][4])

    # capacidad_planta_C capacidad de distribucion de una planta al usar camion

    print 'extrayendo capacidad de distribucion de plantas por camion...'

    capacidad_planta_C = {}
    for row1 in range(len(archivo)):
        capacidad_planta_C[row1 + 1] = float(archivo[row1][5])

    # ahora abriremos otro csv para seguir sacando los parametros necesarios para el modelo
    archivo = []

    with open(str(ruta_archivo) + 'DCs.csv') as File:
        reader = csv.reader(File, delimiter=';')
        for row in reader:
            archivo.append(row)

    archivo.pop(0)
    archivo.sort()

    # cd centro de distribucion

    print 'extrayendo costo de manejo por centro de distribucion ...'

    cd = {}
    for row1 in range(len(archivo)):
        cd[row1 + 1] = float(archivo[row1][3])

    # c_manejo
    # costo de manejo por CD

    c_manejo = {}
    for row1 in range(len(archivo)):
        c_manejo[row1 + 1] = float(archivo[row1][4])

    # ahora abriremos otro csv para seguir sacando los parametros necesarios para el modelo

    archivo = []

    with open(str(ruta_archivo) + 'detergents.csv') as File:
        reader = csv.reader(File, delimiter=';')
        for row in reader:
            archivo.append(row)

    archivo.pop(0)

    # p_venta

    print 'extrayendo precio de venta por producto ...'

    p_venta = {}
    for row1 in range(len(archivo)):
        p_venta[row1 + 1] = float(archivo[row1][4])

    # ahora abriremos otro csv para seguir sacando los parametros necesarios para el modelo

    # primero debemos hacer un diccionario de productos

    print 'generando conjunto de productos demandados ...'

    diccionario_de_productos = {}
    for row1 in range(len(archivo)):
        diccionario_de_productos[row1 + 1] = archivo[row1][0]

    archivo = []

    with open(str(ruta_archivo) + 'orders.csv') as File:
        reader = csv.reader(File, delimiter=';')
        for row in reader:
            archivo.append(row)

    archivo.pop(0)

    # d demanda que esta en funcion del producto y del consumidor d[produto, consumidor]

    diccionario_de_clientes = {}
    for i in range(len(clientes_id)):
        diccionario_de_clientes[i + 1] = clientes_id[i]

    productos_demandados = []

    for k in range(len(archivo)):
        for i in range(len(diccionario_de_productos)):
            if diccionario_de_productos[i + 1] in archivo[k][1]:
                productos_demandados.append(i + 1)


    lista_productos_demandados = []
    for i in range(len(productos_demandados)):
        if productos_demandados[i] not in lista_productos_demandados:
            lista_productos_demandados.append(productos_demandados[i])

    d = {}

    print 'extrayendo demanda de productos por clientes ...'

    for i in lista_productos_demandados:
        for j in Clientes:
            d[i, j] = 0

    for i in range(len(diccionario_de_productos)):
        for j in range(len(diccionario_de_clientes)):
            for k in range(len(archivo)):
                if archivo[k][1] in diccionario_de_productos[i + 1] and archivo[k][2] in diccionario_de_clientes[j + 1]:
                    d[i + 1, j + 1] = int(archivo[k][3])

    # ojo que no tiene porque ser demandado todos los productos

    Productos = []
    for row1 in range(len(diccionario_de_productos)):
        Productos.append(int(row1 + 1))

    # ahora abriremos otro csv para seguir sacando los parametros necesarios para el modelo

    # c_rutas

    print 'generando conjunto de rutas ocupadas ...'

    # costo modo, ruta

    archivo1 = []
    with open(str(ruta_archivo) + 'routes_available.csv') as File:
        reader = csv.reader(File, delimiter=';')
        for row in reader:
            archivo1.append(row)
    archivo1.pop(0)

    archivo2 = []
    with open(str(ruta_archivo) + 'routes_endtoend.csv') as File:
        reader = csv.reader(File, delimiter=';')
        for row in reader:
            archivo2.append(row)
    archivo2.pop(0)

    archivo3 = []
    with open(str(ruta_archivo) + 'routes_segment.csv') as File:
        reader = csv.reader(File, delimiter=';')
        for row in reader:
            archivo3.append(row)
    archivo3.pop(0)

    rutas_ocupadas = []

    for i in range(len(archivo2)):
        for k in range(len(archivo1)):
            for j in range(len(diccionario_de_rutas)):
                if archivo1[k][1] == archivo2[i][0] and archivo1[k][3] == archivo2[i][1] and archivo1[k][4] == '1' and \
                        archivo1[k][0] in diccionario_de_rutas[j + 1]:
                    rutas_ocupadas.append(j + 1)
                elif archivo1[k][1] == archivo2[i][0] and archivo1[k][3] == archivo2[i][1] and archivo1[k][4] == '0' and \
                        archivo1[k][0] in diccionario_de_rutas[j + 1]:
                    rutas_ocupadas.append(j + 1)

    lista_rutas_ocupadas = []
    for i in range(len(rutas_ocupadas)):
        if rutas_ocupadas[i] not in lista_rutas_ocupadas:
            lista_rutas_ocupadas.append(rutas_ocupadas[i])
    lista_rutas_ocupadas.sort()

    # costos inicio-final

    print 'extrayendo costos de rutas ...'

    c_rutas = {}

    for i in range(len(archivo2)):
        for k in range(len(archivo1)):
            for j in range(len(diccionario_de_rutas)):
                if archivo1[k][1] == archivo2[i][0] and archivo1[k][3] == archivo2[i][1] and archivo1[k][4] == '1' and \
                        archivo1[k][0] in diccionario_de_rutas[j + 1] and j + 1 in lista_rutas_ocupadas:
                    c_rutas[(1, j + 1)] = float(archivo2[i][2])
                elif archivo1[k][1] == archivo2[i][0] and archivo1[k][3] == archivo2[i][1] and archivo1[k][4] == '0' and \
                        archivo1[k][0] in diccionario_de_rutas[j + 1] and j + 1 in lista_rutas_ocupadas:
                    c_rutas[(2, j + 1)] = float(archivo2[i][2])

    # suma de costos para segmentos entre inicio y mitad en modo tren y camion

    for i in range(len(archivo3)):
        for k in range(len(archivo1)):
            for j in range(len(diccionario_de_rutas)):
                if archivo1[k][1] == archivo3[i][0] and archivo1[k][2] == archivo3[i][1] and archivo1[k][4] == '1' and \
                        archivo3[i][2] == 'Rail' and archivo1[k][0] in diccionario_de_rutas[
                    j + 1] and j + 1 in lista_rutas_ocupadas:
                    c_rutas[(1, j + 1)] = c_rutas[(1, j + 1)] + float(archivo3[i][3])
                elif archivo1[k][1] == archivo3[i][0] and archivo1[k][2] == archivo3[i][1] and archivo1[k][4] == '0' and \
                        archivo3[i][2] == 'Road' and archivo1[k][0] in diccionario_de_rutas[
                    j + 1] and j + 1 in lista_rutas_ocupadas:
                    c_rutas[(2, j + 1)] = c_rutas[(2, j + 1)] + float(archivo3[i][3])

    # suma de costos para segmentos entre mitad y final en modo camion

    for i in range(len(archivo3)):
        for k in range(len(archivo1)):
            for j in range(len(diccionario_de_rutas)):
                if archivo1[k][2] == archivo3[i][0] and archivo1[k][3] == archivo3[i][1] and archivo1[k][4] == '0' and \
                        archivo3[i][2] == 'Road' and archivo1[k][0] in diccionario_de_rutas[
                    j + 1] and j + 1 in lista_rutas_ocupadas:
                    c_rutas[(1, j + 1)] = c_rutas[(1, j + 1)] + float(archivo3[i][3])
                    c_rutas[(2, j + 1)] = c_rutas[(2, j + 1)] + float(archivo3[i][3])

    # suma de costos para segmentos entre inicio y final en camion

    for i in range(len(archivo3)):
        for k in range(len(archivo1)):
            for j in range(len(diccionario_de_rutas)):
                if archivo1[k][1] == archivo3[i][0] and archivo1[k][3] == archivo3[i][1] and archivo1[k][4] == '0' and \
                        archivo3[i][2] == 'Road' and archivo1[k][0] in diccionario_de_rutas[
                    j + 1] and j + 1 in lista_rutas_ocupadas:
                    c_rutas[(1, j + 1)] = c_rutas[(1, j + 1)] + float(archivo3[i][3])
                    c_rutas[(2, j + 1)] = c_rutas[(2, j + 1)] + float(archivo3[i][3])

    # c_produccion

    print 'extrayendo costos de producion ...'

    archivo = []
    with open(str(ruta_archivo) + 'variable_costs.csv') as File:
        reader = csv.reader(File, delimiter=';')
        for row in reader:
            archivo.append(row)
    archivo.pop(0)

    lista_costo_productos = []
    for row1 in range(len(archivo)):
        if archivo[row1][1] not in lista_costo_productos:
            lista_costo_productos.append(archivo[row1][1])

    c_produccion = {}
    for row1 in range(len(lista_costo_productos)):
        for row2 in range(len(archivo)):
            if archivo[row2][1] == lista_costo_productos[row1]:
                c_produccion[row1 + 1] = float(archivo[row2][2])

    # a partir de aqui trabajamos los datos cargados previamente y escribimos el codigo para trabajar con ellos

    print 'resolviendo ...'

    import pyscipopt
    from pyscipopt import Model, quicksum

    m = pyscipopt.Model()

    # variables de cantidad de producto transportado producto/modo/ruta
    # tren=1 y camion=2

    x = {}
    for product in Productos:
        for mode in Modo:
            for route in Rutas:
                x[product, mode, route] = m.addVar(lb=0, ub=None, name="x(%s,%s,%s)" % (product, mode, route))

    s = m.addVar("s", vtype="I", lb=None)

    # satisfacer la demanda
    for client in Clientes:
        for product in lista_productos_demandados:
            m.addCons(quicksum(x[product, mode, route] for mode in Modo for route in lista_rutas_ocupadas if route in conjunto_rutas_cliente[client]) <= d[product, client],
                      name="Demand(%s,%s)" % (product, client))


    # lo que sale de la planta debe ser menor o igual a capacidad de distribucion segun su modo
    for plant in Plantas:
        m.addCons(quicksum(
            x[product, 1, route] for product in lista_productos_demandados for route in lista_rutas_ocupadas if
            route in conjunto_rutas_plantas[plant]) <= capacidad_planta_tren[plant], name="CapPlantTren(%s)" % (plant))

    for plant in Plantas:
        m.addCons(quicksum(
            x[product, 2, route] for product in lista_productos_demandados for route in lista_rutas_ocupadas if
            route in conjunto_rutas_plantas[plant]) <= capacidad_planta_C[plant], name="CapPlantC(%s)" % (plant))

    # los trenes no puedene ir de planta a cliente
    for product in lista_productos_demandados:
        for mode in Modo:
            for route in lista_rutas_ocupadas:
                m.addCons(x[product, mode, route] == 0 if mode == 1 and route in conjunto_rutas_directas else x[
                                                                                                                  product, mode, route] ==
                                                                                                              x[
                                                                                                                  product, mode, route],
                          name="Direct")

    # lo que sale del centro debe ser menor que la capacidad de distribucion
    for CoD in Centro_distribucion:
        m.addCons(quicksum(
            x[product, mode, route] for product in lista_productos_demandados for mode in Modo for route in
            lista_rutas_ocupadas if route in conjunto_rutas_intermedias[CoD]) <= cd[CoD], name="CD(%s)" % CoD)

    # lo que sale de las plantas debe ser menor o igual a la capacidad de produccion para cada planta
    for plant in Plantas:
        m.addCons(quicksum(x[product, mode, route] for product in lista_productos_demandados for mode in Modo for route in lista_rutas_ocupadas if route in conjunto_rutas_plantas[plant]) <= Capp[plant],
                  name="CapP(%s)" % plant)

    # funcion objetivo
    m.addCons((quicksum(
        p_venta[product] * x[product, mode, route] for product in lista_productos_demandados for mode in Modo for route
        in lista_rutas_ocupadas)) * 0.95 - quicksum(
        c_produccion[product] * x[product, mode, route] for product in lista_productos_demandados for mode in
        Modo for route in lista_rutas_ocupadas) - quicksum(
        c_rutas[mode, route] * x[product, mode, route] for product in lista_productos_demandados for mode in Modo for
        route in lista_rutas_ocupadas) - quicksum(
        c_manejo[CoD] * x[product, mode, route] for CoD in Centro_distribucion for product in lista_productos_demandados
        for mode in Modo for route in conjunto_rutas_intermedias[CoD]) == s)


    m.setObjective(s, "maximize")
    m.optimize()
    m.printBestSol()
    print("El optimo es: ", m.getVal(s))

    # ahora se buscara exportar los datos en excel para luego ser analizado y cargados en power bi

    import pandas as pd
    from pandas import ExcelWriter


    archivo = []

    with open(str(ruta_archivo) + 'routes_available.csv') as File:
        reader = csv.reader(File, delimiter=';')
        for row in reader:
            archivo.append(row)
    archivo.pop(0)


    lista_resultados_diccionario={}

    for j in range(len(Productos)):
        for k in range(len(archivo)):
            for l in range(len(diccionario_de_rutas)):
                if archivo[k][0] in diccionario_de_rutas[l+1] and archivo[k][4] == '0':
                    lista_resultados_diccionario[j+1, archivo[k][0]] = m.getVal(x[j+1, 2, l+1])
                elif archivo[k][0] in diccionario_de_rutas[l+1] and archivo[k][4] == '1':
                    lista_resultados_diccionario[j + 1, archivo[k][0]] = m.getVal(x[j+1, 1, l+1])


    diccionario_resultados_finales={}
    for (i, j) in lista_resultados_diccionario:
        if lista_resultados_diccionario[i, j] != 0:
            diccionario_resultados_finales[i, j] = lista_resultados_diccionario[i, j]


    data1 = {'producto_id': [], 'ruta': [], 'cantidad transportada': []}

    for (i, j) in diccionario_resultados_finales:
        data1['producto_id'].append(int(i))
        data1['ruta'].append(j)
        data1['cantidad transportada'].append(diccionario_resultados_finales[i, j])


    df1 = pd.DataFrame(data1, columns=['producto_id', 'ruta', 'cantidad transportada'])


    writer = pd.ExcelWriter(str(ruta_archivo) + 'resultados.xlsx', engine='xlsxwriter')
    df1.to_excel(writer, sheet_name='Transporte')

    writer.save()
