import tkinter as tk
from Tablas_para_RIPSS import carga_libros
def crea_tabla_1():
    hint.config(text="Elaborando Tabla 1 ...")
    root.update()
    global sheet_ripss  #, archivo_RIPSS
    carga_libros(1)
    tipo_prestador = []
    prestadores  = []
    no_prestadores = []
    eps = []
    for row in sheet_ripss.iter_rows(min_row=2, values_only=True):
        tipo_prestador = str(row[15])
        eps.append(row[0])
        if tipo_prestador == "Instituciones Prestadoras de Servicios de Salud - IPS":
            prestadores.append(str(row[1]))
        else:
            no_prestadores.append(str(row[1]))
    # Resultados
    total_prestadores = len(set(prestadores))
    total_otros = len(set(no_prestadores))
    servicios_no_prestadores = len(no_prestadores)
    servicios_prestadores = len(prestadores)
    # Dibuja la tabla
    data = [
    ["PROVEEDORES",	"CANTIDAD",	"SERVICIOS CONTRATADOS"],
    ["Prestadores de servicios de salud",total_prestadores,servicios_prestadores],
    ["Otros Proveedores",total_otros, servicios_no_prestadores],
    ["Total", total_prestadores + total_otros, servicios_prestadores + servicios_no_prestadores]
    ]
    return data

def crea_tabla_2():
    hint.config(text="Elaborando Tabla 2 ...")
    root.update()
    global sheet_ripss  #, archivo_RIPSS
    carga_libros(1)

    # Obtiene información
    prestadores = []
    for row in sheet_ripss.iter_rows(min_row=2, values_only=True):
        tipo_prestador = str(row[15])
        codigo = str(row[1])
        prestadores.append((codigo, tipo_prestador))
    total_IPS = len(set((x,y) for x,y in prestadores if y==str("Instituciones Prestadoras de Servicios de Salud - IPS")))
    total_diferente = len(set((x,y) for x,y in prestadores if y==str("Objeto Social Diferente a la Prestación de Servicios de Salud")))
    total_otros = len(set((x,y) for x,y in prestadores if y==str("Otros prestadores de servicios")))
    total_profesional = len(set((x,y) for x,y in prestadores if y==str("Profesional Independiente")))
    total_transporte = len(set((x,y) for x,y in prestadores if y==str("Transporte Especial de Pacientes")))
    total = total_IPS + total_diferente + total_otros + total_profesional + total_transporte
    # Crea la matriz (tabla)
    titulos = ["TIPO PRESTADOR","PRESTADORES","%"]
    data = [
            titulos,
            ["IPS",total_IPS,round(total_IPS*100/total,1) ],
            ["OBJETO SOCIAL DIFERENTE",total_diferente,round(total_diferente*100/total,1)],
            ["OTROS PRESTADORES DE SERVICIO",total_otros,round(total_otros*100/total,1)],
            ["PROFESIONAL INDEPENDIENTE",total_profesional,round(total_profesional*100/total,1)],
            ["TRANSPORTE ESPECIAL DE PACIENTES",total_transporte,round(total_transporte*100/total,1)],
            ["TOTAL",total,100]
        ]
    return data

def crea_tabla_3():
    hint.config(text="Elaborando Tabla 3 ...")
    root.update()
    global sheet_ripss  #, archivo_RIPSS
    carga_libros(1)

    # Obtiene información
    prestadores = []
    for row in sheet_ripss.iter_rows(min_row=2, values_only=True):
        if str(row[15]) == "Instituciones Prestadoras de Servicios de Salud - IPS":
            eps = str(row[0])
            codigo = str(row[1])
            prestadores.append((eps,codigo))
    Eps = set(x for x,y in prestadores)
    total_eps = len(Eps)
    
    # Crea la matriz (tabla)
    data = []
    titulos = ["EAPB",	"NÚMERO IPS","AFILIADOS EAPB","PROPORCIÓN por cada 10.000 afiliados"]
    data.append(titulos)
    for nombre_eps in sorted(list(Eps)):
        total_ips = len(set( y for x,y in prestadores if x == nombre_eps))
        data.append([nombre_eps,total_ips,"",""])
    return data

def crea_tabla_4():
    hint.config(text="Elaborando Tabla 4 ...")
    root.update()
    global sheet_ripss  # archivo_RIPSS
    global sheet_prestadores # archivo con prestadores
    carga_libros(1)
    carga_libros(2)
    # Crear un diccionario codigo_ips: array geolocalización
    diccionario_ips = {}
        # Iterar a través de las filas en archivo_prestadores y almacenar los valores en el diccionario
    for row_prestadores in sheet_prestadores.iter_rows(min_row=2, values_only=True):
        codigo_habilitacion_prestadores = str(row_prestadores[2])
        valor_deseado = row_prestadores[0:39]
        diccionario_ips[str(codigo_habilitacion_prestadores)] = valor_deseado
    # Obtiene información
    prestadores = []
    for row in sheet_ripss.iter_rows(min_row=2, values_only=True):
            codigo = str(row[1])
            nom_loc = diccionario_ips.get(codigo, False)
            if nom_loc:
                nombre_localidad = nom_loc[32]
            else:
                nombre_localidad = "Sin localizar"
            prestadores.append((codigo,nombre_localidad))
    localidades = sorted(set(y for x,y in prestadores))

    ips_loc = set(prestadores)
    # Crea la matriz (tabla)
    data = []
    titulos = ["LOCALIDAD", "PRESTADORES",	"%"]
    data.append(titulos)
    tot = len(ips_loc)
    for localidad in localidades:
        total_ips = len(set( x for x,y in ips_loc if y == localidad))
        data.append([localidad,total_ips,round((total_ips/tot)*100,2)])
    data.append(["TOTAL",tot,100])
    return data

def crea_tabla_5():
    hint.config(text="Elaborando Tabla 5 ...")
    root.update()
    global sheet_ripss  # archivo_RIPSS
    # global sheet_prestadores # archivo con prestadores
    carga_libros(1)
    #carga_libros(2)

    # Obtiene información
    grupo_servicio = []
    for row in sheet_ripss.iter_rows(min_row=2, values_only=True): 
        grupo_servicio.append(str(row[16]))
    servicios = set(grupo_servicio)
    
    # Crea la matriz (tabla)
    data = []
    titulos = ["GRUPO SERVICIOS","TOTAL","%"]
    data.append(titulos)
    tot = len(grupo_servicio)
    for grupo in servicios:
        total_grupo = grupo_servicio.count(grupo)
        data.append([grupo,total_grupo,round((total_grupo/tot)*100,2)])
    data.append(["TOTAL",tot,100])
    return data

def crea_tabla_6():
    hint.config(text="Elaborando Tabla 6 ...")
    root.update()
    global sheet_ripss  # archivo_RIPSS
    # global sheet_prestadores # archivo con prestadores
    carga_libros(1)
    #carga_libros(2)

    # Obtiene información
    red_general = []
    for row in sheet_ripss.iter_rows(min_row=2, values_only=True): 
        eps = str(row[0])
        if row[7] == None:
            red_general.append( (eps, "NONE") )
        else:
            red_general.append( (eps, str(row[7])) )
    nombres_eps = set(x.strip().upper() for x,y in red_general)
    tipos = list(set(y.strip().upper() for x,y in red_general))
    # Crea la matriz (tabla)
    data = []
    titulo1 = ["EAPB"] + [x for x in tipos] + ["TOTAL"]
    data.append(titulo1)
    total = {}

    for nombre in sorted(list(nombres_eps)):
        for tipo in tipos:
            total[tipo] = red_general.count((nombre,tipo))
        tot = sum(total.values())
        data.append([nombre] +[total[x] for x in tipos] + [tot])
    
    suma_columnas = [0] * len(data[0])
    for fila in data[1:]:
        for i in range(1,len(fila)):
            suma_columnas[i] += fila[i]
    
    data.append(["TOTAL"] + suma_columnas[1:])
    return data

def crea_tabla_7():
    hint.config(text="Elaborando Tabla 7 ...")
    root.update()
    global sheet_ripss  # archivo_RIPSS
    # global sheet_prestadores # archivo con prestadores
    carga_libros(1)
    #carga_libros(2)

    # Obtiene información
    red_general = []
    for row in sheet_ripss.iter_rows(min_row=2, values_only=True): 
        eps = str(row[0])
        if row[8] == None:
            red_general.append( (eps, "NONE") )
        else:
            red_general.append( (eps, str(row[8]) ))
    nombres_eps = set(x.strip().upper() for x,y in red_general)
    tipos = list(set(y.strip().upper() for x,y in red_general))
    print(tipos)
    # Crea la matriz (tabla)
    data = []
    titulo1 = ["EAPB"] + [x.upper() for x in tipos] + ["TOTAL"]
    data.append(titulo1)
    total = {}
    for nombre in sorted(list(nombres_eps)):
        for tipo in tipos:
            total[tipo] = red_general.count((nombre,tipo))
        tot = sum(total.values())
        data.append([nombre] +[total[x] for x in tipos] + [tot])
    
    suma_columnas = [0] * len(data[0])
    for fila in data[1:]:
        for i in range(1,len(fila)):
            suma_columnas[i] += fila[i]
    
    data.append(["TOTAL"] + suma_columnas[1:])
    return data

def crea_tabla_8():
    hint.config(text="Elaborando Tabla 8 ...")
    root.update()
    global sheet_ripss  # archivo_RIPSS
    # global sheet_prestadores # archivo con prestadores
    carga_libros(1)
    #carga_libros(2)

    # Obtiene información
    red_general = []
    for row in sheet_ripss.iter_rows(min_row=2, values_only=True): 
        eps = str(row[0])
        if row[9] == None:
            red_general.append( (eps, "NONE") )
        else:
            red_general.append( (eps, str(row[9])) )
    nombres_eps = set(x.strip().upper() for x,y in red_general)
    tipos = list(set(y.strip().upper() for x,y in red_general))
    # Crea la matriz (tabla)
    data = []
    titulo1 = ["EAPB"] + [x for x in tipos] + ["TOTAL"]
    data.append(titulo1)
    total = {}
    for nombre in sorted(list(nombres_eps)):
        for tipo in tipos:
            total[tipo] = red_general.count((nombre,tipo))
        tot = sum(total.values())
        data.append([nombre] +[total[x] for x in tipos] + [tot])
    
    suma_columnas = [0] * len(data[0])
    for fila in data[1:]:
        for i in range(1,len(fila)):
            suma_columnas[i] += fila[i]
    
    data.append(["TOTAL"] + suma_columnas[1:])
    return data

def crea_tabla_9():
    hint.config(text="Elaborando Tabla 9 ...")
    root.update()
    global sheet_ripss  # archivo_RIPSS
    # global sheet_prestadores # archivo con prestadores
    carga_libros(1)
    #carga_libros(2)

    # Obtiene información
    red_general = []
    for row in sheet_ripss.iter_rows(min_row=2, values_only=True): 
        eps = str(row[0])
        if row[10] == None:
            red_general.append( (eps, "NONE") )
        else:
            red_general.append( (eps, str(row[10])) )
    nombres_eps = set(x.strip() for x,y in red_general)
    tipos = list(set(y.strip().upper() for x,y in red_general))
    # Crea la matriz (tabla)
    data = []
    titulo1 = ["EAPB"] + [x for x in tipos] + ["TOTAL"]
    data.append(titulo1)
    total = {}
    for nombre in sorted(list(nombres_eps)):
        for tipo in tipos:
            total[tipo] = red_general.count((nombre,tipo))
        tot = sum(total.values())
        data.append([nombre] +[total[x] for x in tipos] + [tot])
    
    suma_columnas = [0] * len(data[0])
    for fila in data[1:]:
        for i in range(1,len(fila)):
            suma_columnas[i] += fila[i]
    
    data.append(["TOTAL"] + suma_columnas[1:])
    return data
