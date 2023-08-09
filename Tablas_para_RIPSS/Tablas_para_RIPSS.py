# Aplicación para la generación de tablas para el Informe de
# análisis de oferta de las RIPPS
#
import tkinter as tk
from tkinter import PhotoImage
from tkinter import ttk
from tkinter import filedialog

import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font
from openpyxl.worksheet.table import Table, TableStyleInfo

import os
import sys

# variables globales
version = "1.0"
archivo_configuracion = "config_tablas.txt"
def carga_libros(sel):
    global sheet_ripss, sheet_prestadores
    
    # Cargar los libros de trabajo
    if sel == 1:
        global archivo_RIPSS, workbook_ripss, archivo_destino
        workbook_ripss = load_workbook(filename=archivo_destino, read_only=True)  
        sheet_name_ripss = "BD_Depurada"
        sheet_ripss = workbook_ripss[sheet_name_ripss]
    if sel == 2: 
        global archivo_prestadores, archivo_origen
        workbook_prestadores = load_workbook(filename=archivo_origen, read_only=True)
        sheet_name_prestadores = "Prestadores (3)"    
        sheet_prestadores = workbook_prestadores[sheet_name_prestadores]

def selecciona_carpeta(sel):
    global archivo_prestadores, archivo_RIPSS
    global archivo_destino, archivo_origen
    ruta_seleccionada = os.path.join(os.environ['USERPROFILE'], 'Documents')
    
    if ruta_seleccionada:
        if sel == 1:
            ruta_seleccionada = recupera_parametro("archivo_destino")
            directorio, nombre_archivo = os.path.split(ruta_seleccionada)
            ruta_seleccionada = filedialog.askopenfilename(title="Ubicación de la fuente de RIPSS", initialdir=directorio,filetypes=(("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")) )
            entry.delete(0, tk.END)  # Limpiar el contenido actual del Entry
            archivo_destino = ruta_seleccionada
            entry.insert(0, ruta_seleccionada)
            archivo_RIPSS = ruta_seleccionada
        if sel == 2:
            ruta_seleccionada = recupera_parametro("archivo_origen")
            directorio, nombre_archivo = os.path.split(ruta_seleccionada)
            ruta_seleccionada = filedialog.askopenfilename(title="Ubicación de la fuente de RIPSS", initialdir=directorio,filetypes=(("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")) )
            entry2.delete(0, tk.END)  # Limpiar el contenido actual del Entry
            archivo_origen = ruta_seleccionada
            #ruta_seleccionada = recupera_parametro("archivo_origen")
            print(ruta_seleccionada)
            entry2.insert(0, ruta_seleccionada)
            archivo_prestadores = ruta_seleccionada
    
def actualiza_ripss():
    global sheet_ripss, sheet_prestadores, workbook_ripss
    global archivo_origen, archivo_destino
    asigna_parametro("archivo_origen")
    asigna_parametro("archivo_destino")
    carga_libros(1) # Excel de destino
    carga_libros(2) # Excel de origen de datos
    # Crea nueva hoja con permisos de escritura y copia sheet_ripss
    workbook_writable = Workbook()  # Aquí se crea el archivo 
    sheet_writable = workbook_writable.active
    # Copiar los datos de la hoja sheet_ripss de solo lectura a la hoja en el archivo editable
    for row in sheet_ripss.iter_rows(values_only=True):
        sheet_writable.append(row)
    
    # Obtener índices de columnas basados en encabezados
    col_index_ripss_codigo_prestador = 1
    col_index_prestadores_codigo_habilitacion = 2 
    
    # Crear un diccionario para almacenar los valores deseados
    valores_deseados = {}
    modificar = []
    # Iterar a través de las filas en archivo_prestadores y almacenar los valores en el diccionario
    for row_prestadores in sheet_prestadores.iter_rows(min_row=2, values_only=True):
        codigo_habilitacion_prestadores = row_prestadores[col_index_prestadores_codigo_habilitacion]
        valor_deseado = [row_prestadores[30],row_prestadores[31],row_prestadores[32],
                         row_prestadores[33],row_prestadores[34],row_prestadores[35],
                         row_prestadores[36],row_prestadores[37],row_prestadores[38]
                         ]
        valores_deseados[str(codigo_habilitacion_prestadores)] = valor_deseado
    
    # Iterar a través de las filas en archivo_RIPSS y actualizar las celdas
    for contador_fila_ripss, row_ripss in enumerate(sheet_writable.iter_rows(min_row=2, values_only=True), start=2):
        codigo_prestador_ripss = row_ripss[col_index_ripss_codigo_prestador]
        valor_deseado = valores_deseados.get(str(codigo_prestador_ripss))
        #print(contador_fila_ripss)
        if valor_deseado is not None:
            sheet_writable.cell(row=contador_fila_ripss, column=22).value = valor_deseado[0]
            sheet_writable.cell(row=contador_fila_ripss, column=23).value = valor_deseado[1]
            sheet_writable.cell(row=contador_fila_ripss, column=24).value = valor_deseado[2]
            sheet_writable.cell(row=contador_fila_ripss, column=25).value = valor_deseado[3]
            sheet_writable.cell(row=contador_fila_ripss, column=26).value = valor_deseado[4]
            sheet_writable.cell(row=contador_fila_ripss, column=27).value = valor_deseado[5]
            sheet_writable.cell(row=contador_fila_ripss, column=28).value = valor_deseado[6]
            sheet_writable.cell(row=contador_fila_ripss, column=29).value = valor_deseado[7]
            sheet_writable.cell(row=contador_fila_ripss, column=30).value = valor_deseado[8]

    # Guardar los cambios en archivo_RIPSS
    workbook_writable.save('archivo_RIPSS_actualizado.xlsx')


def crea_tabla_1():
    global sheet_ripss, archivo_RIPSS
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
    

def dibuja_tabla(tabla, nombre, wb, sheet):

    for row in tabla:
        sheet.append(row)
    # Crear una tabla a partir de los datos
    tab = Table(displayName=nombre, ref=sheet.dimensions)

    # Agregar un estilo a la tabla
    style = TableStyleInfo(
        name="TableStyleMedium9", showFirstColumn=False,
        showLastColumn=False, showRowStripes=True, showColumnStripes=True
    )
    tab.tableStyleInfo = style
    sheet.append([nombre])
    # Agregar la tabla a la hoja de cálculo
    sheet.add_table(tab)



def crea_hoja_Distrital():
    nombre_libro = "Tablas para RIPSS"
    nombre_hoja = "1.Distritales"

    hint.config(text="En ejecución ...")
    root.update()
    # Crear un nuevo libro de Excel y obtener la hoja activa
    wb = Workbook()
    sheet = wb.active
    sheet.title = nombre_hoja
    # Tabla 1
    tabla = crea_tabla_1()
    dibuja_tabla(tabla,"Tabla_1", wb, sheet)
    
    #Tabla 2
    tabla = crea_tabla_1()
    dibuja_tabla(tabla,"Tabla_2", wb, sheet)

    # Guardar el libro de Excel
    wb.save(nombre_libro+".xlsx")
    hint.config(text="Se ha creado el archivo "+nombre_libro+".xlsx")
    root.update()

# Funciones para consultar y almacenar parámetros
def asigna_parametro(variable):
    global archivo_configuracion
    print("El archivo es: "+archivo_configuracion)
    try:
        # Intentamos abrir el archivo en modo lectura
        with open(archivo_configuracion, "r") as file:
            lines = file.readlines()
    except FileNotFoundError:
        # Si el archivo no existe, lo creamos y escribimos el parámetro con su valor
        with open(archivo_configuracion, "w") as file:
            file.write(f"{variable}={globals()[variable]}\n")
    else:
        # Buscamos la variable en las líneas del archivo
        for i, line in enumerate(lines):
            key, value = line.strip().split("=")
            if key == variable:
                # Cambiamos el valor de la variable en el archivo
                lines[i] = f"{variable}={globals()[variable]}\n"
                with open(archivo_configuracion, "w") as file:
                    file.writelines(lines)
                break
        else:
            # Si la variable no está en el archivo, la agregamos al final
            with open(archivo_configuracion, "a") as file:
                file.write(f"{variable}={globals()[variable]}\n")

def recupera_parametro(variable):
    global archivo_configuracion
    file_name = archivo_configuracion
    
    try:
        # Intentamos abrir el archivo en modo lectura
        with open(archivo_configuracion, "r") as file:
            lines = file.readlines()
    except FileNotFoundError:
        # Si el archivo no existe, devolvemos el valor original de la variable
        return globals()[variable]
    else:
        # Buscamos la variable en las líneas del archivo
        for line in lines:
            key, value = line.strip().split("=")
            if key == variable:
                return value
        # Si la variable no está en el archivo, devolvemos el valor original de la variable
        return globals()[variable]


# Programa principal
# Define directorios
script_dir = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
images_dir = os.path.join(script_dir, 'images')
# Configuración de la ventana principal
root = tk.Tk()
root.title("Tablas para informe RIPSS  Ver.("+version+")")
root.geometry("500x440")
logo_path = os.path.join(images_dir, 'Logo.ico')
root.iconbitmap(logo_path)
icon_path = os.path.join(images_dir, 'folder.png')
icon_image_folder = PhotoImage(file=icon_path)

# Selección de carpeta de origen de RIPSS
# Campo de captura de texto
label1 = tk.Label(root, text="  Ubicación de la fuente de RIPSS ")
label1.grid(row=1, column=1, pady=5)
entry = tk.Entry(root, width=40)
entry.grid(row=1, column=2, columnspan=2, pady=10,sticky="w")
archivo_destino = os.path.join(os.environ['USERPROFILE'], 'Documents')
archivo_destino = recupera_parametro("archivo_destino")
entry.insert(0, archivo_destino)
boton_sel_carpeta = tk.Button(root, image=icon_image_folder, command=lambda: selecciona_carpeta(1))
boton_sel_carpeta.grid(row=1, column=4,sticky="w")

label2 = tk.Label(root, text="  Ubicación de la fuente de Prestadores ")
label2.grid(row=2, column=1, pady=5)
entry2 = tk.Entry(root, width=40)
entry2.grid(row=2, column=2, columnspan=2, pady=5,sticky="w")
archivo_origen = os.path.join(os.environ['USERPROFILE'], 'Documents')
archivo_origen = recupera_parametro("archivo_origen")
entry2.insert(0, archivo_origen)
boton_sel_carpeta2 = tk.Button(root, image=icon_image_folder, command=lambda: selecciona_carpeta(2))
boton_sel_carpeta2.grid(row=2, column=4,sticky="w")

boton_actualizaRipss = tk.Button(root, text="Actualiza Georeferenciación en RIPSS", command=actualiza_ripss)
boton_actualizaRipss.grid(row=4, column=3,sticky="w")

boton_Hoja1 = tk.Button(root, text="Crea la Tabla 1", command=crea_hoja_Distrital)
boton_Hoja1.grid(row=5, column=3,sticky="w")

hint = tk.Label(root, text="Seleccione los archivos para trabajar")
hint.grid(column=0, columnspan=2, sticky="se", padx=10, pady=10)
# Mostrar la ventana
root.mainloop()