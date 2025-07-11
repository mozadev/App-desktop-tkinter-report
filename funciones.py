from tkinter import *
import pandas as pd
from datetime import datetime
import math
from tkinter import ttk, messagebox, filedialog, font  # Importa el módulo de barras de progreso

# FUNCIONES
def handle_nan(value):
    if pd.notna(value):
        return int(value)
    else:
        return 0

def handle_nan_vacio(value):
    if pd.notna(value):
        return str(value)
    else:
        return ''

# Obtener la fecha actual
def obtener_fecha_actual():
    return datetime.now()

# def formato_fecha(fecha_str, formato_entrada="%Y-%m-%d %H:%M:%S"):
#     # Verificar si el valor es NaN o cadena vacía
#     if isinstance(fecha_str, float) and math.isnan(fecha_str):
#         return "Fecha no disponible"
    
#     # Si la fecha es una cadena vacía, devolver un mensaje indicando que no hay fecha disponible
#     if not fecha_str:
#         return "Fecha no disponible"

#     # Si la fecha ya es una cadena, simplemente formatearla
#     if isinstance(fecha_str, str):
#         fecha_obj = datetime.strptime(fecha_str, formato_entrada)
#     # Si la fecha es un número (int o float), convertirlo a cadena y luego formatear
#     elif isinstance(fecha_str, (int, float)):
#         fecha_obj = datetime.fromtimestamp(fecha_str)
#     else:
#         # Manejar otros casos según sea necesario
#         return "Fecha no válida"

#     # Formatear la fecha en el nuevo formato
#     nueva_fecha_str = fecha_obj.strftime("%d/%m/%Y %H:%M")
    
#     return nueva_fecha_str

def transformar_fecha(fecha_str):
    # Verificar si el valor es NaN o cadena vacía
    if isinstance(fecha_str, float) and math.isnan(fecha_str):
        return "Fecha no disponible"
    
    # Si la fecha es una cadena vacía, devolver un mensaje indicando que no hay fecha disponible
    if not fecha_str:
        return "Fecha no disponible"
    
    try:
        # Intenta analizar la fecha en varios formatos
        fecha = datetime.strptime(fecha_str, "%Y-%m-%d %H:%M:%S")
    except ValueError:
        try:
            fecha = datetime.strptime(fecha_str, "%d/%m/%Y %H:%M:%S")
        except ValueError:
            try:
                fecha = datetime.strptime(fecha_str, "%d/%m/%Y %H:%M")
            except ValueError:
                try:
                    fecha = datetime.strptime(fecha_str, "%Y-%m-%d")
                except ValueError:
                    try:
                        fecha = datetime.strptime(fecha_str, "%d/%m/%Y")
                    except ValueError:
                        # Si no se puede analizar la fecha, lanza una excepción
                        raise ValueError("Formato de fecha no reconocido")

    # Formatea la fecha en el formato deseado
    fecha_formateada = fecha.strftime("%d/%m/%Y %H:%M")
    return fecha_formateada


def seleccionar_archivo(label, ruta_variable):
    archivo = filedialog.askopenfilename(title="Seleccionar archivo")
    if archivo:
        ruta_variable.set(archivo)
        label.config(text=f"Archivo: {archivo}")
        print("Ruta del archivo seleccionado:", archivo)