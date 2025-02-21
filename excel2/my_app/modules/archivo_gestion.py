import os
from tkinter import filedialog, messagebox
import customtkinter as ctk

def seleccionar_archivo(entry, archivos, tipo):
    """
    Permite seleccionar un archivo Excel, muestra solo el nombre en el campo de entrada,
    y guarda la ruta completa en el diccionario 'archivos'.
    """
    archivo = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xlsx;*.xls")])
    if archivo:
        # Aquí podrías agregar validaciones específicas (extensión, nombre, etc.)
        nombre_archivo = os.path.basename(archivo)
        entry.delete(0, ctk.END)
        entry.insert(0, nombre_archivo)
        archivos[tipo] = archivo

def borrar_archivos(archivos, entradas):
    """
    Limpia el diccionario de archivos y borra el contenido de las entradas.
    """
    for tipo in archivos:
        archivos[tipo] = ""
    for entry in entradas:
        entry.delete(0, ctk.END)

