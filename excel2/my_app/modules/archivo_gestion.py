# modules/archivo_gestion.py
import os
from tkinter import filedialog, messagebox
import customtkinter as ctk

def seleccionar_archivo(entry, archivos, tipo):
    try:
        archivo = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xlsx;*.xls")])
        if archivo:
            # Conversión si es .xls a .xlsx (podrías importar la función de conversion aquí)
            if archivo.lower().endswith(".xls") and not archivo.lower().endswith(".xlsx"):
                respuesta = messagebox.askyesno("Conversión", "El archivo está en formato .xls. ¿Deseas convertirlo a .xlsx para continuar?")
                if respuesta:
                    from modules.conversion import convertir_xls_a_xlsx
                    nuevo_archivo = convertir_xls_a_xlsx(archivo)
                    if nuevo_archivo:
                        archivo = nuevo_archivo
                        messagebox.showinfo("Conversión exitosa", f"Archivo convertido a {nuevo_archivo}")
                    else:
                        messagebox.showerror("Error", "No se pudo convertir el archivo.")
                        return
                else:
                    return
            
            # Validación de extensión y nombre
            if not archivo.lower().endswith(".xlsx"):
                messagebox.showerror("Error", f"El archivo seleccionado para {tipo} no es válido.\nDebe ser un archivo Excel (.xlsx).")
                return

            nombre_archivo = os.path.basename(archivo)
            entry.delete(0, ctk.END)
            entry.insert(0, nombre_archivo)
            archivos[tipo] = archivo
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al seleccionar el archivo para {tipo}: {e}")

def borrar_archivos(archivos, entradas):
    for tipo in archivos.keys():
        archivos[tipo] = ""
    for entry in entradas:
        entry.delete(0, ctk.END)
    # Se asume que btn_abrir_carpeta es global o se maneja de otra forma en la UI
    try:
        from gui.main_window import ocultar_boton_ubicacion  # Por ejemplo, se implementa en el módulo de UI
        ocultar_boton_ubicacion()
    except Exception:
        pass
