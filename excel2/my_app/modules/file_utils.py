# modules/file_utils.py
import os
import subprocess
from tkinter import messagebox

def abrir_ubicacion_archivo(ruta_archivo):
    carpeta = os.path.dirname(ruta_archivo)
    if os.path.exists(carpeta):
        try:
            if os.name == 'nt':
                subprocess.run(['explorer', '/select,', os.path.abspath(ruta_archivo)])
            elif os.name == 'posix':
                subprocess.run(['open', carpeta])
            else:
                subprocess.run(['xdg-open', carpeta])
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir la carpeta: {e}")
    else:
        messagebox.showerror("Error", f"La carpeta no existe: {carpeta}")
