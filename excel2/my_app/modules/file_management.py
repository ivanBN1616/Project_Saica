# modules/file_management.py
import os
from pathlib import Path
from tkinter import filedialog, messagebox

def eliminar_archivos():
    archivos_a_borrar = filedialog.askopenfilenames(title="Seleccionar archivos a eliminar",
                                                    filetypes=[("Archivos de Excel", "*.xlsx;*.xls")])
    if not archivos_a_borrar:
        return
    
    archivos_permitidos = ["CIM3", "CIM4", "OT", "REAL"]
    archivos_validos = [archivo for archivo in archivos_a_borrar if any(nombre in Path(archivo).stem.upper() for nombre in archivos_permitidos)]
    
    if not archivos_validos:
        messagebox.showwarning("Advertencia", "Solo puedes eliminar archivos con nombres: CIM3, CIM4, OT o REAL.")
        return
    
    confirmacion = messagebox.askyesno("Confirmación", f"¿Seguro que quieres eliminar estos archivos?\n\n" + "\n".join(archivos_validos))
    if confirmacion:
        errores = []
        for archivo in archivos_validos:
            try:
                os.remove(archivo)
            except Exception as e:
                errores.append(f"No se pudo eliminar {archivo}.\nError: {e}")
        if errores:
            messagebox.showerror("Errores al eliminar archivos", "\n\n".join(errores))
        else:
            messagebox.showinfo("Éxito", "Los archivos se eliminaron correctamente.")
