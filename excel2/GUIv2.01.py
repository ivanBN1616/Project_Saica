import tkinter as tk
import os
from tkinter import filedialog, messagebox
from openpyxl import Workbook, load_workbook
from pathlib import Path
from datetime import datetime
from openpyxl.styles import PatternFill, Border, Side, Alignment

def seleccionar_archivo(entry, archivos, tipo):
    archivo = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xlsx")])
    if archivo:
        entry.delete(0, tk.END)
        entry.insert(0, archivo)
        archivos[tipo] = archivo

def extraer_datos(archivo, columnas):
    ruta = Path(archivo)
    if not ruta.exists():
        messagebox.showerror("Error", f"El archivo {archivo} no existe.")
        return []
    
    wb = load_workbook(archivo, data_only=True)
    hoja = wb.active  
    return [tuple(row[:columnas]) for row in hoja.iter_rows(min_row=2, values_only=True)]

def formatear_fecha(fecha):
    if isinstance(fecha, str) and ' ' in fecha:
        return fecha.split()[0]
    elif isinstance(fecha, datetime):
        return fecha.strftime('%Y-%m-%d')
    return fecha

def generar_reporte(archivos):
    if not all(archivos.values()):
        messagebox.showerror("Error", "Debe seleccionar todos los archivos antes de generar el reporte.")
        return
    
    datos_cim3 = extraer_datos(archivos['cim3'], 21)
    datos_cim4 = extraer_datos(archivos['cim4'], 12)
    datos_ots = extraer_datos(archivos['ots'], 12)
    datos_trabajo_real = extraer_datos(archivos['trabajo_real'], 12)
    
    averias = [(fila[20], fila[12], fila[0], fila[14], fila[2]) for fila in datos_cim3 if fila[12] and ('Avería' in fila[12] or 'Averia' in fila[12])]
    averias += [(fila[11], fila[5], fila[2], fila[9], fila[4]) for fila in datos_cim4 if fila[5] and ('Avería' in fila[5] or 'Averia' in fila[5])]

    resultado = []
    for averia in averias:
        ubicacion, descripcion, hora, fecha, duracion = averia
        ot_asignada = "-"
        desc_ot = "-"
        fecha_comienzo = "-"
        hora_comienzo = "-"
        horas_trabajadas = "-"
        
        for ot in datos_ots:
            if ot[1][:5] == hora[:5]:
                ot_asignada = ot[4]
                desc_ot = ot[11]
                break
        
        trabajos_realizados = next((trabajo[8] for trabajo in datos_trabajo_real if trabajo[2] == ot_asignada), "Sin información")
        trabajadores = next((trabajo[9] for trabajo in datos_trabajo_real if trabajo[2] == ot_asignada), "Sin trabajador asignado")
        horas_trabajadas = next((trabajo[6] for trabajo in datos_trabajo_real if trabajo[2] == ot_asignada), "-")
        
        # Obtener fecha y hora de comienzo
        fecha_comienzo = next((formatear_fecha(trabajo[4]) for trabajo in datos_trabajo_real if trabajo[2] == ot_asignada), "-")
        hora_comienzo = next((trabajo[5].strftime('%H:%M:%S') if isinstance(trabajo[5], datetime) else str(trabajo[5]) for trabajo in datos_trabajo_real if trabajo[2] == ot_asignada), "-")
        
        # Combinamos fecha y hora de comienzo
        if fecha_comienzo != "-" and hora_comienzo != "-":
            fecha_hora_comienzo = f"{fecha_comienzo} {hora_comienzo}"
        else:
            fecha_hora_comienzo = "-"
        
        resultado.append((ubicacion, descripcion, fecha, hora, duracion, ot_asignada, desc_ot, trabajos_realizados, trabajadores, fecha_hora_comienzo, horas_trabajadas))

    wb_nuevo = Workbook()
    ws = wb_nuevo.active
    
    borde = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    relleno_verde = PatternFill("solid", fgColor="90FE93")
    alineacion_centro = Alignment(horizontal="center", vertical="center", wrap_text=True)

    titulos = ["Máquina", "Descripción", "Fecha", "Hora", "Duración", "OT Asignada", "Descripción OT", "Trabajos Realizados", "Trabajadores", "Fecha y Hora Comienzo", "Horas Trabajadas"]
    for col, titulo in enumerate(titulos, start=1):
        celda = ws.cell(row=1, column=col, value=titulo)
        celda.fill = relleno_verde
        celda.border = borde
        celda.alignment = alineacion_centro

    ancho_columnas = [15, 25, 12, 10, 10, 15, 50, 50, 25, 25, 12]
    for col, ancho in enumerate(ancho_columnas, start=1):
        ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = ancho

    for fila_idx, registro in enumerate(resultado, start=2):
        for col_idx, valor in enumerate(registro, start=1):
            celda = ws.cell(row=fila_idx, column=col_idx, value=valor)
            celda.border = borde
            celda.alignment = alineacion_centro

    ruta_salida = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Archivos de Excel", "*.xlsx"), ("Todos los archivos", "*.*")], title="Guardar reporte como", initialfile=f"{datetime.now().strftime('%Y-%m-%d')}_MTO_Registre_diari_OT_paro.xlsx")
    if not ruta_salida:
        return

    wb_nuevo.save(ruta_salida)
    messagebox.showinfo("Reporte Generado", f"El reporte se ha guardado en:\n{ruta_salida}")

# Inicialización de la interfaz gráfica
root = tk.Tk()
root.title("Procesador de Reportes Excel")
root.geometry("500x300")

# Obtiene la ruta absoluta del archivo ico en la misma carpeta que el script
icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "paros.ico")

# Aplica el icono si el archivo existe
if os.path.exists(icon_path):
    root.iconbitmap(icon_path)
else:
    print("⚠️ Advertencia: El archivo de icono no se encontró:", icon_path)


archivos = {"cim3": "", "cim4": "", "ots": "", "trabajo_real": ""}

frame_archivos = tk.Frame(root)
frame_archivos.pack(pady=10)

for i, tipo in enumerate(archivos.keys()):
    lbl = tk.Label(frame_archivos, text=f"Archivo {tipo.upper()}:")
    lbl.grid(row=i, column=0, padx=5, pady=5)
    entry = tk.Entry(frame_archivos, width=40)
    entry.grid(row=i, column=1, padx=5, pady=5)
    btn = tk.Button(frame_archivos, text="Buscar", command=lambda t=tipo, e=entry: seleccionar_archivo(e, archivos, t))
    btn.grid(row=i, column=2, padx=5, pady=5)

btn_generar = tk.Button(root, text="Generar Reporte", command=lambda: generar_reporte(archivos))
btn_generar.pack(pady=20)

root.mainloop()
