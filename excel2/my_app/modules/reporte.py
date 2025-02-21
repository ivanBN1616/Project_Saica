from openpyxl import Workbook
from openpyxl.styles import Border, Side, PatternFill, Alignment
from datetime import datetime
from modules.utils import formatear_fecha

def generar_reporte(archivos):
    """
    Procesa los datos de los archivos y genera un reporte en Excel.
    Devuelve la ruta del reporte generado.
    """
    # Aquí iría tu lógica de extracción y procesamiento de datos
    # Ejemplo simplificado:
    resultado = [
        ("Máquina", "Descripción", "Fecha", "Hora", "Duración", "OT Asignada", "Descripción OT", "Trabajos Realizados", "Trabajadores", "Fecha y Hora Comienzo", "Horas Trabajadas"),
        # Datos procesados...
    ]
    
    wb_nuevo = Workbook()
    ws = wb_nuevo.active

    # Definir estilos
    borde = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    relleno_verde = PatternFill("solid", fgColor="90FE93")
    alineacion_centro = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Escribir encabezados
    titulos = resultado[0]
    for col, titulo in enumerate(titulos, start=1):
        celda = ws.cell(row=1, column=col, value=titulo)
        celda.fill = relleno_verde
        celda.border = borde
        celda.alignment = alineacion_centro

    # Escribir datos (aquí deberías usar tus datos procesados)
    for fila in resultado[1:]:
        ws.append(fila)

    # Ruta de guardado (esto puede provenir de config o diálogo)
    ruta_salida = f"{datetime.now().strftime('%Y-%m-%d')}_Reporte.xlsx"
    wb_nuevo.save(ruta_salida)
    return ruta_salida
