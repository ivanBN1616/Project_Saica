import os
import pandas as pd

def convertir_xls_a_xlsx(ruta_xls, ruta_xlsx=None):
    """
    Convierte un archivo Excel de formato XLS a XLSX.
    Si ruta_xlsx no se especifica, se crea en el mismo directorio con extensión .xlsx.
    """
    if ruta_xlsx is None:
        base, _ = os.path.splitext(ruta_xls)
        ruta_xlsx = base + ".xlsx"
    try:
        excel = pd.read_excel(ruta_xls, sheet_name=None, engine='xlrd')
        with pd.ExcelWriter(ruta_xlsx, engine='openpyxl') as writer:
            for hoja, df in excel.items():
                df.to_excel(writer, sheet_name=hoja, index=False)
        return ruta_xlsx
    except Exception as e:
        print(f"Error al convertir {ruta_xls}: {e}")
        return None
