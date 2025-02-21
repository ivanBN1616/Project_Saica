def formatear_fecha(fecha):
    """
    Convierte una fecha a formato 'YYYY-MM-DD'.
    """
    if isinstance(fecha, str) and ' ' in fecha:
        return fecha.split()[0]
    elif hasattr(fecha, "strftime"):
        return fecha.strftime('%Y-%m-%d')
    return fecha

def obtener_ruta_relativa(ruta_archivo):
    """
    Retorna la ruta del archivo relativa al directorio del script.
    """
    import os
    base_path = os.path.dirname(os.path.realpath(__file__))
    return os.path.join(base_path, ruta_archivo)
