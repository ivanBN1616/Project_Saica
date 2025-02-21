import customtkinter as ctk
from modules.archivo_gestion import seleccionar_archivo, borrar_archivos
from modules.reporte import generar_reporte

class MainWindow(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Procesador de Reportes Excel")
        self.geometry("600x300")
        
        # Crear el frame para selección de archivos
        self.frame_archivos = ctk.CTkFrame(self)
        self.frame_archivos.pack(pady=10)
        
        self.archivos = {"cim3": "", "cim4": "", "ots": "", "trabajo_real": ""}
        self.entradas = []
        
        for i, tipo in enumerate(self.archivos.keys()):
            lbl = ctk.CTkLabel(self.frame_archivos, text=f"{tipo.upper()}:")
            lbl.grid(row=i, column=0, padx=5, pady=5)
            entry = ctk.CTkEntry(self.frame_archivos, width=180)
            entry.grid(row=i, column=1, padx=5, pady=5)
            self.entradas.append(entry)
            btn = ctk.CTkButton(self.frame_archivos, text="", command=lambda t=tipo, e=entry: seleccionar_archivo(e, self.archivos, t))
            btn.grid(row=i, column=2, padx=5, pady=5)
        
        # Crear el frame para botones de acción
        self.frame_botones = ctk.CTkFrame(self)
        self.frame_botones.pack(pady=10)
        
        btn_generar = ctk.CTkButton(self.frame_botones, text="Generar Reporte", command=self.generar_reporte_callback)
        btn_generar.grid(row=0, column=0, padx=10)
        
        btn_borrar = ctk.CTkButton(self.frame_botones, text="Borrar Archivos", command=lambda: borrar_archivos(self.archivos, self.entradas))
        btn_borrar.grid(row=0, column=1, padx=10)

    def generar_reporte_callback(self):
        # Llama a la función de generación de reporte y muestra el resultado en un messagebox
        ruta_reporte = generar_reporte(self.archivos)
        if ruta_reporte:
            ctk.messagebox.showinfo("Reporte Generado", f"El reporte se ha guardado en:\n{ruta_reporte}")
