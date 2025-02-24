# gui/main_window.py
import customtkinter as ctk
from tkinter import filedialog, messagebox
from modules.archivo_gestion import seleccionar_archivo, borrar_archivos
from modules.reporte import generar_reporte
from modules.file_management import eliminar_archivos
import os
from PIL import ImageTk, Image

class MainWindow(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Procesador de Reportes Excel")
        self.geometry("600x300")
        
        # Configurar íconos y rutas de imágenes (pueden venir de config o un módulo de recursos)
        self.icono_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "icon", "paros.ico")
        try:
            icono = ImageTk.PhotoImage(Image.open(self.icono_path))
            self.iconphoto(False, icono)
        except Exception as e:
            print(f"Error al cargar el ícono: {e}")

        # Frame para selección de archivos
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
        
        # Frame para botones de acción
        self.frame_botones = ctk.CTkFrame(self)
        self.frame_botones.pack(pady=10)
        
        btn_generar = ctk.CTkButton(self.frame_botones, text="Generar Reporte", command=self.generar_reporte_callback)
        btn_generar.grid(row=0, column=0, padx=10)
        
        btn_borrar = ctk.CTkButton(self.frame_botones, text="Borrar Archivos", command=lambda: borrar_archivos(self.archivos, self.entradas))
        btn_borrar.grid(row=0, column=1, padx=10)
        
        btn_eliminar = ctk.CTkButton(self, text="Eliminar antiguos Excel", command=eliminar_archivos)
        btn_eliminar.pack(side="left", anchor="sw", padx=10, pady=10)
    
    def generar_reporte_callback(self):
        nombre_reporte, wb = generar_reporte(self.archivos)
        if not nombre_reporte:
            messagebox.showerror("Error", "No se generó el reporte.")
            return
        ruta_salida = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Archivos de Excel", "*.xlsx")], title="Guardar reporte como", initialfile=nombre_reporte)
        if not ruta_salida:
            return
        try:
            wb.save(ruta_salida)
            messagebox.showinfo("Reporte Generado", f"El reporte se ha guardado en:\n{ruta_salida}")
            # Aquí se puede hacer visible un botón para abrir la carpeta, etc.
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar el reporte: {str(e)}")
