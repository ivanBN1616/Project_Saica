# app.py
import customtkinter as ctk
from gui.main_window import MainWindow

def main():
    # Configuración de la apariencia
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")
    
    # Inicia la ventana principal
    app = MainWindow()
    app.mainloop()

if __name__ == "__main__":
    main()
