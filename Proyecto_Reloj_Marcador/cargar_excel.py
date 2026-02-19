from tkinter import ttk, filedialog, messagebox
import pandas as pd
from procesador import leer_excel_global
from MensajeCargando import MensajeCargando


def mostrar_datos_en_tabla(app):
    for fila in app.tree.get_children():
        app.tree.delete(fila)
    for _, row in app.df_global.iterrows():
        app.tree.insert("", "end", values=list(row))

def cargar_excel(app):
    ventana = MensajeCargando()
    
    ruta = filedialog.askopenfilename(
        title="Seleccionar archivo Excel",
        filetypes=[("Archivos Excel", "*.xls;*.xlsx")]
    )
    if not ruta:
        ventana.withdraw()
        return
    
    try:
        ventana.deiconify()
        ventana.lift()
        ventana.attributes("-topmost", True)  
          
        app.df_global = leer_excel_global(ruta)
        if app.df_global.empty:
            ventana.withdraw()
            messagebox.showwarning("Atención", "No se encontraron registros válidos.")
            return
        mostrar_datos_en_tabla(app)
        ventana.withdraw()
        messagebox.showinfo("Éxito", "Archivo cargado correctamente.")
    except Exception as e:
        ventana.withdraw()
        messagebox.showerror("Error", f"No se pudo leer el archivo.\n{e}")
        