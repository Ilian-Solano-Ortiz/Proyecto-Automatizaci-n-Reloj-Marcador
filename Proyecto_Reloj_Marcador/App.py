import tkinter as tk
import pandas as pd
from tkinter import ttk
from cargar_excel import cargar_excel, mostrar_datos_en_tabla
from ReportesGenerados.generador_pdf import exportar_todos_pdf


class App(tk.Tk):

    def __init__(self):
        super().__init__()
        self.df_global=pd.DataFrame()
        # ---------------------------------------
        # CONFIGURACIÓN DE LA VENTANA
        # ---------------------------------------
        self.title("Visor de Asistencias MEP")
        self.geometry("1300x700")
        self.configure(bg="#e6f0fa")

        # Construcción de la interfaz
        self.crear_encabezado()
        self.crear_controles()
        self.crear_tabla()

    # =========================================================
    #  Encabezado institucional
    # =========================================================
    def crear_encabezado(self):
        frame_titulo = tk.Frame(self, bg="#002d5a", height=60)
        frame_titulo.pack(fill="x")

        lbl_titulo = tk.Label(
            frame_titulo,
            text="📘 VISOR DE ASISTENCIAS - MINISTERIO DE EDUCACIÓN PÚBLICA",
            bg="#002d5a",
            fg="white",
            font=("Arial", 14, "bold")
        )
        lbl_titulo.pack(pady=10)

    # =========================================================
    # Botones superiores
    # =========================================================
    def crear_controles(self):
        frame_top = tk.Frame(self, bg="#e6f0fa")
        frame_top.pack(pady=10)

        estilo_boton = {
            "font": ("Arial", 11, "bold"),
            "bg": "#002d5a",
            "fg": "white",
            "activebackground": "#f2c94c",
            "activeforeground": "black",
            "bd": 0,
            "width": 18,
            "height": 2
        }

        btn_cargar = tk.Button(
            frame_top,
            text="Cargar Excel",
            command=lambda: cargar_excel(self),
            **estilo_boton
        )
        btn_cargar.pack(side="left", padx=10)

        btn_pdf = tk.Button(
            frame_top,
            text="Exportar todos los PDF",
            command=lambda: exportar_todos_pdf(self),
            **estilo_boton
        )
        btn_pdf.pack(side="left", padx=10)

    # =========================================================
    # Tabla TreeView
    # =========================================================
    def crear_tabla(self):
        cols = ["ID", "Nombre", "Departamento", "Fecha", "Entrada", "Salida"]
        self.tree = ttk.Treeview(self, columns=cols, show="headings")

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview.Heading",
                        background="#002d5a",
                        foreground="white",
                        font=("Arial", 10, "bold"))
        style.configure("Treeview",
                        background="white",
                        foreground="black",
                        rowheight=25,
                        fieldbackground="white",
                        font=("Arial", 10))
        style.map("Treeview", background=[("selected", "#f2c94c")])

        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=200, anchor="center")

        scroll_y = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_y.set)
        scroll_y.pack(side="right", fill="y")

        self.tree.pack(fill="both", expand=True, padx=15, pady=10)
