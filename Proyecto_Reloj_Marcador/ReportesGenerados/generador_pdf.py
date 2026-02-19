from fpdf import FPDF
import os
from datetime import datetime
from tkinter import messagebox, filedialog
from pathlib import Path
from MensajeCargando import MensajeCargando

class PDFAsistencia(FPDF):
    def header(self):
        # Encabezado institucional con el logo del MEP
        ruta_encabezado = os.path.join(os.path.dirname(__file__), "encabezado_mep.png")
        if os.path.exists(ruta_encabezado):
            self.image(ruta_encabezado, x=10, y=8, w=190)
        self.ln(28)  # espacio después del encabezado

    def footer(self):
        # Pie de página con fecha y hora de generación
        self.set_y(-15)
        self.set_font("Arial", "I", 8)
        self.cell(0, 10, limpiar(f"Generado automáticamente el {datetime.now().strftime('%d/%m/%Y %H:%M')}"), 0, 0, "C")

#Genera el PDF individual con el formato oficial del MEP.
def generar_pdf_profesor(nombre, df_global, rutaGuardar):
    df_prof = df_global[df_global["Nombre"] == nombre]
    if df_prof.empty:
        return None
    # ===== Datos generales =====
    departamento = str(df_prof["Departamento"].iloc[0])
    id_prof = str(df_prof["ID"].iloc[0])
    periodo = f"{df_prof['Fecha'].min()} ~ {df_prof['Fecha'].max()}"
    #=== Creación del pdf====
    pdf=crear_pdf(departamento,id_prof,periodo,nombre)
    # ===== Tabla de asistencia =====
    diseño_tabla_asistencia(pdf, df_prof)
    return guardar_pdf(pdf,nombre,rutaGuardar)
    
def crear_pdf(departamento, id_prof, periodo, nombre):
    total_faltas = 0
    total_permiso = 0
    total_entrada = 0
    total_salida = 0
    total_extra = "0:00"
    total_retardos = 0
    total_temprano = 0
    pdf = PDFAsistencia()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    # ===== Título principal =====
    diseño_titulo_principal(pdf)
    # ===== Bloque de datos generales =====
    diseño_bloque_datos_generales(pdf, departamento, nombre, id_prof, periodo)
    # ===== Resumen General =====
    diseño_resumen_general(pdf,total_faltas,total_permiso, total_entrada,
                           total_salida, total_extra, total_retardos,
                             total_temprano)
    return pdf


def guardar_pdf(pdf, nombre,rutaGuardar):
    fecha_actual = datetime.now().strftime("%Y-%m-%d")

    
    carpeta_fecha = rutaGuardar / fecha_actual

    carpeta_fecha.mkdir(parents=True, exist_ok=True)

    ruta_pdf = carpeta_fecha / f"{nombre.replace(' ', '_')}.pdf"

    pdf.output(str(ruta_pdf))

    print(f"PDF generado: {ruta_pdf}")
    return ruta_pdf

#Genera todos los PDFs agrupados por fecha actual.
def generar_todos_los_pdfs(df_global,rutaGuardar):
    fecha_actual = datetime.now().strftime("%Y-%m-%d")
    carpeta_fecha = rutaGuardar / fecha_actual
    carpeta_fecha.mkdir(parents=True, exist_ok=True)
    
    nombres = df_global["Nombre"].unique().tolist()

    for nombre in nombres:
        generar_pdf_profesor(nombre, df_global,rutaGuardar)

    
def exportar_todos_pdf(app):
    ventanaMsj= MensajeCargando() 
    try:
        rutaGuardar=seleccionar_ruta_guardado()
        if not rutaGuardar:
            return
        ventanaMsj.deiconify()
        ventanaMsj.lift()
        ventanaMsj.attributes("-topmost", True) 
        if app.df_global.empty:
            ventanaMsj.withdraw()
            messagebox.showwarning("Aviso", "No hay datos cargados para exportar.")
            return
        generar_todos_los_pdfs(app.df_global,rutaGuardar) 
        ventanaMsj.withdraw() 
        messagebox.showinfo("Éxito", "Se generaron todos los reportes PDF correctamente.") 
    except Exception as e:
        ventanaMsj.withdraw()
        messagebox.showerror("Error", f"No se pudieron generar los PDF.\n{e}")

def seleccionar_ruta_guardado():
    rutaGuardar= filedialog.askdirectory(title="Guardar archivo en")
    if not rutaGuardar:
        return
    return Path(rutaGuardar)

#-------------------------------------------
#       MÉTODOS DE DISEÑO PARA EL PDF
#-------------------------------------------
def diseño_titulo_principal(pdf):
    pdf.set_font("Arial", "B", 14)
    pdf.set_text_color(0, 70, 130)  # azul institucional
    pdf.cell(0, 10, "Reporte General de Asistencia", 0, 1, "C")
    pdf.ln(4)


def diseño_bloque_datos_generales(pdf, departamento, nombre, id_prof, periodo):
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Arial", "", 11)
    pdf.cell(40, 8, limpiar(f"Depto.: {departamento}"), 0, 0)
    pdf.cell(70, 8, limpiar(f"Nombre: {nombre}"), 0, 0)
    pdf.cell(40, 8, limpiar(f"ID: {id_prof}"), 0, 1)
    pdf.cell(40, 8, limpiar(f"Periodo: {periodo}"), 0, 1)
    pdf.ln(3)
    
def diseño_resumen_general(pdf, total_faltas,total_permiso, total_entrada,
                           total_salida, total_extra, total_retardos,
                             total_temprano):
    pdf.set_font("Arial", "B", 10)
    pdf.cell(190, 8, "Resumen General", 0, 1, "C")
    pdf.set_font("Arial", "", 9)
    columnas = ["Falta", "Permiso", "Entrada", "Salida", "Tiempo Extra", "Retardos", "Salida Temprano"]
    valores = [
        total_faltas, total_permiso, total_entrada,
        total_salida, total_extra, total_retardos, total_temprano
    ]
    col_width = 190 / len(columnas)
    for col in columnas:
        pdf.cell(col_width, 8, limpiar(col), 1, 0, "C")
    pdf.ln()
    for val in valores:
        pdf.cell(col_width, 8, limpiar(str(val)), 1, 0, "C")
    pdf.ln(12)

def diseño_tabla_asistencia(pdf, df_prof):
    pdf.set_font("Arial", "B", 11)
    pdf.cell(0, 8, "Reporte de Asistencia", 0, 1, "C")
    pdf.ln(3)
    # Colores institucionales
    pdf.set_draw_color(0, 45, 90)   # borde azul oscuro
    pdf.set_text_color(0, 80, 0)    # texto verde oscuro
    pdf.set_fill_color(230, 230, 230)
    # Fila superior
    pdf.set_font("Arial", "B", 10)
    pdf.cell(38, 10, limpiar("Fecha Semana"), 1, 0, "C", False)
    pdf.cell(76, 10, limpiar("Primer Horario"), 1, 0, "C", False)
    pdf.cell(76, 10, limpiar("Segundo Horario"), 1, 1, "C", False)

    pdf.set_font("Arial", "", 9)
    pdf.cell(38, 8, "", 1, 0, "C", False)
    pdf.cell(38, 8, limpiar("Entrada"), 1, 0, "C", False)
    pdf.cell(38, 8, limpiar("Salida"), 1, 0, "C", False)
    pdf.cell(38, 8, limpiar("Entrada"), 1, 0, "C", False)
    pdf.cell(38, 8, limpiar("Salida"), 1, 1, "C", False)
    # ===== Cuerpo de tabla =====
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Arial", "", 10)
    for _, row in df_prof.iterrows():
        pdf.cell(38, 8, limpiar(str(row.get("Fecha", ""))), 1, 0, "C")
        pdf.cell(38, 8, limpiar(str(row.get("Entrada", ""))), 1, 0, "C")
        pdf.cell(38, 8, limpiar(str(row.get("Salida", ""))), 1, 0, "C")
        pdf.cell(38, 8, "", 1, 0, "C")  # Segundo horario entrada vacío
        pdf.cell(38, 8, "", 1, 1, "C")  # Segundo horario salida vacío

def limpiar(texto):
    return texto.encode("latin-1", "ignore").decode("latin-1")
