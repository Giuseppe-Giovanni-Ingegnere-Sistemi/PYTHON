import pandas as pd
from docx import Document
from docx.shared import Pt
import os
from tkinter import Tk, filedialog

# 📌 Función para seleccionar archivos con un cuadro de diálogo
def seleccionar_archivo(tipo):
    root = Tk()
    root.withdraw()  # Oculta la ventana de Tkinter
    if tipo == "excel":
        archivo = filedialog.askopenfilename(title="Selecciona el archivo de Excel", filetypes=[("Archivos de Excel", "*.xlsx")])
    elif tipo == "word":
        archivo = filedialog.askopenfilename(title="Selecciona la plantilla de Word", filetypes=[("Archivos de Word", "*.docx")])
    else:
        archivo = filedialog.askdirectory(title="Selecciona la carpeta donde guardar los documentos")
    return archivo

# 📌 Pedir al usuario que seleccione los archivos
ruta_excel = seleccionar_archivo("excel")
ruta_plantilla = seleccionar_archivo("word")
ruta_guardar = seleccionar_archivo("carpeta")

# 📌 Si el usuario no selecciona nada, salir del programa
if not ruta_excel or not ruta_plantilla or not ruta_guardar:
    print("⚠️ No seleccionaste todos los archivos. Saliendo del programa...")
    exit()

# 📌 Leer el archivo Excel correctamente
df = pd.read_excel(ruta_excel, sheet_name="CALCULO", header=5)
df.columns = df.columns.str.strip()  # Eliminar espacios en los nombres de columnas

# 📌 Crear la carpeta de guardado si no existe
if not os.path.exists(ruta_guardar):
    os.makedirs(ruta_guardar)

# 📌 Función para formatear valores numéricos correctamente
def formatear_valor(valor):
    if pd.isna(valor):
        return "0.00"
    elif isinstance(valor, (int, float)):
        return f"{valor:,.2f}".replace(",", " ")
    return str(valor).strip()

# 📌 Función para reemplazar texto sin perder formato
def reemplazar_texto(doc, buscar, reemplazo):
    reemplazo = formatear_valor(reemplazo)
    for parrafo in doc.paragraphs:
        if buscar in parrafo.text:
            for run in parrafo.runs:
                if buscar in run.text:
                    run.text = run.text.replace(buscar, reemplazo)
                    run.bold = True
                    run.font.name = "Verdana"
                    run.font.size = Pt(11)
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for parrafo in celda.paragraphs:
                    if buscar in parrafo.text:
                        for run in parrafo.runs:
                            if buscar in run.text:
                                run.text = run.text.replace(buscar, reemplazo)
                                run.bold = True
                                run.font.name = "Verdana"
                                run.font.size = Pt(11)

# 📌 Iterar sobre todas las filas del Excel para generar un documento por cada persona
for index, fila in df.iterrows():
    if pd.isna(fila.get("Nombre completo", "")):
        continue

    # 📌 Abrir la plantilla de Word
    doc = Document(ruta_plantilla)

    # 📌 Reemplazar los marcadores en el documento
    reemplazar_texto(doc, "«Nombre_completo»", fila.get("Nombre completo", ""))
    reemplazar_texto(doc, "«Puesto»", fila.get("Puesto", ""))
    reemplazar_texto(doc, "«Salario_por_día»", fila.get("Salario por día", ""))
    reemplazar_texto(doc, "«Fecha_de_alta»", fila.get("Fecha de alta", ""))
    reemplazar_texto(doc, "«Fecha_de_baja»", fila.get("Fecha de baja", ""))
    reemplazar_texto(doc, "«SUELDO»", fila.get("SUELDO", ""))
    reemplazar_texto(doc, "«IMPORTE_AGUINALDO»", fila.get("IMPORTE AGUINALDO", ""))
    reemplazar_texto(doc, "«IMPORTE_VACACIONES»", fila.get("IMPORTE VACACIONES", ""))
    reemplazar_texto(doc, "«GRATIFICACION»", fila.get("GRATIFICACION", ""))
    reemplazar_texto(doc, "«IMPORTE_PRIMA_VACACIONAL»", fila.get("IMPORTE PRIMA VACACIONAL", ""))
    reemplazar_texto(doc, "«Total_de_Percepciones»", fila.get("TOTAL PERCEPCIONES", ""))
    reemplazar_texto(doc, "«TOTAL_ISR»", fila.get("ISR MENSUAL", ""))
    reemplazar_texto(doc, "«IMSS»", fila.get("IMSS", ""))
    reemplazar_texto(doc, "«TOTAL_DEDUCCIONES»", fila.get("TOTAL DEDUCCIONES", ""))
    reemplazar_texto(doc, "«NETO»", fila.get("NETO", ""))
    reemplazar_texto(doc, "«Banco»", fila.get("Banco", ""))
    reemplazar_texto(doc, "«CUENTA»", fila.get("cuenta", ""))

    # 📌 Guardar el documento con el nombre del trabajador
    nombre_archivo = os.path.join(ruta_guardar, f"{fila.get('Nombre completo', 'SinNombre').replace('/', '-').replace(':', '-')}.docx")
    doc.save(nombre_archivo)
    print(f"✅ Documento generado: {nombre_archivo}")

print("✅ TODOS los documentos han sido generados correctamente.")
