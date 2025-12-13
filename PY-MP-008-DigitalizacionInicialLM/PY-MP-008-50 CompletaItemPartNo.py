"""
ESTADO: MÓDULO AUXILIAR (VÁLIDO / NO LEGACY)

COMPLETA ITEMS SEGÚN PARTNo
Este script automatiza el llenado de información cuando el PART NUMBER
corresponde a un componente conocido dentro de una tabla maestra.
En estos casos, no es necesario realizar OCR ni extracción por coordenadas.

FUNCIONALIDAD:
- Lee el PARTNo detectado previamente (ej. '18743').
- Consulta una tabla maestra local con información técnica previamente verificada
  (ej. 'HYDRAULIC PUMP V20 VICKERS', especificaciones, fabricante).
- Completa automáticamente los campos MATERIAL / DESCRIPTION / SPECIFICATION
  sin pasar por los módulos de extracción OCR.
- Ahorra tiempo y reduce errores en componentes frecuentes o estandarizados.

NOTA:
Este módulo funciona como atajo inteligente dentro del pipeline
cuando el PARTNo pertenece al catálogo interno de la empresa.
"""

import pandas as pd
import tkinter as tk
from tkinter import filedialog


def seleccionar_archivo(titulo):
    root = tk.Tk()
    root.withdraw()  # Oculta la ventana principal de Tkinter
    archivo = filedialog.askopenfilename(title=titulo, filetypes=[("Excel files", "*.xlsx;*.xls")])
    return archivo


def limpiar_nombres_columnas(df):
    df.columns = df.columns.str.strip()
    return df


# Seleccionar el archivo ExcelBase
ruta_excel_base = seleccionar_archivo("Selecciona el archivo ExcelBase")
ExcelBase = pd.read_excel(ruta_excel_base, usecols="A:E", header=None)
TablaExcelBase = ExcelBase.copy()
print("TablaExcelBase inicial (sin cabecera):")
print(TablaExcelBase)

# Seleccionar el archivo ExcelItems
ruta_excel_items = seleccionar_archivo("Selecciona el archivo ExcelItems")
ExcelItems = pd.read_excel(ruta_excel_items, usecols="A:E", header=0)
TablaItems = limpiar_nombres_columnas(ExcelItems.copy())

# Imprimir nombres de columnas de TablaItems
print("Nombres de columnas de TablaItems:")
print(TablaItems.columns)

# Procesamiento de las tablas
for idx_base, fila_base in TablaExcelBase.iterrows():
    PartNoBase = pd.to_numeric(fila_base[0], errors='coerce')

    # Definir CondicionValidez
    if (10000 < PartNoBase < 20000) or (80000 < PartNoBase < 90000):
        for idx_items, fila_items in TablaItems.iterrows():
            PartNoItems = pd.to_numeric(fila_items[TablaItems.columns[0]], errors='coerce')

            if PartNoBase == PartNoItems:
                TablaExcelBase.at[idx_base, 2] = fila_items[TablaItems.columns[2]]
                TablaExcelBase.at[idx_base, 3] = fila_items[TablaItems.columns[3]]
                TablaExcelBase.at[idx_base, 4] = fila_items[TablaItems.columns[4]]
                break

print("\nTablaExcelBase después del procesamiento:")
print(TablaExcelBase)

# Guardar el resultado en ExcelBase
TablaExcelBase.to_excel(ruta_excel_base, index=False, header=False)
