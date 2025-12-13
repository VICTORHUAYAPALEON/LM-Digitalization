import subprocess
import webbrowser
from tkinter import Tk, filedialog

# Función para abrir ventana emergente y seleccionar archivo
def seleccionar_archivo(tipo_archivo):
    root = Tk()
    root.withdraw()  # Oculta la ventana principal de Tkinter
    if tipo_archivo == 'pdf':
        archivo = filedialog.askopenfilename(title="Selecciona el archivo PDF", filetypes=[("PDF files", "*.pdf")])
    elif tipo_archivo == 'excel':
        archivo = filedialog.askopenfilename(title="Selecciona el archivo Excel", filetypes=[("Excel files", "*.xlsx *.xls")])
    root.destroy()
    return archivo

# Seleccionar archivo PDF y Excel
archivo_pdf = seleccionar_archivo('pdf')
archivo_excel = seleccionar_archivo('excel')

# Abrir el archivo PDF en Microsoft Edge
try:
    edge_path = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"
    subprocess.run([edge_path, archivo_pdf], check=True)
except FileNotFoundError:
    print("No se encontró Microsoft Edge en la ruta especificada. Abriendo con el navegador predeterminado.")
    webbrowser.open(archivo_pdf)


# Lista de archivos .py a ejecutar
archivos = [
    "PY-MP-008-21 ExtraccionQttyDescription.py",
    "PY-MP-008-31 ProcesamientoQttyDescription.py",
    "PY-MP-008-22 ExtraccionMaterial.py",
    "PY-MP-008-32 ProcesamientoMaterial.py",
    "PY-MP-008-23 ExtraccionSpecification.py",
    "PY-MP-008-33 ProcesamientoSpecification.py"
]

# Ejecutar cada archivo en secuencia con los argumentos de archivo_pdf y archivo_excel
for archivo in archivos:
    try:
        print(f"Ejecutando {archivo}...")
        subprocess.run(["python", archivo, archivo_pdf, archivo_excel], check=True)
        print(f"Terminó de ejecutar {archivo}.\n")
    except subprocess.CalledProcessError as e:
        print(f"Error al ejecutar {archivo}: {e}\n")


