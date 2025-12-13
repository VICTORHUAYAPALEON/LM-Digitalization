import os
import sys
from tkinter import Tk, filedialog
from pdf2image import convert_from_path
from PIL import Image, ImageEnhance, ImageFilter
import pytesseract
import cv2
import openpyxl
import numpy as np  # Asegúrate de importar numpy


# Configurar la ruta de Tesseract
tesseract_path = r'C:/Program Files/Tesseract-OCR/tesseract.exe'
pytesseract.pytesseract.tesseract_cmd = tesseract_path

# # Función para abrir ventana emergente y seleccionar archivo
# def seleccionar_archivo(tipo_archivo):
#     root = Tk()
#     root.withdraw()  # Oculta la ventana principal de Tkinter
#     if tipo_archivo == 'excel':
#         archivo = filedialog.askopenfilename(title="Selecciona el archivo Excel",
#                                              filetypes=[("Excel files", "*.xlsx *.xls")])
#     elif tipo_archivo == 'pdf':
#         archivo = filedialog.askopenfilename(title="Selecciona el archivo PDF", filetypes=[("PDF files", "*.pdf")])
#     root.destroy()
#     return archivo


# Función para procesar la imagen
def preprocesar_imagen(imagen):

    # Convertir a escala de grises
    gray_img = imagen.convert('L')

    # Aumento de contraste
    enhancer = ImageEnhance.Contrast(gray_img)
    high_contrast_img = enhancer.enhance(2)

    # Binarización
    image = np.array(high_contrast_img)
    _, binary_image = cv2.threshold(image, 128, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

    # Eliminación de ruido
    denoised_img = cv2.fastNlMeansDenoising(binary_image, h=30)

    # Dilatación y Erosión
    kernel = np.ones((2, 2), np.uint8)
    dilated_img = cv2.dilate(denoised_img, kernel, iterations=1)
    eroded_img = cv2.erode(dilated_img, kernel, iterations=2)

    preprocessed_img = Image.fromarray(eroded_img)

    return preprocessed_img

# Función principal
def main():

    # Verificar que se recibieron los argumentos necesarios
    if len(sys.argv) != 3:
        print("Uso: PY-MP-008-21 ExtraccionQttyDescription.py <archivo_pdf> <archivo_excel>")
        sys.exit(1)

    # Seleccionar archivo Excel y PDF
    archivo_pdf = sys.argv[1]
    archivo_excel = sys.argv[2]

    # Convertir PDF a imágenes
    imagenes = convert_from_path(archivo_pdf,
                                 poppler_path=r'D:\03. PROFESSIONAL EXP\02. MARCO PERUANA\UIPath+Python\PY-MP-008 DigitalizacionInicialLM\poppler-24.02.0\Library\bin',
                                 dpi=300)

    # Inicializar variable QttyDescription
    QttyDescription = ""

    for imagen in imagenes:

        width, height = imagen.size

        if width > height:
            izquierda = int(0.125 * width)
            top = int(0.08 * height)
            derecha = int(0.31 * width)
            bottom = int(0.9 * height)

        elif height > width:
            width, height = imagen.size
            izquierda = int(0.265 * width)
            top = int(0.08 * height)
            derecha = int(0.59 * width)
            bottom = int(0.8 * height)

        # Recortar la imagen
        imagen_recortada = imagen.crop((izquierda, top, derecha, bottom))

        # Preprocesar la imagen
        imagen_preprocesada = preprocesar_imagen(imagen_recortada)

        # # Mostrar la imagen preprocesada
        # imagen_preprocesada.show()

        # Realizar OCR en la imagen preprocesada
        custom_config = r'--dpi 300 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz,.-& ' \
                        r'-c preserve_interword_spaces=1 -c chop_enable=0 -c segment_segcost_rating=2 ' \
                        r'-c textord_space_size_is_variable=0'

        texto_extraido = pytesseract.image_to_string(imagen_preprocesada, config=custom_config)

        # Concatenar el texto extraído a QttyDescription
        QttyDescription += texto_extraido + '\n'

    # Imprimir QttyDescription
    QttyDescription = QttyDescription.replace("\n", "|")
    print(QttyDescription)

    # Generar la lista de palabras ListaQttyDescription
    ListaQttyDescription = QttyDescription.split("||")

    # Cargar el archivo Excel
    workbook = openpyxl.load_workbook(archivo_excel)
    sheet = workbook.active

    # Insertar las palabras en el Excel de forma vertical a partir de la celda C3
    fila_inicial = 3
    columna = 3  # Columna C

    for index, palabra in enumerate(ListaQttyDescription):
        celda = sheet.cell(row=fila_inicial + index, column=columna)
        celda.value = palabra

    # Guardar el archivo Excel
    workbook.save(archivo_excel)


if __name__ == "__main__":
    main()
