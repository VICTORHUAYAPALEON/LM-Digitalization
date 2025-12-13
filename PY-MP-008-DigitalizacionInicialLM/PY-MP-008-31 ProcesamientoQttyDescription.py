import pandas as pd
import sys
import string
import wordninja

# Función para procesar los datos de la tercera columna (columna de índice 2)
def process_column_3(data):
    listaqttydescription_input = []

    # Extraer las celdas de la tercera columna con contenido
    for value in data.iloc[:, 2]:  # Índice 2 es la tercera columna (C)
        if pd.notna(value):  # Verificar si la celda no está vacía
            listaqttydescription_input.append(str(value))

    # Eliminar el contenido de la tercera columna
    data.iloc[:, 2] = ""

    # Expansión de los elementos que contengan '|'
    expanded_list = []
    for element in listaqttydescription_input:
        parts = element.split('|')
        expanded_list.extend(parts)

    listaqttydescription_input = expanded_list

    #print("Lista después de la expansión:", listaqttydescription_input)

    # Eliminar elementos con 3 o menos caracteres o que sean "QUAN" o "DESCRIPTION"
    listaqttydescription_input = [
        item for item in listaqttydescription_input if len(item) > 3 and item not in ["QUAN", "DESCRIPTION"]
    ]

    #print("Lista después de eliminar elementos no válidos:", listaqttydescription_input)

    # Modificaciones de formato en los strings
    def format_string(item):
        item = item.strip().upper().replace(',', ', ')  # Convertir a mayúsculas, agregar espacio después de coma
        # Eliminar caracteres no alfabéticos al final del string
        item = item.rstrip(string.punctuation + string.digits)
        # Usando wordninja para separar las palabras
        segmented = wordninja.split(item)
        # Uniéndolas con espacios
        item = " ".join(segmented)
        return item

    listaqttydescription_input = [format_string(item) for item in listaqttydescription_input]

    print("Lista después del formato:", listaqttydescription_input)

    return listaqttydescription_input

# Función para guardar la lista resultante en la tercera columna de nuevo
def save_to_excel(file_path, data, listaqttydescription_output):
    # Asegúrate de que haya suficientes filas en el DataFrame
    required_rows = len(listaqttydescription_output) + 2  # +2 porque empezamos en la fila 3
    if required_rows > len(data):
        # Si hay menos filas, podemos agregar filas vacías
        for _ in range(required_rows - len(data)):
            data.loc[len(data)] = [None] * data.shape[1]  # Agregar una fila vacía

    # Sobrescribir la quinta columna con la nueva lista
    for i, value in enumerate(listaqttydescription_output, start=2):  # Empezar desde la fila 3 (índice 2)
        data.iloc[i, 2] = value  # Índice 2 es la columna C

    # Guardar el archivo Excel
    data.to_excel(file_path, index=False)


# Flujo principal
if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Uso: PY-MP-008-31 ProcesamientoQttyDescription.py <ruta_archivo_pdf> <ruta_archivo_excel>")
        sys.exit(1)

    pdf_file_path = sys.argv[1]
    excel_file_path = sys.argv[2]

    # Leer el archivo Excel
    data = pd.read_excel(excel_file_path, engine='openpyxl')

    # Verificar si tiene al menos 3 columnas
    if data.shape[1] >= 3:
        # Procesar la tercera columna
        listaqttydescription_input = process_column_3(data)

        # Crear la lista de salida
        listaqttydescription_output = listaqttydescription_input

        # Guardar los resultados en el archivo Excel
        save_to_excel(excel_file_path, data, listaqttydescription_output)

        print("Los datos se han guardado correctamente en el archivo.")
    else:
        print("El archivo Excel no tiene suficientes columnas.")
