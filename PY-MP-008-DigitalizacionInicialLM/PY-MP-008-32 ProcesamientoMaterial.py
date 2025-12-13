import pandas as pd
import sys
import string

# Función para procesar los datos de la cuarta columna (columna de índice 3)
def process_column_4(data):
    listamaterial_input = []

    # Extraer las celdas de la cuarta columna con contenido
    for value in data.iloc[:, 3]:  # Índice 3 es la cuarta columna (D)
        if pd.notna(value):  # Verificar si la celda no está vacía
            listamaterial_input.append(str(value))

    # Eliminar el contenido de la cuarta columna
    data.iloc[:, 3] = ""

    # Expansión de los elementos que contengan '|'
    expanded_list = []
    for element in listamaterial_input:
        parts = element.split('|')
        expanded_list.extend(parts)

    listamaterial_input = expanded_list

    #print("Lista después de la expansión:", listamaterial_input)

    # Eliminar elementos con 2 o menos caracteres
    listamaterial_input = [
        item for item in listamaterial_input if len(item) > 2
    ]

    #print("Lista después de eliminar elementos no válidos:", listamaterial_input)

    # Modificaciones de formato en los strings
    def format_string(item):
        item = item.strip().upper().replace(',', ', ')  # Convertir a mayúsculas, agregar espacio después de coma
        # Eliminar caracteres no alfabéticos al final del string
        item = item.rstrip(string.punctuation + string.digits)
        return item

    listamaterial_input = [format_string(item) for item in listamaterial_input]

    print("Lista después del formato:", listamaterial_input)

    return listamaterial_input

# Función para guardar la lista resultante en la cuarta columna de nuevo
def save_to_excel(file_path, data, listamaterial_output):
    # Asegúrate de que haya suficientes filas en el DataFrame
    required_rows = len(listamaterial_output) + 2  # +2 porque empezamos en la fila 3
    if required_rows > len(data):
        # Si hay menos filas, podemos agregar filas vacías
        for _ in range(required_rows - len(data)):
            data.loc[len(data)] = [None] * data.shape[1]  # Agregar una fila vacía

    # Sobrescribir la quinta columna con la nueva lista
    for i, value in enumerate(listamaterial_output, start=2):  # Empezar desde la fila 3 (índice 2)
        data.iloc[i, 3] = value  # Índice 3 es la columna D

    # Guardar el archivo Excel
    data.to_excel(file_path, index=False)


# Flujo principal
if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Uso: PY-MP-008-32 ProcesamientoMaterial.py <ruta_archivo_pdf> <ruta_archivo_excel>")
        sys.exit(1)

    pdf_file_path = sys.argv[1]
    excel_file_path = sys.argv[2]

    # Leer el archivo Excel
    data = pd.read_excel(excel_file_path, engine='openpyxl')

    # Verificar si tiene al menos 4 columnas
    if data.shape[1] >= 4:
        # Procesar la cuarta columna
        listamaterial_input = process_column_4(data)

        # Crear la lista de salida
        listamaterial_output = listamaterial_input

        # Guardar los resultados en el archivo Excel
        save_to_excel(excel_file_path, data, listamaterial_output)

        print("Los datos se han guardado correctamente en el archivo.")
    else:
        print("El archivo Excel no tiene suficientes columnas.")
