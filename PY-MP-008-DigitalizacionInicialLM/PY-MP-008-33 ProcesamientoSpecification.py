import pandas as pd
import sys
import string


# Función para procesar los datos de la quinta columna (columna de índice 4)
def process_column_5(data):
    listaspecification_input = []

    # Extraer las celdas de la quinta columna con contenido
    for value in data.iloc[:, 4]:  # Índice 4 es la quinta columna (E)
        if pd.notna(value):  # Verificar si la celda no está vacía
            listaspecification_input.append(str(value))

    # Eliminar el contenido de la quinta columna
    data.iloc[:, 4] = ""

    # Expansión de los elementos que contengan '|'
    expanded_list = []
    for element in listaspecification_input:
        parts = element.split('|')
        expanded_list.extend(parts)

    listaspecification_input = expanded_list

    # print("Lista después de la expansión:", listaspecification_input)

    # Eliminar elementos con 3 o menos caracteres o que sean "SPECIFICATION"
    listaspecification_input = [
        item for item in listaspecification_input if len(item) > 3 and item not in ["SPECIFICATION"]
    ]

    # print("Lista después de eliminar elementos no válidos:", listaspecification_input)

    # Modificaciones de formato en los strings
    def format_string(item):
        item = item.strip().upper().replace(',', ', ')  # Convertir a mayúsculas, agregar espacio después de coma

        # Sustituir '/' según las reglas especificadas
        new_item = []
        for i, char in enumerate(item):
            if char == '/':
                # Chequear si está entre dos consonantes
                if (i > 0 and i < len(item) - 1 and
                        item[i - 1].lower() not in "aeiou" and
                        item[i + 1].lower() not in "aeiou"):
                    new_item.append('I')
                # Chequear si está junto a un número
                elif (i > 0 and item[i - 1].isdigit()) or (i < len(item) - 1 and item[i + 1].isdigit()):
                    new_item.append('1')
                else:
                    new_item.append(char)  # Mantener el '/' si no cumple ninguna condición
            else:
                new_item.append(char)
        item = ''.join(new_item)

        # Sustituir caracteres específicos
        item = item.replace('*', '#').replace('%', ' 1/4').replace('YA', ' 1/4').replace('§', '&').replace('2B', '5/8').replace('VG','3/8').replace('4B','3/8').replace(
            'IG', '16').replace('(B', '1/8').replace('5B', '5/8').replace('YB', '5/8').replace('MOM', 'MDM').replace('|', '/').replace('I', '/').replace('l', '/')
        if item.count('X') >= 2 or item.count('x') >= 2:
            item = item.replace('x', ' x ').replace('X', ' x ')



        # Eliminar caracteres no alfabéticos al final del string
        item = item.rstrip(string.punctuation + string.digits)
        return item

    listaspecification_input = [format_string(item) for item in listaspecification_input]

    print("Lista después del formato:", listaspecification_input)

    return listaspecification_input


# Función para guardar la lista resultante en la quinta columna de nuevo
def save_to_excel(file_path, data, listaspecification_output):
    # Asegúrate de que haya suficientes filas en el DataFrame
    required_rows = len(listaspecification_output) + 2  # +2 porque empezamos en la fila 3
    if required_rows > len(data):
        # Si hay menos filas, podemos agregar filas vacías
        for _ in range(required_rows - len(data)):
            data.loc[len(data)] = [None] * data.shape[1]  # Agregar una fila vacía

    # Sobrescribir la quinta columna con la nueva lista
    for i, value in enumerate(listaspecification_output, start=2):  # Empezar desde la fila 3 (índice 2)
        data.iloc[i, 4] = value  # Índice 4 es la columna E

    # Guardar el archivo Excel
    data.to_excel(file_path, index=False)


# Flujo principal
if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Uso: PY-MP-008-33 ProcesamientoSpecification.py <ruta_archivo_pdf> <ruta_archivo_excel>")
        sys.exit(1)

    pdf_file_path = sys.argv[1]  # Argumento del PDF que no se utilizará
    excel_file_path = sys.argv[2]

    # Leer el archivo Excel
    data = pd.read_excel(excel_file_path, engine='openpyxl')

    # Verificar si tiene al menos 5 columnas
    if data.shape[1] >= 5:
        # Procesar la quinta columna
        listaspecification_input = process_column_5(data)

        # Crear la lista de salida
        listaspecification_output = listaspecification_input

        # Guardar los resultados en el archivo Excel
        save_to_excel(excel_file_path, data, listaspecification_output)

        print("Los datos se han guardado correctamente en el archivo.")
    else:
        print("El archivo Excel no tiene suficientes columnas.")
