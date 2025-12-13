import os
import shutil


def obtener_ruta_entrada_y_salida():
    """
    Solicita al usuario las rutas de la carpeta de entrada y salida.
    """
    carpeta_entrada = input("Por favor, ingresa la ruta de la carpeta de entrada: ")
    carpeta_salida = input("Por favor, ingresa la ruta de la carpeta de salida: ")
    return carpeta_entrada, carpeta_salida


def validar_subcarpeta(subcarpeta):
    """
    Valida si el nombre de la subcarpeta sigue el formato M-XXXXX revY.
    XXXXX debe ser un número entre 20000 y 80000.
    revY puede ser cualquier letra del abecedario o 'revNOIND'.
    """
    # Verificar el patrón de la subcarpeta
    partes = subcarpeta.split(" ")
    if len(partes) != 2:
        return False

    nombre, revision = partes
    if not nombre.startswith("M-") or not nombre[2:].isdigit() or not (20000 <= int(nombre[2:]) <= 80000):
        return False

    if not revision.startswith("rev"):
        return False

    revision_sin_rev = revision[3:]
    if not revision_sin_rev.isalpha() and revision_sin_rev != "NOIND":
        return False

    return True


def procesar_carpetas(carpeta_entrada, carpeta_salida):
    """
    Procesa las subcarpetas dentro de la carpeta de entrada y extrae los archivos Excel.
    """
    for subcarpeta in os.listdir(carpeta_entrada):
        subcarpeta_path = os.path.join(carpeta_entrada, subcarpeta)

        # Verifica si es una subcarpeta y si tiene el formato correcto
        if os.path.isdir(subcarpeta_path) and validar_subcarpeta(subcarpeta):
            archivo_excel = subcarpeta + ".xlsx"
            archivo_excel_path = os.path.join(subcarpeta_path, archivo_excel)

            # Verificar si el archivo Excel existe en la subcarpeta
            if os.path.exists(archivo_excel_path):
                # Copiar el archivo Excel a la carpeta de salida
                destino = os.path.join(carpeta_salida, archivo_excel)
                shutil.copy(archivo_excel_path, destino)
                print(f"Archivo {archivo_excel} copiado a la carpeta de salida.")
            else:
                print(f"No se encontró el archivo Excel {archivo_excel} en la subcarpeta {subcarpeta}.")
        else:
            print(f"La subcarpeta {subcarpeta} no sigue el formato esperado.")


def main():
    # Obtener rutas de entrada y salida
    carpeta_entrada, carpeta_salida = obtener_ruta_entrada_y_salida()

    # Verificar si las carpetas existen
    if not os.path.exists(carpeta_entrada):
        print("La carpeta de entrada no existe.")
        return
    if not os.path.exists(carpeta_salida):
        print("La carpeta de salida no existe.")
        return

    # Procesar las subcarpetas
    procesar_carpetas(carpeta_entrada, carpeta_salida)


if __name__ == "__main__":
    main()
