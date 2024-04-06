import pandas as pd
import os


# Leer un archivo Excel
def read_file(file_path):
    # Intentar leer el archivo Excel
    try:
        # Devolver un DataFrame
        return pd.read_excel(file_path)

    # Capturar la excepción si el archivo no se encuentra
    except FileNotFoundError:
        print("Error: El archivo especificado no se encontró.")
        return None

    # Captura cualquier otra excepción
    except Exception as e:
        print(f"Error: Ocurrió un error al leer el archivo: {e}")
        return None


# Validar si un encabezado existe en el DataFrame
def validate_header(df, header):
    # Verificar si el encabezado no está presente en las columnas del DataFrame
    if header not in df.columns:
        print(f"Error: El encabezado '{header}' no existe en el archivo.")
        return False
    return True


# Validar si un dato está presente en una columna específica del DataFrame
def validate_data(df, header, data):
    # Verificar si el dato no está presente en la columna especificada
    if data not in df[header].unique():
        print(f"Error: El dato '{data}' no se encontró en el encabezado '{header}'.")
        return False
    return True


# Procesar los datos
def process_data(df, file_path, output_folder, header, data):
    # Verificar si el DataFrame filtrado no está vacío
    if not df.empty:
        try:
            # Genera un nuevo nombre de archivo
            new_file_name = os.path.splitext(os.path.basename(file_path))[0] + "_indicted.xlsx"
            # Crea la ruta de salida para el nuevo archivo
            output_path = os.path.join(output_folder, new_file_name)
            # Guarda el DataFrame filtrado en un nuevo archivo Excel sin índice
            df.to_excel(output_path, index=False)

            print(f'El archivo con los cambios se ha guardado como: "{output_path}"')

        # Captura la excepción si no se tienen permisos para guardar el archivo
        except PermissionError:
            print(f"Error: No tienes permiso para guardar el archivo en '{output_folder}'.")

        # Captura cualquier otra excepción
        except Exception as e:
            print(f"Error: Ocurrió un error al guardar el archivo: {e}")

    # Si el DataFrame filtrado está vacío
    else:
        print(f'Error: El dato especificado: "{data}" para el encabezado "{header}" no fue encontrado en el archivo.')


def main():
    # Carpeta de salida donde se guardarán los archivos modificados
    output_folder = "files"

    # Intenta crear la carpeta de salida si no existe
    try:
        os.makedirs(output_folder, exist_ok=True)

    # Captura la excepción si no se puede crear la carpeta de salida
    except OSError as e:
        print(f"Error: No se pudo crear la carpeta de salida: {e}")
        return

    print("POR FAVOR, INGRESE LOS DATOS EXACTAMENTE")

    while True:
        # Solicita al usuario el nombre del archivo
        file_name = input("Ingrese el nombre del archivo (o 'exit' para salir): ")
        # Verifica si el usuario quiere salir del programa
        if file_name.lower() == 'exit':
            break

        # Crea la ruta completa del archivo
        file_path = os.path.join(output_folder, file_name)

        # Verifica si el archivo tiene la extensión .xlsx
        if not file_name.lower().endswith('.xlsx'):
            print("Error: El archivo debe tener la extensión '.xlsx'.")
            continue

        # Lee el archivo Excel y obtiene un DataFrame
        df = read_file(file_path)
        # Verifica si ocurrió un error al leer el archivo
        if df is None:
            continue

        # Solicita al usuario el nombre del encabezado
        header = input("Ingrese el nombre del encabezado: ")
        # Verifica si el encabezado proporcionado existe en el archivo
        if not validate_header(df, header):
            continue

        # Solicita al usuario el dato a buscar
        data = input("Ingrese el dato a buscar: ")
        # Verifica si el dato proporcionado está presente en el encabezado especificado
        if not validate_data(df, header, data):
            continue

        # Filtra el DataFrame con el dato proporcionado
        df_filtered = df[df[header].str.lower() == data.lower()]

        # Procesa y guarda los datos filtrados en un nuevo archivo Excel
        process_data(df_filtered, file_path, output_folder, header, data)


# Punto de entrada del programa
if __name__ == "__main__":
    main()
