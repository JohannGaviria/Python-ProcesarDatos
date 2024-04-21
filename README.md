# Script de Manipulación de Archivos Excel

Este script Python permite al usuario filtrar y procesar datos en archivos Excel.

## Instalación

1. Clona el repositorio con el siguiente comando:
    ```bash
    git clone https://github.com/JohannGaviria/Python-ProcesarDatos.git
    ```

2. Crea un entorno virtual con `virtualenv` u otro gestor de entornos virtuales:
    ```bash
    cd Python-ProcesarDatos
    python -m virtualenv venv
    ```

3. Inicia el entorno virtual:
    - En Windows:
        ```bash
        venv\Scripts\activate
        ```
    - En macOS y Linux:
        ```bash
        source venv/bin/activate
        ```

4. Instala las dependencias necesarias del archivo `requeriments.txt`:
    ```bash
    pip install -r requeriments.txt
    ```

## Funciones

### `read_file(file_path: str) -> pandas.DataFrame or None`

Lee un archivo Excel y devuelve un DataFrame.

- **Parámetros:**
    - `file_path` (`str`): Ruta del archivo Excel a leer.

- **Devuelve:**
    - `pandas.DataFrame or None`: DataFrame con los datos del archivo leído o `None` si ocurre un error.

### `validate_header(df: pandas.DataFrame, header: str) -> bool`

Valida si un encabezado existe en el DataFrame.

- **Parámetros:**
    - `df` (`pandas.DataFrame`): DataFrame para validar.
    - `header` (`str`): Encabezado a verificar.

- **Devuelve:**
    - `bool`: `True` si el encabezado existe, `False` si no.

### `validate_data(df: pandas.DataFrame, header: str, data: str) -> bool`

Valida si un dato está presente en una columna específica del DataFrame.

- **Parámetros:**
    - `df` (`pandas.DataFrame`): DataFrame para validar.
    - `header` (`str`): Encabezado en el que se busca el dato.
    - `data` (`str`): Dato a buscar.

- **Devuelve:**
    - `bool`: `True` si el dato está presente, `False` si no.

### `process_data(df: pandas.DataFrame, file_path: str, output_folder: str, header: str, data: str) -> None`

Procesa los datos y los guarda en un nuevo archivo Excel.

- **Parámetros:**
    - `df` (`pandas.DataFrame`): DataFrame con los datos filtrados.
    - `file_path` (`str`): Ruta del archivo de entrada.
    - `output_folder` (`str`): Carpeta de salida donde se guardarán los archivos modificados.
    - `header` (`str`): Encabezado del DataFrame.
    - `data` (`str`): Dato a buscar en el encabezado.

- **Devuelve:**
    - `None`

## Punto de Entrada del Programa

### `main() -> None`

Función principal que gestiona la ejecución del script.

- **Parámetros:**
    - Ninguno.

- **Devuelve:**
    - `None`
