# Gestor de Unificación Henry

Esta aplicación de escritorio, desarrollada en Python con `tkinter` y `pandas`, está diseñada para automatizar la extracción, normalización, consolidación y enriquecimiento de datos de archivos Excel y CSV. Es particularmente útil para procesar grandes volúmenes de datos distribuidos en múltiples archivos y carpetas, y para preparar informes unificados.

## Funcionalidades Principales

La aplicación se divide en dos secciones principales: el "Proceso Principal" para la consolidación de datos por etapas, y la "Herramienta de Cruce" para enriquecer archivos existentes.

### 1. Proceso Principal

Esta sección permite procesar y unificar datos en dos etapas:

#### Etapa 1: Procesar Lotes por Carpetas

1.  **Selección de Carpeta de Entrada (Lotes):** El usuario selecciona una carpeta principal que contiene subcarpetas. Cada subcarpeta se considera un "lote" de archivos.
2.  **Extracción de Datos:** La aplicación busca archivos `.xlsx`, `.xls` y `.csv` dentro de cada subcarpeta. Para cada archivo, busca una tabla específica identificada por la celda 'Tipo' en la columna B. Extrae filas de datos a partir de esa cabecera.
3.  **Normalización Numérica:** Los valores numéricos (Cantidad, Precio unitario eD, Precio total eD, Precio unitario cliente, Precio total cliente) se normalizan para manejar diferentes separadores decimales (coma/punto) y formatos. Se aplica una heurística para asegurar la consistencia entre Cantidad, Precio unitario eD y Precio total eD.
4.  **Guardado de Lotes Procesados:** Los datos extraídos y normalizados de cada lote se guardan como un archivo `.xlsx` individual en una "Carpeta de Destino" especificada.
5.  **Limpieza Opcional:** Se ofrece la opción de limpiar (eliminar todos los archivos de) las subcarpetas de entrada una vez que han sido procesadas exitosamente.

#### Etapa 2: Unificación Final

1.  **Selección de Carpeta de Lotes Procesados:** El usuario selecciona la carpeta donde se guardaron los archivos `.xlsx` de la Etapa 1.
2.  **Consolidación:** La aplicación lee todos los archivos `.xlsx` de esta carpeta y los concatena en un único DataFrame.
3.  **Filtro Opcional:** Permite aplicar un filtro a los datos consolidados. El usuario puede seleccionar una columna y especificar una lista de valores permitidos. Solo las filas que contengan alguno de esos valores en la columna seleccionada se incluirán en el resultado final.
4.  **Normalización Numérica (Consolidado):** Se aplica una normalización numérica adicional al DataFrame consolidado para asegurar la consistencia final.
5.  **Guardado del Archivo Unificado:** El DataFrame consolidado (y opcionalmente filtrado) se guarda como un único archivo `.xlsx` en la ubicación especificada por el usuario.

### 2. Herramienta de Cruce (Enriquecimiento)

Esta herramienta permite enriquecer un archivo base con información de un archivo secundario (de enriquecimiento).

1.  **Selección de Archivos:** El usuario selecciona un "Archivo Base" y un "Archivo de Enriquecimiento" (ambos pueden ser Excel o CSV).
2.  **Definición de Claves:** Se especifican las columnas clave en ambos archivos para realizar la unión (merge).
3.  **Selección de Columnas a Agregar:** El usuario elige qué columnas del archivo de enriquecimiento desea añadir al archivo base.
4.  **Selección de Columnas a Quitar:** Opcionalmente, el usuario puede seleccionar columnas del archivo base que desea eliminar del resultado final.
5.  **Normalización Numérica:** Se aplica normalización numérica a ambos DataFrames antes de la unión.
6.  **Unión y Guardado:** La aplicación realiza una unión `left_on` (izquierda) para agregar las columnas seleccionadas del archivo de enriquecimiento al archivo base. El resultado se guarda como un nuevo archivo `.xlsx`.

### Formato de Salida Numérico

Ambas secciones permiten elegir el formato de salida para las columnas numéricas:
-   **Número Excel:** Los números se guardan en formato numérico estándar de Excel.
-   **Texto con coma:** Los números se convierten a texto, utilizando la coma como separador decimal (ej. `1.23` se convierte en `1,23`).

### Registro de Actividad

La aplicación incluye una consola integrada que muestra el progreso de las operaciones, advertencias y errores, proporcionando retroalimentación en tiempo real al usuario.

## Cómo Ejecutar

1.  Asegúrate de tener Python instalado (versión 3.x recomendada).
2.  Instala las dependencias necesarias: `pandas`, `openpyxl`, `tkinter`. Puedes hacerlo con `pip`:
    ```bash
    pip install pandas openpyxl
    ```
3.  Ejecuta el script principal:
    ```bash
    python main.py
    ```
```