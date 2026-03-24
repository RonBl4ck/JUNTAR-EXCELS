import pandas as pd
import os
import glob
import warnings
import re
from config import settings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl") 

def extract_table_from_file(file_path, lcl_name):
    """
    Finds the row where column B (index 1) is 'Tipo' and extracts data.
    Handles both Excel and CSV files.
    """
    found_rows = []
    expected_cols_count = 11
    file_ext = os.path.splitext(file_path)[1].lower()

    try:
        # --- CSV Processing ---
        if file_ext == '.csv':
            # For CSV, we read the whole file and then find the table
            # Try with different encodings and delimiters if needed
            try:
                df = pd.read_csv(file_path, header=None, encoding='utf-8', sep=';')
            except (UnicodeDecodeError, pd.errors.ParserError):
                try:
                    df = pd.read_csv(file_path, header=None, encoding='latin1', sep=';')
                except pd.errors.ParserError:
                    try:
                        df = pd.read_csv(file_path, header=None, encoding='utf-8', sep=',')
                    except Exception as e:
                         print(f"Warning: Could not parse CSV {os.path.basename(file_path)} with any common format: {e}")
                         return []

            if df.shape[1] < 2:
                return []

            i = 0
            while i < len(df):
                try:
                    cell_b = df.iloc[i, 1]
                except IndexError:
                    i += 1
                    continue

                if isinstance(cell_b, str) and str(cell_b).strip() == "Tipo":
                    if df.shape[1] < 1 + expected_cols_count:
                        i += 1
                        continue
                    
                    current_row = i + 1
                    while current_row < len(df):
                        val_b = df.iloc[current_row, 1]
                        if pd.isna(val_b) or str(val_b).strip() == "":
                            break
                        
                        raw_data = df.iloc[current_row, 1:1+expected_cols_count].tolist()
                        found_rows.append([lcl_name] + raw_data)
                        current_row += 1
                    i = current_row
                    continue
                i += 1
        
        # --- Excel Processing ---
        elif file_ext in ['.xlsx', '.xls']:
            xls = pd.ExcelFile(file_path)
            for sheet_name in xls.sheet_names:
                try:
                    df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                    
                    if df.shape[1] < 2:
                        continue

                    i = 0
                    while i < len(df):
                        try:
                            cell_b = df.iloc[i, 1]
                        except IndexError:
                            i += 1
                            continue

                        if isinstance(cell_b, str) and str(cell_b).strip() == "Tipo":
                            if df.shape[1] < 1 + expected_cols_count:
                                i += 1
                                continue
                            
                            current_row = i + 1
                            while current_row < len(df):
                                val_b = df.iloc[current_row, 1]
                                if pd.isna(val_b) or str(val_b).strip() == "":
                                    break
                                
                                raw_data = df.iloc[current_row, 1:1+expected_cols_count].tolist()
                                found_rows.append([lcl_name] + raw_data)
                                current_row += 1
                            i = current_row
                            continue
                        i += 1

                except Exception as e:
                    print(f"Warning: Error reading sheet '{sheet_name}' in {os.path.basename(file_path)}: {e}")

    except Exception as e:
        print(f"Error reading file {os.path.basename(file_path)}: {e}")
        return []

    return found_rows

def process_stage1_by_subfolders(input_dir, output_dir):
    """
    Stage 1: Find subfolders in input_dir and process each as a batch.
    Returns a list of successfully processed folder paths for cleaning.
    """
    subfolders = [f.path for f in os.scandir(input_dir) if f.is_dir()]
    
    if not subfolders:
        print(f"No se encontraron subcarpetas de lotes en {input_dir}.")
        return [], False

    print(f"Se encontraron {len(subfolders)} subcarpetas para procesar.")
    overall_success = False
    successfully_processed_folders = []

    for folder_path in subfolders:
        batch_name = os.path.basename(folder_path)
        print(f"\n--- Procesando Lote: {batch_name} ---")

        input_files = glob.glob(os.path.join(folder_path, "*.xlsx")) + \
                      glob.glob(os.path.join(folder_path, "*.xls")) + \
                      glob.glob(os.path.join(folder_path, "*.csv"))
        
        if not input_files:
            print(f" [!] No se encontraron archivos Excel o CSV en la carpeta '{batch_name}'.")
            continue

        print(f"Procesando {len(input_files)} archivos para el lote '{batch_name}'...")
        
        all_rows = []
        for file_path in input_files:
            filename = os.path.basename(file_path)
            name_no_ext = os.path.splitext(filename)[0]

            if 'lcl' in name_no_ext.lower():
                # Find all sequences of digits
                numbers = re.findall(r'\d+', name_no_ext)
                if numbers:
                    # Assume the longest number is the one we want (to distinguish from LCL1, LCL2 etc)
                    lcl_name = max(numbers, key=len)
                else:
                    # Fallback if 'LCL' is in the name but no numbers are found
                    lcl_name = name_no_ext.split('-')[0].strip()
            else:
                # Original logic for names that do not contain 'LCL'
                lcl_name = name_no_ext.split('-')[0].strip()
            
            rows = extract_table_from_file(file_path, lcl_name)
            if rows:
                all_rows.extend(rows)
                print(f" [OK] {filename} -> {len(rows)} filas (LCL: {lcl_name})")
            else:
                print(f" [!] {filename} -> No se encontró tabla válida")

        if not all_rows:
            print(f"No se extrajeron datos de ningún archivo en el lote '{batch_name}'.")
            continue

        headers = [
            "LCL_Origen", "Tipo", "Contador", "Mat./Prest.", "Descripción Mat./Serv.", 
            "Cantidad", "Unidad medida base", "Imputación", 
            "Precio unitario eD", "Precio unitario cliente", 
            "Precio total eD", "Precio total cliente"
        ]

        df_batch = pd.DataFrame(all_rows, columns=headers)

        # --- Data Cleaning ---
        # Convert columns with comma decimals to dot decimals
        cols_to_clean = [
            "Cantidad", "Precio unitario eD", "Precio unitario cliente", 
            "Precio total eD", "Precio total cliente"
        ]
        print("Limpiando datos numéricos (reemplazando comas)...")
        for col in cols_to_clean:
            if col in df_batch.columns:
                try:
                    # Ensure column is string, replace comma, then convert to numeric
                    df_batch[col] = pd.to_numeric(
                        df_batch[col].astype(str).str.replace(',', '.'),
                        errors='coerce' # Invalid parsing will be set as NaT
                    )
                except Exception as e:
                    print(f" [Advertencia] No se pudo limpiar la columna '{col}': {e}")
        
        output_path = os.path.join(output_dir, f"{batch_name}.xlsx")
        try:
            df_batch.to_excel(output_path, index=False)
            print(f"\n[ÉXITO] Lote '{batch_name}' guardado en: {output_path}")
            print(f"Total filas: {len(df_batch)}")
            overall_success = True
            successfully_processed_folders.append(folder_path)
        except Exception as e:
            print(f"\n[ERROR] No se pudo guardar el lote '{batch_name}': {e}")

    return successfully_processed_folders, overall_success

def process_stage2_consolidation(input_dir, output_file_path, filter_by_tipo):
    """
    Stage 2: Consolidate all files from the input directory, with an option to filter by 'Tipo'.
    """
    input_files = glob.glob(os.path.join(input_dir, "*.xlsx"))
    
    if not input_files:
        print(f"No hay archivos procesados en {input_dir}.")
        return False

    print("--- Unificando Lotes ---")
    dfs = []
    for file_path in input_files:
        try:
            df = pd.read_excel(file_path)
            
            # Filter by 'Tipo' if the option is selected
            if filter_by_tipo:
                # Ensure 'Tipo' column exists
                if 'Tipo' in df.columns:
                    original_rows = len(df)
                    df = df[df['Tipo'] == 'Materiales']
                    print(f" [FILTRO] Lote '{os.path.basename(file_path)}': {len(df)} de {original_rows} filas son 'Materiales'.")
                else:
                    print(f" [ADVERTENCIA] La columna 'Tipo' no se encontró en '{os.path.basename(file_path)}', no se pudo filtrar.")

            # Add 'Tipo_Obra' column based on filename (e.g., 'DD001')
            batch_name = os.path.splitext(os.path.basename(file_path))[0]
            df.insert(0, "Tipo_Obra", batch_name)
            
            dfs.append(df)
            print(f" [OK] Leído lote: {batch_name} ({len(df)} filas)")
        except Exception as e:
            print(f" [Error] Falló al leer {os.path.basename(file_path)}: {e}")

    if not dfs:
        print("No se encontraron datos para consolidar (después de filtrar).")
        return False

    # Consolidate
    df_consolidated = pd.concat(dfs, ignore_index=True)
    print(f"\nTotal consolidado: {len(df_consolidated)} filas.")

    try:
        # Save Final Output
        df_consolidated.to_excel(output_file_path, index=False)
        print(f"\n[ÉXITO] Archivo final guardado en: {output_file_path}")
        return True
    except Exception as e:
        print(f"[ERROR] Falló al guardar el archivo consolidado: {e}")
        return False

def _read_file_to_df(file_path):
    """Reads an Excel or CSV file into a pandas DataFrame."""
    if not file_path or not os.path.exists(file_path):
        return None
    
    file_ext = os.path.splitext(file_path)[1].lower()
    try:
        if file_ext in ['.xlsx', '.xls']:
            return pd.read_excel(file_path)
        elif file_ext == '.csv':
            # Try common CSV formats
            try:
                return pd.read_csv(file_path, sep=',')
            except Exception:
                return pd.read_csv(file_path, sep=';', encoding='latin1')
    except Exception as e:
        print(f"Error al leer el archivo {os.path.basename(file_path)}: {e}")
        return None

def enrich_file(base_path, side_path, base_key, side_key, cols_to_add, output_path):
    """
    Enriches a base file with columns from a side file based on a key match.
    """
    try:
        print("--- Iniciando Proceso de Enriquecimiento ---")
        print(f"Archivo Base: {os.path.basename(base_path)}")
        print(f"Archivo Enriquecimiento: {os.path.basename(side_path)}")

        df_base = _read_file_to_df(base_path)
        df_side = _read_file_to_df(side_path)

        if df_base is None or df_side is None:
            print("[ERROR] No se pudieron leer uno o ambos archivos.")
            return False

        # --- Data Cleaning ---
        df_base.columns = [str(c).strip() for c in df_base.columns]
        df_side.columns = [str(c).strip() for c in df_side.columns]
        
        base_key = base_key.strip()
        side_key = side_key.strip()
        cols_to_add = [c.strip() for c in cols_to_add]

        # --- Validation ---
        if base_key not in df_base.columns:
            print(f"[ERROR] La columna clave '{base_key}' no existe en el archivo base. Columnas disponibles: {list(df_base.columns)}")
            return False
        if side_key not in df_side.columns:
            print(f"[ERROR] La columna clave '{side_key}' no existe en el archivo de enriquecimiento. Columnas disponibles: {list(df_side.columns)}")
            return False
        for col in cols_to_add:
            if col not in df_side.columns:
                print(f"[ERROR] La columna a agregar '{col}' no existe en el archivo de enriquecimiento. Columnas disponibles: {list(df_side.columns)}")
                return False

        # --- Merging ---
        print(f"Uniendo por '{base_key}' (base) y '{side_key}' (enriquecimiento)...")
        
        # Ensure keys are strings for reliable matching
        df_base[base_key] = df_base[base_key].astype(str).str.strip()
        df_side[side_key] = df_side[side_key].astype(str).str.strip()
        
        # Exclude the key column from the columns to be added
        if side_key in cols_to_add:
            cols_to_add.remove(side_key)

        # Keep only necessary columns from side file
        df_side_subset = df_side[[side_key] + cols_to_add].drop_duplicates(subset=[side_key])

        # Perform the merge
        df_enriched = pd.merge(
            df_base,
            df_side_subset,
            left_on=base_key,
            right_on=side_key,
            how='left'
        )

        # Remove the side key column if it's different from the base key
        if base_key != side_key and side_key in df_enriched.columns:
            df_enriched.drop(columns=[side_key], inplace=True)
        
        # --- Reporting ---
        original_rows = len(df_base)
        matched_rows = df_enriched[cols_to_add[0]].notna().sum() if cols_to_add else 0
        print(f"Se encontraron coincidencias para {matched_rows} de {original_rows} filas.")

        # --- Saving ---
        print(f"Guardando archivo enriquecido en: {output_path}")
        df_enriched.to_excel(output_path, index=False)
        print("[ÉXITO] El archivo ha sido enriquecido y guardado.")
        return True

    except Exception as e:
        print(f"[ERROR] Ocurrió un error inesperado durante el enriquecimiento: {e}")
        return False
