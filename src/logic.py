import glob
import os
import re
import warnings

import pandas as pd

from config import settings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

NUMERIC_OUTPUT_COLUMNS = [
    "Cantidad",
    "Precio unitario eD",
    "Precio unitario cliente",
    "Precio total eD",
    "Precio total cliente",
]

CONSISTENCY_NUMERIC_COLUMNS = [
    "Cantidad",
    "Precio unitario eD",
    "Precio total eD",
]

CONSISTENCY_TOLERANCE = 0.01


def _normalize_mixed_number(value):
    """Normalize strings with mixed decimal separators into a float."""
    if pd.isna(value):
        return pd.NA

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return value

    text = str(value).strip()
    if not text:
        return pd.NA

    text = text.replace(" ", "")

    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "")
            text = text.replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        # If the string has more than one comma, they are treated as thousands separators and removed.
        # Otherwise, the single comma is treated as a decimal separator.
        if text.count(',') > 1:
            text = text.replace(",", "")
        else:
            text = text.replace(",", ".")
    elif text.count(".") > 1:
        # If the last part after a dot has 3 digits, and it's not the only dot,
        # it's likely a number with dots as thousands separators.
        parts = text.split('.')
        # Heuristic: if segments aren't all 3 digits, the last dot is likely a decimal point.
        is_thousands = all(len(p) == 3 for p in parts[1:-1]) and len(parts[-1]) == 3
        if is_thousands:
            text = text.replace(".", "")
        else:
            text = "".join(parts[:-1]) + "." + parts[-1]

    text = re.sub(r"[^0-9.\-+]", "", text)
    if text in {"", "-", "+", ".", "-.", "+."}:
        return pd.NA

    try:
        return float(text)
    except ValueError:
        return pd.NA


def _parse_number_variant(value, decimal_mode):
    if pd.isna(value):
        return pd.NA

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)

    text = str(value).strip()
    if not text:
        return pd.NA

    text = text.replace(" ", "")
    text = re.sub(r"[^0-9,.\-+]", "", text)
    if text in {"", "-", "+", ".", ",", "-.", "+.", "-,", "+,"}:
        return pd.NA

    if decimal_mode == "comma":
        text = text.replace(".", "")
        text = text.replace(",", ".")
    elif decimal_mode == "dot":
        text = text.replace(",", "")
    else:
        return pd.NA

    try:
        return float(text)
    except ValueError:
        return pd.NA


def _candidate_values_for_quantity(value):
    candidates = []
    for mode in ("comma", "dot"):
        parsed = _parse_number_variant(value, mode)
        if pd.notna(parsed):
            candidates.append((mode, float(parsed)))

    fallback = _normalize_mixed_number(value)
    if pd.notna(fallback):
        fallback_value = float(fallback)
        if not any(abs(candidate - fallback_value) <= 1e-12 for _, candidate in candidates):
            candidates.append(("mixed", fallback_value))

    return candidates


def _relative_error(expected, actual):
    scale = max(abs(expected), abs(actual), 1.0)
    return abs(expected - actual) / scale


def _choose_consistent_triplet(row, row_number=None):
    raw_qty = row.get("Cantidad", pd.NA)
    raw_unit = row.get("Precio unitario eD", pd.NA)
    raw_total = row.get("Precio total eD", pd.NA)
    messages = []

    # If 'total' is zero (common for uncalculated formulas), try to calculate it
    # from quantity and unit price. This avoids warnings and fixes the data.
    if pd.notna(raw_total) and str(raw_total).strip() in ('0', '0.0', '0,0', '0.00', '0,00'):
        updated_row = row.copy()
        cantidad = _normalize_mixed_number(raw_qty)
        unitario = _normalize_mixed_number(raw_unit)

        updated_row["Cantidad"] = cantidad
        updated_row["Precio unitario eD"] = unitario

        # If qty and unit price are valid numbers, calculate the total. Otherwise, it remains 0.
        if pd.notna(cantidad) and pd.notna(unitario):
            updated_row["Precio total eD"] = cantidad * unitario
        else:
            updated_row["Precio total eD"] = 0.0
        return updated_row, messages

    best = None
    for price_mode in ("comma", "dot"):
        unit = _parse_number_variant(raw_unit, price_mode)
        total = _parse_number_variant(raw_total, price_mode)
        if pd.isna(unit) or pd.isna(total):
            continue
        if abs(unit) <= 1e-12:
            continue

        expected_qty = total / unit
        qty_candidates = _candidate_values_for_quantity(raw_qty)
        for qty_mode, qty in qty_candidates:
            error = _relative_error(expected_qty, qty)
            candidate = {
                "price_mode": price_mode,
                "qty_mode": qty_mode,
                "Cantidad": qty,
                "Precio unitario eD": unit,
                "Precio total eD": total,
                "error": error,
                "expected_qty": expected_qty,
            }
            if best is None or candidate["error"] < best["error"]:
                best = candidate

    if best is None:
        return row, messages

    updated_row = row.copy()
    updated_row["Cantidad"] = best["expected_qty"]
    updated_row["Precio unitario eD"] = best["Precio unitario eD"]
    updated_row["Precio total eD"] = best["Precio total eD"]

    return updated_row, messages


def _resolve_numeric_consistency(df):
    required_cols = [col for col in CONSISTENCY_NUMERIC_COLUMNS if col in df.columns]
    if len(required_cols) < len(CONSISTENCY_NUMERIC_COLUMNS):
        return _normalize_numeric_columns(df), []

    resolved_rows = []
    messages = []
    for idx, (_, row) in enumerate(df.iterrows(), start=2):
        resolved_row, row_messages = _choose_consistent_triplet(row, row_number=idx)
        resolved_rows.append(resolved_row)
        if row_messages:
            messages.extend(row_messages)

    resolved_df = pd.DataFrame(resolved_rows, columns=df.columns)

    remaining_cols = [col for col in NUMERIC_OUTPUT_COLUMNS if col not in CONSISTENCY_NUMERIC_COLUMNS]
    resolved_df = _normalize_numeric_columns(resolved_df, remaining_cols)
    return resolved_df, messages


def _safe_resolve_numeric_consistency(df, context_label=""):
    messages = []
    try:
        resolved_df, consistency_messages = _resolve_numeric_consistency(df)
        if consistency_messages:
            messages.extend(consistency_messages)
        return resolved_df, messages
    except Exception as e:
        label = f" en {context_label}" if context_label else ""
        messages.append(
            f"[ADVERTENCIA] Fallo la heuristica numerica{label}: {e}. "
            "Se usara normalizacion simple."
        )
        return _normalize_numeric_columns(df.copy()), messages


def _normalize_numeric_columns(df, numeric_columns=None):
    numeric_columns = numeric_columns or NUMERIC_OUTPUT_COLUMNS
    for col in numeric_columns:
        if col in df.columns:
            df[col] = df[col].apply(_normalize_mixed_number)
    return df


def _format_number_with_comma(value):
    if pd.isna(value):
        return ""

    if not isinstance(value, (int, float)):
        return str(value)

    if isinstance(value, float) and value.is_integer():
        return str(int(value))

    text = f"{value:.15g}"
    return text.replace(".", ",")


def _apply_numeric_output_format(df, output_numeric_format, numeric_columns=None):
    numeric_columns = numeric_columns or NUMERIC_OUTPUT_COLUMNS
    if output_numeric_format == "comma_text":
        for col in numeric_columns:
            if col in df.columns:
                df[col] = df[col].apply(_format_number_with_comma)
    return df


def extract_table_from_file(file_path, lcl_name):
    """
    Finds the row where column B (index 1) is 'Tipo' and extracts data.
    Handles both Excel and CSV files.
    """
    found_rows = []
    errors = []
    expected_cols_count = 11
    file_ext = os.path.splitext(file_path)[1].lower()

    try:
        if file_ext == ".csv":
            try:
                df = pd.read_csv(file_path, header=None, encoding="utf-8", sep=";")
            except (UnicodeDecodeError, pd.errors.ParserError):
                try:
                    df = pd.read_csv(file_path, header=None, encoding="latin1", sep=";")
                except pd.errors.ParserError:
                    try:
                        df = pd.read_csv(file_path, header=None, encoding="utf-8", sep=",")
                    except Exception as e:
                        errors.append(
                            f"Warning: Could not parse CSV with any common format: {e}"
                        )
                        return [], errors

            if df.shape[1] < 2:
                return [], errors

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

                        raw_data = df.iloc[current_row, 1 : 1 + expected_cols_count].tolist()
                        found_rows.append([lcl_name] + raw_data)
                        current_row += 1
                    i = current_row
                    continue
                i += 1

        elif file_ext in [".xlsx", ".xls"]:
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

                                raw_data = df.iloc[current_row, 1 : 1 + expected_cols_count].tolist()
                                found_rows.append([lcl_name] + raw_data)
                                current_row += 1
                            i = current_row
                            continue
                        i += 1

                except Exception as e:
                    errors.append(
                        f"Warning: Error reading sheet '{sheet_name}': {e}"
                    )

    except Exception as e:
        errors.append(f"Error reading file: {e}")
        return [], errors

    return found_rows, errors


def process_stage1_by_subfolders(input_dir, output_dir, output_numeric_format="excel"):
    """
    Stage 1: Find subfolders in input_dir and process each as a batch.
    Returns a list of successfully processed folder paths for cleaning.
    """
    subfolders = [f.path for f in os.scandir(input_dir) if f.is_dir()]
    stage1_log = []
    relative_error_log = []

    if not subfolders:
        print(f"No se encontraron subcarpetas de lotes en {input_dir}.")
        return [], False

    print(f"Se encontraron {len(subfolders)} subcarpetas para procesar.")
    overall_success = False
    successfully_processed_folders = []

    for folder_path in subfolders:
        batch_name = os.path.basename(folder_path)
        print(f"\n--- Procesando Lote: {batch_name} ---")

        input_files = (
            glob.glob(os.path.join(folder_path, "*.xlsx"))
            + glob.glob(os.path.join(folder_path, "*.xls"))
            + glob.glob(os.path.join(folder_path, "*.csv"))
        )

        if not input_files:
            print(f" [!] No se encontraron archivos Excel o CSV en la carpeta '{batch_name}'.")
            continue

        print(f"Procesando {len(input_files)} archivos para el lote '{batch_name}'...")

        all_rows = []
        for file_path in input_files:
            filename = os.path.basename(file_path)
            name_no_ext = os.path.splitext(filename)[0]

            if "lcl" in name_no_ext.lower():
                numbers = re.findall(r"\d+", name_no_ext)
                if numbers:
                    lcl_name = max(numbers, key=len)
                else:
                    lcl_name = name_no_ext.split("-")[0].strip()
            else:
                lcl_name = name_no_ext.split("-")[0].strip()

            rows, file_errors = extract_table_from_file(file_path, lcl_name)
            if file_errors:
                for error in file_errors:
                    stage1_log.append(f"[Lote: {batch_name} | Archivo: {filename}] {error}")

            if rows:
                # Add filename to each row for better error tracking
                for r in rows:
                    r.insert(0, filename)
                all_rows.extend(rows)
                print(f" [OK] {filename} -> {len(rows)} filas (LCL: {lcl_name})")
            else:
                msg = f"No se encontro tabla valida en '{filename}'"
                print(f" [!] {msg}")
                if not file_errors:
                    stage1_log.append(f"[Lote: {batch_name}] {msg}")

        if not all_rows:
            print(f"No se extrajeron datos de ningun archivo en el lote '{batch_name}'.")
            continue

        headers = [
            "Archivo_Origen",
            "LCL_Origen",
            "Tipo",
            "Contador",
            "Mat./Prest.",
            "Descripcion Mat./Serv.",
            "Cantidad",
            "Unidad medida base",
            "Imputacion",
            "Precio unitario eD",
            "Precio unitario cliente",
            "Precio total eD",
            "Precio total cliente",
        ]

        df_batch = pd.DataFrame(all_rows, columns=headers)

        output_path = os.path.join(output_dir, f"{batch_name}.xlsx")
        try:
            print("Normalizando columnas numericas con heuristica por fila...")
            context_label = f"lote {batch_name}"
            df_batch, numeric_messages = _safe_resolve_numeric_consistency(
                df_batch, context_label=context_label
            )

            if numeric_messages:
                for msg in numeric_messages:
                    log_entry = f"[Lote: {batch_name}] {msg}"
                    match = re.search(r"Fila (\d+):", msg)
                    if match:
                        row_number = int(match.group(1))
                        df_index = row_number - 2
                        if 0 <= df_index < len(df_batch):
                            source_file = df_batch.iloc[df_index]["Archivo_Origen"]
                            log_entry = f"[Archivo: {source_file}] {msg}"

                    if "No se pudo reconciliar" in msg:
                        relative_error_log.append(log_entry)
                    else:
                        stage1_log.append(log_entry)

            df_to_save = df_batch.drop(columns=["Archivo_Origen"])
            df_to_save = _apply_numeric_output_format(df_to_save, output_numeric_format)
            df_to_save.to_excel(output_path, index=False)

            print(f"\n[EXITO] Lote '{batch_name}' guardado en: {output_path}")
            print(f"Total filas: {len(df_batch)}")
            overall_success = True
            successfully_processed_folders.append(folder_path)
        except Exception as e:
            error_msg = f"[ERROR] No se pudo guardar el lote '{batch_name}': {e}"
            print(f"\n{error_msg}")
            stage1_log.append(error_msg)

    if stage1_log:
        print("\n\n--- Resumen de Otras Advertencias de la Fase 1 ---")
        for msg in stage1_log:
            print(f" - {msg}")
        print("--- Fin del Resumen ---")

    if relative_error_log:
        print("\n\n--- Resumen de Errores de Reconciliacion (error relativo) ---")
        for msg in sorted(relative_error_log):
            print(f" - {msg}")
        print("--- Fin del Resumen ---")

    return successfully_processed_folders, overall_success


def process_stage2_consolidation(
    input_dir,
    output_file_path,
    filter_column=None,
    allowed_values=None,
    output_numeric_format="excel",
):
    """
    Stage 2: Consolidate all files from the input directory with an optional
    filter by column and allowed values.
    """
    input_files = glob.glob(os.path.join(input_dir, "*.xlsx"))

    if not input_files:
        print(f"No hay archivos procesados en {input_dir}.")
        return False

    print("--- Unificando Lotes ---")
    dfs = []
    allowed_values_set = {str(value).strip() for value in (allowed_values or []) if str(value).strip()}

    for file_path in input_files:
        try:
            df = pd.read_excel(file_path)

            if filter_column and allowed_values_set:
                if filter_column in df.columns:
                    original_rows = len(df)
                    comparable_series = df[filter_column].astype(str).str.strip()
                    df = df[comparable_series.isin(allowed_values_set)]
                    print(
                        f" [FILTRO] Lote '{os.path.basename(file_path)}': {len(df)} de "
                        f"{original_rows} filas cumplen {filter_column} en {sorted(allowed_values_set)}."
                    )
                else:
                    print(
                        f" [ADVERTENCIA] La columna '{filter_column}' no se encontro "
                        f"en '{os.path.basename(file_path)}', no se pudo filtrar."
                    )

            batch_name = os.path.splitext(os.path.basename(file_path))[0]
            df.insert(0, "Tipo_Obra", batch_name)

            dfs.append(df)
            print(f" [OK] Leido lote: {batch_name} ({len(df)} filas)")
        except Exception as e:
            print(f" [Error] Fallo al leer {os.path.basename(file_path)}: {e}")

    if not dfs:
        print("No se encontraron datos para consolidar (despues de filtrar).")
        return False

    df_consolidated = pd.concat(dfs, ignore_index=True)
    df_consolidated, _ = _safe_resolve_numeric_consistency(
        df_consolidated, context_label="consolidado"
    )
    df_consolidated = _apply_numeric_output_format(df_consolidated, output_numeric_format)
    print(f"\nTotal consolidado: {len(df_consolidated)} filas.")

    try:
        df_consolidated.to_excel(output_file_path, index=False)
        print(f"\n[EXITO] Archivo final guardado en: {output_file_path}")
        return True
    except Exception as e:
        print(f"[ERROR] Fallo al guardar el archivo consolidado: {e}")
        return False


def _read_file_to_df(file_path):
    """Reads an Excel or CSV file into a pandas DataFrame."""
    if not file_path or not os.path.exists(file_path):
        return None

    file_ext = os.path.splitext(file_path)[1].lower()
    try:
        if file_ext in [".xlsx", ".xls"]:
            return pd.read_excel(file_path)
        if file_ext == ".csv":
            try:
                return pd.read_csv(file_path, sep=",")
            except Exception:
                return pd.read_csv(file_path, sep=";", encoding="latin1")
    except Exception as e:
        print(f"Error al leer el archivo {os.path.basename(file_path)}: {e}")
        return None
    return None


def enrich_file(
    base_path,
    side_path,
    base_key,
    side_key,
    cols_to_add,
    cols_to_drop,
    output_path,
    output_numeric_format="excel",
):
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

        df_base.columns = [str(c).strip() for c in df_base.columns]
        df_side.columns = [str(c).strip() for c in df_side.columns]
        df_base, _ = _safe_resolve_numeric_consistency(df_base, context_label="archivo base")
        df_side, _ = _safe_resolve_numeric_consistency(
            df_side, context_label="archivo enriquecimiento"
        )

        base_key = base_key.strip()
        side_key = side_key.strip()
        cols_to_add = [c.strip() for c in cols_to_add]
        cols_to_drop = [c.strip() for c in (cols_to_drop or [])]

        if base_key not in df_base.columns:
            print(
                f"[ERROR] La columna clave '{base_key}' no existe en el archivo base. "
                f"Columnas disponibles: {list(df_base.columns)}"
            )
            return False
        if side_key not in df_side.columns:
            print(
                f"[ERROR] La columna clave '{side_key}' no existe en el archivo de enriquecimiento. "
                f"Columnas disponibles: {list(df_side.columns)}"
            )
            return False
        for col in cols_to_add:
            if col not in df_side.columns:
                print(
                    f"[ERROR] La columna a agregar '{col}' no existe en el archivo de "
                    f"enriquecimiento. Columnas disponibles: {list(df_side.columns)}"
                )
                return False
        for col in cols_to_drop:
            if col not in df_base.columns:
                print(
                    f"[ERROR] La columna a quitar '{col}' no existe en el archivo base. "
                    f"Columnas disponibles: {list(df_base.columns)}"
                )
                return False

        print(f"Uniendo por '{base_key}' (base) y '{side_key}' (enriquecimiento)...")

        df_base[base_key] = df_base[base_key].astype(str).str.strip()
        df_side[side_key] = df_side[side_key].astype(str).str.strip()

        if side_key in cols_to_add:
            cols_to_add.remove(side_key)

        df_side_subset = df_side[[side_key] + cols_to_add].drop_duplicates(subset=[side_key])

        df_enriched = pd.merge(
            df_base,
            df_side_subset,
            left_on=base_key,
            right_on=side_key,
            how="left",
        )

        if base_key != side_key and side_key in df_enriched.columns:
            df_enriched.drop(columns=[side_key], inplace=True)

        removable_cols = [col for col in cols_to_drop if col in df_enriched.columns]
        if removable_cols:
            df_enriched.drop(columns=removable_cols, inplace=True)
            print(f"Se quitaron {len(removable_cols)} columnas del resultado: {removable_cols}")

        df_enriched = _apply_numeric_output_format(df_enriched, output_numeric_format)

        original_rows = len(df_base)
        matched_rows = df_enriched[cols_to_add[0]].notna().sum() if cols_to_add else 0
        print(f"Se encontraron coincidencias para {matched_rows} de {original_rows} filas.")

        print(f"Guardando archivo enriquecido en: {output_path}")
        df_enriched.to_excel(output_path, index=False)
        print("[EXITO] El archivo ha sido enriquecido y guardado.")
        return True

    except Exception as e:
        print(f"[ERROR] Ocurrio un error inesperado durante el enriquecimiento: {e}")
        return False
