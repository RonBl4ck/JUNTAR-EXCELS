import os

# Base directory
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Data directories
DATA_DIR = os.path.join(BASE_DIR, "data")
STAGE1_RAW_DIR = os.path.join(DATA_DIR, "stage1_raw")
STAGE2_PROCESSED_DIR = os.path.join(DATA_DIR, "stage2_processed_work_types")
OUTPUT_DIR = os.path.join(DATA_DIR, "output_final")
MASTER_DIR = os.path.join(DATA_DIR, "master")

# File paths
MASTER_FILE = os.path.join(MASTER_DIR, "Maestra_LCL.xlsx")
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "MASTER_ENRIQUECIDO.xlsx")

# Master file columns
MASTER_KEY_COL = "LCL"
PROCESS_KEY_COL = "LCL_Origen"

# Columns to extract from Master
MASTER_COLS_TO_ADD = ['Zona', 'Distrito', 'Tipo de Obra', 'Origen']
