import os
import sys
import pandas as pd

# Path to the project
project_root = r"c:\Users\P723919021\Documents\vscode\JUNTAR EXCELS"
sys.path.append(os.path.join(project_root, "src"))

import logic

def test_user_row():
    # User's case:
    # Cantidad: 15.6
    # Unit: 684,36
    # Total: 10676,00
    
    row_data = {
        "Cantidad": "15.6",
        "Precio unitario eD": "684,36",
        "Precio total eD": "10676,00",
        "Other": "Value"
    }
    
    # Simulate the logic pipeline
    df = pd.DataFrame([row_data])
    print("--- Datos Originales ---")
    print(df)
    
    resolved_df, messages = logic._safe_resolve_numeric_consistency(df, "test_context")
    
    print("\n--- Datos Procesados ---")
    print(resolved_df)
    
    print("\n--- Mensajes ---")
    for msg in messages:
        print(msg)

    # Verification
    final_total = resolved_df.iloc[0]["Precio total eD"]
    final_unit = resolved_df.iloc[0]["Precio unitario eD"]
    final_qty = resolved_df.iloc[0]["Cantidad"]
    
    expected_total = 10676.01 # Since 15.6 * 684.36 = 10676.016, and we round(n,4) we expect 10676.016 (or .02)
    # Actually 15.6 * 684.36 = 10676.016
    
    if final_total < 100000:
        print("\n[SUCCESS] El total no fue inflado (se mantiene cerca de 10676).")
    else:
        print("\n[FAILURE] El total fue inflado a", final_total)

    if final_qty == 15.6 or abs(final_qty - 15.6) < 0.001:
        print("[SUCCESS] La cantidad se mantiene correcta (redondeada).")
    else:
        print("[FAILURE] La cantidad tiene error de precision:", final_qty)

if __name__ == "__main__":
    test_user_row()
