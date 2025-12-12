import pandas as pd
import os

files = [
    "Consolidado_Facturas.xlsx",
    "Maestro_Pasillos.xlsx",
    "Perdida_Vacio_MIGO.xlsx"
]

base_path = r"c:\Users\ariel.mella\OneDrive - CIAL Alimentos\Escritorio\respaldo\Antigravity\Apps\Nexus_Jarvis"

print("--- EXCEL FILE INSPECTION ---")
for f in files:
    path = os.path.join(base_path, f)
    if os.path.exists(path):
        try:
            df = pd.read_excel(path, nrows=5)
            print(f"\nFILE: {f}")
            print(f"COLUMNS: {list(df.columns)}")
            print(f"SAMPLE DATA:\n{df.head(2).to_string()}")
        except Exception as e:
            print(f"\nFILE: {f} - ERROR: {e}")
    else:
        print(f"\nFILE: {f} - NOT FOUND")
