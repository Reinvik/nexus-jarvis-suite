import pandas as pd
import os
from datetime import datetime
import numpy as np

# --- CONFIGURATION ---
INPUT_FILE = "Perdida_Vacio_MIGO.xlsx"
OUTPUT_FILE = "Dashboard_Source_Mermas.csv"
BASE_PATH = os.path.dirname(os.path.abspath(__file__))
# ---------------------

def prepare_data():
    input_path = os.path.join(BASE_PATH, "..", INPUT_FILE)
    output_path = os.path.join(BASE_PATH, "..", OUTPUT_FILE)

    print(f"Reading file: {input_path}")
    
    if not os.path.exists(input_path):
        print("ERROR: Input file not found.")
        return

    try:
        # Read Excel
        df = pd.read_excel(input_path)
        
        # --- CLEANING & ENRICHMENT ---
        
        # 1. Add Report Date (Simulated for this demo, usually would be file date)
        # We use today's date so the dashboard always looks "fresh"
        df['Fecha_Reporte'] = datetime.now().strftime("%Y-%m-%d")
        
        # 2. Add 'Total_Impacto_Estimado' (Simulating a cost if not present, or just Quantity)
        # If 'Cantidad' exists, ensure it's numeric
        if 'Cantidad' in df.columns:
            df['Cantidad'] = pd.to_numeric(df['Cantidad'], errors='coerce').fillna(0)
        
        # 3. Add 'Semana' for trends
        df['Semana'] = datetime.now().isocalendar()[1]
        
        # 4. Clean text fields
        if 'Nombre' in df.columns:
            df['Nombre'] = df['Nombre'].str.strip()
            
        # 5. Add a 'Categoria_Simulada' for better visuals (Optional)
        # Simple logic based on name to create groups
        def categorize(name):
            name = str(name).upper()
            if 'VIENESA' in name: return 'Embutidos'
            if 'JAMON' in name: return 'Fiambres'
            if 'PATE' in name: return 'Untables'
            return 'Otros'
            
        if 'Nombre' in df.columns:
            df['Categoria'] = df['Nombre'].apply(categorize)

        # --- EXPORT TO CSV ---
        # CSV is faster for Power BI to read and doesn't lock like Excel
        df.to_csv(output_path, index=False, encoding='utf-8-sig', sep=';')
        
        print(f"OK - Data prepared successfully!")
        print(f"Output: {output_path}")
        print(f"Rows processed: {len(df)}")
        print("\nPreparation Complete. You can now load this CSV into Power BI.")

    except Exception as e:
        print(f"ERROR - Error preparing data: {e}")

if __name__ == "__main__":
    prepare_data()
