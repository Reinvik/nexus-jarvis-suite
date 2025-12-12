import pandas as pd
import os

OUTPUT_DIR = "Nexus_Jarvis_Build_v5/Entregable_Usuarios/Plantillas"

if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)

# 1. Plantilla MIGO
df_migo = pd.DataFrame(columns=[
    "Material", "Cantidad", "Unidad", "Centro_Orig", "Alm_Orig", "Lote_Orig", 
    "Centro_Dest", "Alm_Dest", "Lote_Dest", "Texto_Cabecera"
])
# Ejemplo
df_migo.loc[0] = ["1000123", "10", "UN", "SGSJ", "NCD1", "L001", "SGSJ", "NCD2", "", "Traspaso Manual"]
df_migo.to_excel(os.path.join(OUTPUT_DIR, "Plantilla_MIGO.xlsx"), index=False)

# 2. Plantilla LT01
df_lt01 = pd.DataFrame(columns=[
    "Material", "Cantidad", "Unidad", "Alm_Dest", "Ubicación"
])
# Ejemplo
df_lt01.loc[0] = ["1000123", "50", "UN", "920", "TRANSFER"]
df_lt01.to_excel(os.path.join(OUTPUT_DIR, "Plantilla_LT01.xlsx"), index=False)

# 3. Plantilla Auditor Altura (Pallet)
# Este bot pega la data, así que dejamos una hoja vacía o con cabecera simple
df_pallet = pd.DataFrame(columns=["Pegar LX02 Aquí"])
df_pallet.to_excel(os.path.join(OUTPUT_DIR, "Plantilla_Auditor_Altura.xlsx"), index=False)

# 4. Plantilla Conversiones
df_conv = pd.DataFrame(columns=["Material"])
df_conv.loc[0] = ["1000123"]
df_conv.to_excel(os.path.join(OUTPUT_DIR, "Plantilla_Conversiones.xlsx"), index=False)

print(f"✅ Plantillas creadas en {OUTPUT_DIR}")
