import sys
import os
import pandas as pd
import time

# Add the current directory to sys.path to allow importing Tx_MIGO3
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.append(current_dir)

from Tx_MIGO3 import SapMigoBotTurbo

class SapMigoTransferBot(SapMigoBotTurbo):
    def setup_transfer_header(self):
        """Configura la cabecera de MIGO para Traspaso 301"""
        print("--- Configurando Cabecera para Traspaso 301 ---")
        try:
            # A08 = Traspaso (Transfer Posting)
            self.session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0006/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").Key = "A08"
            time.sleep(0.5)
            
            # R10 = Otros (Others)
            self.session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0006/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-REFDOC").Key = "R10"
            time.sleep(0.5)
            
            # BWART = 301
            self.session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0006/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0112/ctxtGOHEAD-BWARTWA").Text = "301"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(1.0)
            
        except Exception as e:
            print(f"Error configurando cabecera: {e}")

    def run(self, excel_path):
        self.start_transaction()
        self.setup_transfer_header()
        
        # Read Excel using the parent class method or custom logic if needed
        # We use the parent's read_excel_dynamic but we need to map columns first
        df = self.read_excel_dynamic(excel_path)
        if df is None or df.empty:
            print("Error: DataFrame vacío o no se pudo leer.")
            return

        print("--- Mapeando Columnas de Excel (Formato Antiguo) a MIGO Bot ---")
        # Mapping based on user's VBScript:
        # Col 1 (A) -> Material
        # Col 11 (K) -> Cantidad
        # Col 10 (J) -> UME (Unidad)
        # Col 5 (E) -> Lote
        
        # We need to rename columns to match what Tx_MIGO3 expects:
        # "Material", "Cantidad", "Unidad", "Lote_Orig"
        # And add fixed columns: "Centro_Orig", "Alm_Orig", "Centro_Dest", "Alm_Dest", "Lote_Dest"
        
        # Since read_excel_dynamic reads headers, we assume the Excel HAS headers in row 2 (as per VBS comments)
        # The VBS says: "TITULOS DEL EXCEL TIENEN QUE ESTAR EN EL RENGLON 2" -> Row 2 is header, Data starts Row 3.
        # Tx_MIGO3.read_excel_dynamic reads the UsedRange. If Row 1 is empty and Row 2 has headers, it might pick up Row 2 as headers.
        # Let's inspect the DF columns to be sure, but for now we'll try to map by index if names don't match, 
        # OR we rely on the user having standard headers. 
        # The VBS used .Cells(i, 1) etc, implying positional reliance.
        
        # Let's try to map by position if possible, or just rename standard columns if they exist.
        # Since we can't easily know the exact header names the user has, we might need to rely on the user 
        # ensuring the Excel has headers that match, OR we map by index if the DF has enough columns.
        
        # Strategy: Create a new DF with the expected columns for Tx_MIGO3
        new_df = pd.DataFrame()
        
        # Helper to get column by index safely
        def get_col_by_index(df, idx): # idx is 0-based
            if idx < len(df.columns):
                return df.iloc[:, idx]
            return None

        # VBS indices were 1-based: 1=A, 5=E, 10=J, 11=K
        # Pandas indices are 0-based: 0=A, 4=E, 9=J, 10=K
        
        col_material = get_col_by_index(df, 0) # A
        col_lote = get_col_by_index(df, 4)     # E
        col_ume = get_col_by_index(df, 9)      # J
        col_cant = get_col_by_index(df, 10)    # K
        
        if col_material is not None: new_df["Material"] = col_material
        if col_cant is not None: new_df["Cantidad"] = col_cant
        if col_ume is not None: new_df["Unidad"] = col_ume
        if col_lote is not None: 
            new_df["Lote_Orig"] = col_lote
            new_df["Lote_Dest"] = col_lote # VBS: Lote_receptor = Lote
            
        # Hardcoded values from VBS
        new_df["Centro_Orig"] = "sgsj"
        new_df["Alm_Orig"] = "SDIF"
        new_df["Centro_Dest"] = "sgsj"
        new_df["Alm_Dest"] = "NCD1"
        new_df["Texto_Cabecera"] = "Ajuste ciclico SDIF"
        
        # Filter out rows where Quantity < 1 (as per VBS: If cantidad >= 1)
        # Convert to numeric for filtering
        try:
            new_df["Cantidad_Num"] = pd.to_numeric(new_df["Cantidad"], errors='coerce').fillna(0)
            new_df = new_df[new_df["Cantidad_Num"] > 0]
            new_df = new_df.drop(columns=["Cantidad_Num"])
        except:
            pass # If conversion fails, ignore filtering or let MIGO handle it
            
        print(f"--- Datos preparados: {len(new_df)} registros ---")
        
        # Now we have a DF compatible with Tx_MIGO3 logic
        # We can reuse the logic from run(), but we need to inject this DF.
        # Tx_MIGO3.run() calls read_excel_dynamic again. 
        # We will override the 'df' inside the class or modify how run works.
        # Actually, Tx_MIGO3.run() does: df = self.read_excel_dynamic(excel_path)
        # We can just override read_excel_dynamic to return our processed DF? 
        # No, simpler to copy the logic of run() here but use our new_df.
        
        self.process_dataframe(new_df)

    def process_dataframe(self, df):
        """Logic copied and adapted from Tx_MIGO3.run to work with a pre-loaded DataFrame"""
        self.table = self.find_migo_table()
        if self.table is None:
            print("Error: No se pudo encontrar la tabla MIGO")
            return
            
        self.map_columns()
        
        if 'Texto_Cabecera' in df.columns:
             header_val = df.iloc[0]['Texto_Cabecera']
             self.write_header(header_val)

        import math
        BLOCK_SIZE = 20
        total_rows = len(df)
        num_blocks = math.ceil(total_rows / BLOCK_SIZE)
        
        print(f"Iniciando carga: {total_rows} registros en {num_blocks} bloques.")
        current_sap_scroll = 0 

        for block_idx in range(num_blocks):
            start_idx = block_idx * BLOCK_SIZE
            end_idx = min((block_idx + 1) * BLOCK_SIZE, total_rows)
            print(f"\n--- Bloque {block_idx + 1} (Filas {start_idx} a {end_idx}) ---")
            block_df = df.iloc[start_idx:end_idx].reset_index(drop=True)
            
            # ... (Rest of the logic from Tx_MIGO3.run, simplified or called directly if refactored)
            # Since we can't easily call the 'middle' of a function, we duplicate the loop logic here.
            # It's better to duplicate than to modify the original file too heavily and break other bots.
            
            self.table = self.find_migo_table()
            try: self.table.VerticalScrollbar.Position = current_sap_scroll
            except: 
                self.table = self.find_migo_table()
                self.table.VerticalScrollbar.Position = current_sap_scroll
            time.sleep(0.2)
            self.table = self.find_migo_table()

            cols_phase1 = [
                ("MAT", "Material"), 
                ("QTY", "Cantidad"), 
                ("UNIT", "Unidad"), 
                ("LOC_O", "Alm_Orig"),
                ("PLANT_O", "Centro_Orig")
            ]

            print("   -> Llenando Datos Origen...")
            for key, excel_col in cols_phase1:
                if self.cols[key] != -1 and excel_col in block_df.columns:
                    for i in range(len(block_df)):
                        val = block_df.iloc[i][excel_col]
                        self.set_val_robust(self.cols[key], i, val)

            print("   -> Validando Origen...")
            self.session.findById("wnd[0]").sendVKey(0) 
            time.sleep(0.8)
            
            self.table = self.find_migo_table()
            try: 
                if self.table.VerticalScrollbar.Position != current_sap_scroll:
                    self.table.VerticalScrollbar.Position = current_sap_scroll
            except: pass

            cols_phase2 = [
                ("BATCH_O", "Lote_Orig"),
                ("PLANT_D", "Centro_Dest"),
                ("LOC_D", "Alm_Dest")
            ]

            print("   -> Llenando Lotes y Destinos...")
            has_dest = False
            for key, excel_col in cols_phase2:
                if self.cols[key] != -1 and excel_col in block_df.columns:
                    for i in range(len(block_df)):
                        val = block_df.iloc[i][excel_col]
                        self.set_val_robust(self.cols[key], i, val)
                        if key == "PLANT_D" and str(val).strip() != "":
                            has_dest = True

            if has_dest:
                print("   -> Validando Destinos...")
                self.session.findById("wnd[0]").sendVKey(0)
                time.sleep(0.8)
                self.table = self.find_migo_table()
                try: 
                    if self.table.VerticalScrollbar.Position != current_sap_scroll:
                        self.table.VerticalScrollbar.Position = current_sap_scroll
                except: pass

                print("   -> Llenando Lote Destino...")
                try: self.table.FirstVisibleColumn = 12 
                except: pass
                
                if "Lote_Dest" in block_df.columns and self.cols["BATCH_D"] != -1:
                    for i in range(len(block_df)):
                        val = block_df.iloc[i]["Lote_Dest"]
                        self.set_val_robust(self.cols["BATCH_D"], i, val)
                        
                self.session.findById("wnd[0]").sendVKey(0)
                time.sleep(0.6)
            else:
                 self.session.findById("wnd[0]").sendVKey(0)
                 time.sleep(0.6)

            current_sap_scroll += len(block_df)
            
        print("--- Carga finalizada ---")
        self.finalizar_interactivo()

if __name__ == "__main__":
    if len(sys.argv) > 1:
        excel_path = sys.argv[1]
        print(f"Iniciando Bot MIGO Traspaso 301 con archivo: {excel_path}")
        bot = SapMigoTransferBot()
        bot.run(excel_path)
        
        # Keep window open for user to see result
        input("\nPresione ENTER para cerrar esta ventana...")
    else:
        print("Error: Debes arrastrar un archivo Excel a este script o ejecutarlo desde la Macro.")
        input("Presione ENTER para salir...")
