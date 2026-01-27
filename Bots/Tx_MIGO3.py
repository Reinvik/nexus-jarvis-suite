import win32com.client
import sys
import time
import pandas as pd
import os
import math
import shutil
import re 
from datetime import datetime
import pythoncom
import ctypes # [NEW]

class SapMigoBotTurbo:
    def __init__(self):
        self.session = None
        self.table = None
        self.cols = {} 
        self.connect_to_sap()

    def connect_to_sap(self, max_retries=3):
        """Conexión robusta a SAP con reintentos"""
        try:
            pythoncom.CoInitialize()
        except:
            pass
        
        for attempt in range(max_retries):
            try:
                sap_gui = win32com.client.GetObject("SAPGUI")
                application = sap_gui.GetScriptingEngine
                connection = application.Children(0)
                self.session = connection.Children(0)
                
                # Verificar que la sesión está activa
                _ = self.session.findById("wnd[0]")
                print("--- Conectado a SAP exitosamente ---")
                return True
                
            except Exception as e:
                print(f"   Intento {attempt + 1}/{max_retries} fallido: {e}")
                self.session = None
                try:
                    pythoncom.CoUninitialize()
                    time.sleep(1)
                    pythoncom.CoInitialize()
                except:
                    pass
                
                if attempt < max_retries - 1:
                    print(f"   Reintentando en 2 segundos...")
                    time.sleep(2)
        
        print("Error conectando a SAP. Asegúrate de tener SAP abierto y logueado.")
        return False

    def reconnect_if_needed(self):
        """Reconecta a SAP si la sesión se perdió"""
        try:
            _ = self.session.findById("wnd[0]")
            return True
        except:
            print("⚠️ Sesión SAP perdida. Reconectando...")
            return self.connect_to_sap()

    def start_transaction(self):
        print("--- Iniciando Transacción MIGO ---")
        if not self.reconnect_if_needed():
            raise Exception("No se pudo conectar a SAP")
        
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nMIGO"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(1.0) # Optimizado (antes 2.0)
            try: self.session.findById("wnd[0]").maximize()
            except: pass
            try:
                self.session.findById("wnd[0]/shellcont/shell/shellcont[2]/shell/shellcont[0]/shell").pressButton("CLOSE")
            except: pass
        except Exception as e:
            print(f"Advertencia al abrir MIGO: {e}")

    def clean_data(self, df):
        df = df.astype(str)
        def remove_trailing_zero(val):
            val = val.strip()
            if val.endswith(".0"):
                return val[:-2]
            return val
        for col in df.columns:
            df[col] = df[col].apply(remove_trailing_zero)
            df[col] = df[col].replace({'nan': '', 'None': '', 'NaT': ''})
            df[col] = df[col].fillna('')
        return df

    def read_excel_dynamic(self, excel_path):
        filename = os.path.basename(excel_path)
        try:
            excel_app = win32com.client.GetActiveObject("Excel.Application")
            wb_found = None
            for wb in excel_app.Workbooks:
                if wb.Name == filename:
                    wb_found = wb
                    break
            if wb_found:
                print("   [INFO] Leyendo desde Excel ABIERTO (En vivo)...")
                ws = wb_found.Worksheets(1)
                used_range = ws.UsedRange.Value
                if used_range:
                    headers = list(used_range[0])
                    data = used_range[1:]
                    df = pd.DataFrame(data, columns=headers)
                    df = df.loc[:, ~df.columns.duplicated()]
                    return self.clean_data(df)
        except Exception:
            pass

        print("   [INFO] Leyendo desde archivo en disco...")
        temp_file = excel_path + ".temp_bot.xlsx"
        try:
            shutil.copyfile(excel_path, temp_file)
            df = pd.read_excel(temp_file, dtype=str)
            df = df.loc[:, ~df.columns.duplicated()]
            df = df.fillna("")
            return self.clean_data(df)
        except Exception as e:
            print(f"[ERROR] No se pudo leer el Excel: {e}")
            return None
        finally:
            if os.path.exists(temp_file):
                try: os.remove(temp_file)
                except: pass

    def find_migo_table(self):
        possible_paths = [
            "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0006/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM",
            "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM",
            "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0005/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM",
            "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0008/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM"
        ]
        start_time = time.time()
        while time.time() - start_time < 5:
            for path in possible_paths:
                try:
                    table = self.session.findById(path)
                    _ = table.RowCount 
                    return table
                except: continue
            time.sleep(0.2)
        print("Error Crítico: No se encontró la tabla. Verifique que MIGO esté abierto y en la pestaña correcta.")
        return None

    def map_columns(self):
        print("--- Mapeando columnas según variante ---")
        search_keys = {
            "MAT": ["MATNR", "MAKTX"], "QTY": ["ERFMG"], "UNIT": ["ERFME"],
            "PLANT_O": ["WERKS", "NAME1"], "LOC_O": ["LGORT", "LGOBE"], "BATCH_O": ["CHARG"],
            "PLANT_D": ["UMWRK", "UMNAME1"], "LOC_D": ["UMLGO", "UMLGOBE"], "BATCH_D": ["UMCHA"]
        }
        self.cols = {k: -1 for k in search_keys}
        try:
            for i in range(self.table.Columns.Count):
                name = self.table.Columns.Item(i).Name
                for key, possibilities in search_keys.items():
                    if self.cols[key] == -1:
                        for p in possibilities:
                            if p in name:
                                self.cols[key] = i
                                break
        except: pass
        
    def write_header(self, text):
        if pd.isna(text) or str(text).strip() == "": return
        try:
            self.session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0006/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0112/txtGOHEAD-BKTXT").Text = str(text)
        except: pass

    def set_val_robust(self, col_idx, row_vis, val):
        if hasattr(val, 'iloc'): val = val.iloc[0]
        if pd.isna(val) or str(val).strip() == "": return
        if col_idx == -1: return

        for _ in range(2):
            try:
                if col_idx >= 15:
                    try:
                        if self.table.FirstVisibleColumn < (col_idx - 2):
                            self.table.FirstVisibleColumn = col_idx - 2
                    except: pass
                else:
                    try:
                        if self.table.FirstVisibleColumn > 0:
                            self.table.FirstVisibleColumn = 0
                    except: pass
                
                cell = self.table.GetCell(row_vis, col_idx)
                if cell.Changeable:
                    cell.Text = str(val)
                    return 
            except:
                try: self.table = self.find_migo_table()
                except: pass

    def ir_a_mb51(self, doc_material):
        print(f"Detectado Documento: {doc_material}. Yendo a MB51...")
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nMB51"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(1.0) # Optimizado
            
            self.session.findById("wnd[0]/usr/ctxtMBLNR-LOW").text = doc_material
            self.session.findById("wnd[0]/usr/txtMJAHR-LOW").text = str(datetime.now().year)
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            print("--- MB51 Ejecutada ---")
        except Exception as e:
            print(f"Error llenando campos de MB51: {e}")

    def finalizar_interactivo(self):
        print("\n--- Esperando confirmación manual del usuario ---")
        
        MB_YESNO = 4
        MB_ICONQUESTION = 0x20
        MB_TOPMOST = 0x40000
        
        msg = (
            "El bot ha terminado de llenar los datos.\n\n"
            "1. Revise la información en SAP.\n"
            "2. Presione 'Contabilizar' (Guardar) manualmente en SAP.\n"
            "3. Una vez aparezca el mensaje 'Documento material ... contabilizado', haga clic en SÍ aquí para ir a MB51."
        )
        
        ret = ctypes.windll.user32.MessageBoxW(0, msg, "Acción Requerida - MIGO Bot", MB_YESNO | MB_ICONQUESTION | MB_TOPMOST)
        
        if ret == 6: # YES
            print("Usuario confirmó guardado. Leyendo mensaje SAP...")
            time.sleep(0.5)
            try:
                mensaje = self.session.findById("wnd[0]/sbar").Text
                print(f"Mensaje SAP detectado: {mensaje}")
                
                numeros = re.findall(r'\d+', mensaje)
                doc_material = ""
                for num in numeros:
                    if len(num) >= 9:
                        doc_material = num
                        break
                
                if doc_material:
                     self.ir_a_mb51(doc_material)
                else:
                    print("No se detectó número de documento en la barra de estado. (Tal vez ya desapareció o no se guardó).")
            except Exception as e:
                print(f"Error leyendo estado SAP: {e}")
        else:
            print("Usuario canceló o indicó que no guardó.")

    def run(self, excel_path):
        self.start_transaction()
        df = self.read_excel_dynamic(excel_path)
        if df is None or df.empty: return

        self.table = self.find_migo_table()
        if self.table is None:
            print("Error: No se pudo encontrar la tabla MIGO")
            return
            
        self.map_columns()
        df.columns = df.columns.str.strip()
        
        if 'Texto_Cabecera' in df.columns:
             header_val = df.iloc[0]['Texto_Cabecera']
             if hasattr(header_val, 'iloc'): header_val = header_val.iloc[0]
             self.write_header(header_val)

        BLOCK_SIZE = 20
        total_rows = len(df)
        num_blocks = math.ceil(total_rows / BLOCK_SIZE)
        
        print(f"Iniciando carga: {total_rows} registros en {num_blocks} bloques de {BLOCK_SIZE}.")
        current_sap_scroll = 0 

        for block_idx in range(num_blocks):
            start_idx = block_idx * BLOCK_SIZE
            end_idx = min((block_idx + 1) * BLOCK_SIZE, total_rows)
            print(f"\n--- Bloque {block_idx + 1} (Filas {start_idx} a {end_idx}) ---")
            block_df = df.iloc[start_idx:end_idx].reset_index(drop=True)
            
            self.table = self.find_migo_table()
            try: self.table.VerticalScrollbar.Position = current_sap_scroll
            except: 
                self.table = self.find_migo_table()
                self.table.VerticalScrollbar.Position = current_sap_scroll
            time.sleep(0.2) # Opt
            self.table = self.find_migo_table()

            cols_phase1 = [
                ("MAT", "Material"), 
                ("QTY", "Cantidad"), 
                ("UNIT", "Unidad"), 
                ("LOC_O", "Alm_Orig"),
                ("PLANT_O", "Centro_Orig")
            ]

            print("   -> Llenando Datos Origen (Por Columnas)...")
            
            for key, excel_col in cols_phase1:
                if self.cols[key] != -1 and excel_col in block_df.columns:
                    for i in range(len(block_df)):
                        val = block_df.iloc[i][excel_col]
                        self.set_val_robust(self.cols[key], i, val)

            print("   -> Validando Origen...")
            self.session.findById("wnd[0]").sendVKey(0) 
            time.sleep(0.8) # Opt
            
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

            print("   -> Llenando Lotes y Destinos (Por Columnas)...")
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
                time.sleep(0.8) # Opt
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
                time.sleep(0.6) # Opt
            else:
                 self.session.findById("wnd[0]").sendVKey(0)
                 time.sleep(0.6) # Opt

            current_sap_scroll += len(block_df)
            
        print("--- Carga finalizada ---")
        self.finalizar_interactivo()

if __name__ == "__main__":
    carpeta = r"c:\Users\ariel.mella\OneDrive - CIAL Alimentos\Archivos de Operación  Outbound CD - 16.-Inventario Critico"
    nombre_archivo = "carga_migo.xlsx" 
    archivo_excel = os.path.join(carpeta, nombre_archivo)

    if os.path.exists(archivo_excel):
        bot = SapMigoBotTurbo()
        bot.run(archivo_excel)
    else:
        print(f"No encuentro el archivo: {archivo_excel}")