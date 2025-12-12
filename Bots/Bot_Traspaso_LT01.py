import win32com.client
import sys
import time
import pandas as pd
import os
import pyperclip
import pythoncom
import shutil
from datetime import datetime

class SapBotTraspasoLT01:
    def __init__(self):
        self.session = None
        # Inicializar COM para este hilo
        try:
            pythoncom.CoInitialize()
        except:
            pass
        self.connect_to_sap()

    def connect_to_sap(self):
        try:
            # Intentar obtener el objeto SAPGUI
            sap_gui = win32com.client.GetObject("SAPGUI")
            if not sap_gui:
                print("No se encontró SAPGUI. Intentando Dispatch...")
                sap_gui = win32com.client.Dispatch("SAPGUI")
            
            application = sap_gui.GetScriptingEngine
            connection = application.Children(0)
            self.session = connection.Children(0)
            print("--- Conectado a SAP exitosamente ---")
        except Exception as e:
            print(f"Error conectando a SAP: {e}")
            print("Asegúrate de tener SAP abierto y logueado.")
            # No hacer sys.exit aquí para no matar el thread silenciosamente sin log
            # Dejar que falle más adelante o manejarlo en run


    def clean_float(self, val):
        try:
            val = str(val).replace(',', '.').strip()
            return float(val)
        except:
            return 0.0
    
    def clean_value(self, val):
        """Limpia valores de Excel eliminando .0 de números enteros"""
        if val is None:
            return ""
        val_str = str(val).strip()
        # Si termina en .0, quitarlo
        if val_str.endswith('.0'):
            val_str = val_str[:-2]
        return val_str
    
    def format_ubicacion(self, ubic):
        """Formatea ubicación a 7 caracteres con ceros a la izquierda"""
        if not ubic or ubic.upper() in ["TRANSFER", "SCHROTT"]:
            return ubic  # Ubicaciones especiales no se formatean
        
        # Limpiar y quitar espacios
        ubic_clean = str(ubic).strip()
        
        # Si es numérico, rellenar con ceros a la izquierda hasta 7 caracteres
        if ubic_clean.isdigit():
            return ubic_clean.zfill(7)
        
        return ubic_clean

    def descargar_stock_lx02(self, excel_path):
        print("--- Descargando Stock desde LX02 (Método Pallet) ---")
        
        try:
            # Ir a LX02
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nLX02"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(1.0)
            
            # Filtros (IDs corregidos S1)
            self.session.findById("wnd[0]/usr/ctxtS1_LGNUM").text = "NCD"
            
            # Selección Múltiple de Tipos de Almacén
            tipos_almacen = ["PCG", "PF1", "PPB"]
            print(f"   -> Consultando Tipos: {', '.join(tipos_almacen)}")
            
            self.session.findById("wnd[0]/usr/btn%_S1_LGTYP_%_APP_%-VALU_PUSH").press()
            time.sleep(0.5)
            
            # Pegar tipos desde portapapeles
            pyperclip.copy("\n".join(tipos_almacen))
            self.session.findById("wnd[1]/tbar[0]/btn[24]").press() 
            time.sleep(0.5)
            self.session.findById("wnd[1]/tbar[0]/btn[8]").press()  
            time.sleep(0.5)
            
            # Ejecutar Reporte
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            time.sleep(2.0)
            
            # Verificar si hay datos
            msg = self.session.findById("wnd[0]/sbar").Text
            if "No existen datos" in msg or "No data" in msg:
                print("      [WARN] LX02 sin datos.")
                return pd.DataFrame()

            # --- LÓGICA DE COPIA (IDÉNTICA A BOT PALLET) ---
            print("   -> Copiando datos al portapapeles...")
            try:
                # Intentar menú lista (List -> Export -> Local File / Clipboard)
                # Bot Pallet usa: wnd[0]/mbar/menu[0]/menu[1]/menu[2]
                self.session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                time.sleep(0.5)
                try: self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").select()
                except: self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").select()
                self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            except:
                # Fallback Grid
                try:
                    self.session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").pressToolbarContextButton("&MB_COPY")
                except:
                    print("No se pudo copiar por Menú ni por Grid. Intentando Ctrl+Y (Select All)...")
                    # Último recurso: Select All + Copy (si fuera lista simple)
                    # self.session.findById("wnd[0]/usr").SetFocus() # Foco en área usuario
                    # self.session.findById("wnd[0]").sendVKey(75) # Select All? No estándar.
                    pass

            time.sleep(1)
            
            # Salir de LX02 para limpiar
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            self.session.findById("wnd[0]").sendVKey(0)

            # Procesar Texto (Pipes a Tabs)
            raw_text = pyperclip.paste()
            if not raw_text:
                print("      [ERROR] Portapapeles vacío.")
                return pd.DataFrame()
                
            clean_text = raw_text.replace('|', '\t')
            # Eliminar líneas de separación (---)
            clean_text = "\n".join([line for line in clean_text.split('\n') if "---" not in line and "___" not in line])
            
            pyperclip.copy(clean_text)

            # Pegar en Excel
            print("   -> Pegando en Excel (Pestaña Lx02)...")
            # Pegar en Excel
            print("   -> Pegando en Excel (Pestaña Lx02)...")
            try:
                excel = None
                wb = None
                
                # 1. Intentar conectar a Excel existente
                try:
                    excel = win32com.client.GetActiveObject("Excel.Application")
                    print("      Conectado a Excel existente.")
                except:
                    print("      Iniciando nueva instancia de Excel.")
                    excel = win32com.client.Dispatch("Excel.Application")
                
                excel.Visible = True
                excel.DisplayAlerts = False
                
                # 2. Verificar si el libro ya está abierto
                filename = os.path.basename(excel_path).lower()
                for w in excel.Workbooks:
                    if w.Name.lower() == filename:
                        wb = w
                        print("      Libro ya abierto. Usando instancia.")
                        break
                
                # 3. Si no está abierto, abrirlo
                if not wb:
                    try:
                        wb = excel.Workbooks.Open(excel_path)
                    except Exception as e:
                        print(f"      [ERROR] No se pudo abrir el archivo (¿Está en modo edición?): {e}")
                        return pd.DataFrame()
                
                # 4. Seleccionar/Crear hoja Lx02
                ws_lx02 = None
                for s in wb.Sheets:
                    if s.Name.lower() == "lx02":
                        ws_lx02 = s
                        break
                
                if not ws_lx02:
                    ws_lx02 = wb.Sheets.Add()
                    ws_lx02.Name = "Lx02"
                
                # 5. Pegar
                ws_lx02.Activate()
                ws_lx02.Cells.ClearContents()
                ws_lx02.Range("A1").Select()
                ws_lx02.Paste() # Usar Paste simple en lugar de PasteSpecial

                
                # Intentar alinear encabezados (Bot Pallet Logic Simplified)
                # Buscar fila con "Material"
                header_row = 0
                used_range = ws_lx02.UsedRange
                # Leer primeras 20 filas para buscar header
                # Optimización: Leer valores a lista python para buscar rápido
                
                # Guardar cambios
                wb.Save()
                
                # --- LECTURA DIRECTA DESDE COM (Evita bloqueo de archivo) ---
                print("   -> Leyendo datos desde Excel en memoria...")
                
                # Leer UsedRange a una lista de listas (tupla de tuplas)
                used_range = ws_lx02.UsedRange
                data = used_range.Value # Esto devuelve una tupla de tuplas con los valores
                
                if not data:
                    print("      [WARN] Hoja Lx02 vacía.")
                    return pd.DataFrame()

                # Convertir a DataFrame
                # data[0] es la fila 1 (pero COM es 1-based, python 0-based, Value devuelve tupla python)
                # Buscar fila header
                header_idx = -1
                for i, row in enumerate(data):
                    # Convertir row a string para buscar
                    row_str = " ".join([str(x) for x in row if x is not None])
                    if "Material" in row_str and "Lote" in row_str:
                        header_idx = i
                        break
                
                if header_idx == -1:
                    print("      [WARN] No se encontró encabezado (Material/Lote) en Lx02.")
                    return pd.DataFrame()
                
                # Extraer headers y datos - MANTENER TODOS LOS ÍNDICES
                header_row = data[header_idx]
                headers = []
                for x in header_row:
                    if x is not None:
                        headers.append(str(x).strip())
                    else:
                        headers.append("")  # Mantener columnas vacías para alineación
                
                # Debug: Mostrar headers encontrados
                print(f"      Headers LX02: {[h for h in headers if h][:10]}")  # Primeros 10 no vacíos
                
                # Datos desde header_idx + 1
                rows = []
                for i in range(header_idx + 1, len(data)):
                    row_raw = data[i]
                    # Verificar que no sea fila vacía o de separación
                    if not row_raw or all(x is None or str(x).strip() in ['', '---', '___'] for x in row_raw):
                        continue
                    
                    # Crear diccionario alineado
                    row_dict = {}
                    for j in range(min(len(headers), len(row_raw))):
                        if headers[j]:  # Solo agregar si el header no está vacío
                            row_dict[headers[j]] = row_raw[j]
                    
                    if row_dict:  # Solo agregar si tiene datos
                        rows.append(row_dict)
                
                df_stock = pd.DataFrame(rows)
                
                # Eliminar columnas duplicadas ANTES de renombrar
                df_stock = df_stock.loc[:, ~df_stock.columns.duplicated()]
                
                # Limpiar columnas
                df_stock.columns = df_stock.columns.str.strip()
                
                # Mapeo de columnas (Nombres REALES de LX02 según la imagen)
                # Material -> Material
                # Lote -> Lote  
                # St. disp. -> Cantidad (Stock disponible)
                # Ubicación -> Ubicacion
                # Tp. -> Tipo_Almacen (Tipo de almacén)
                
                # Normalizar nombres
                renames = {}
                for col in df_stock.columns:
                    c = col.lower()
                    if col == "Material":
                        renames[col] = "Material"
                    elif "lote" in c:
                        renames[col] = "Lote"
                    elif "st." in c and "disp" in c:  # St. disp.
                        renames[col] = "Cantidad"
                    elif "ubic" in c:
                        renames[col] = "Ubicacion"
                    elif col == "Tp.":  # Tipo de almacén
                        renames[col] = "Tipo_Almacen"
                
                df_stock = df_stock.rename(columns=renames)
                
                # Verificar que tenemos las columnas necesarias
                required_cols = ["Material", "Lote", "Cantidad", "Ubicacion", "Tipo_Almacen"]
                missing = [col for col in required_cols if col not in df_stock.columns]
                if missing:
                    print(f"      [WARN] Faltan columnas en LX02: {missing}")
                    print(f"      Columnas disponibles: {df_stock.columns.tolist()}")
                
                # Limpiar datos
                if "Cantidad" in df_stock.columns:
                    df_stock["Cantidad"] = df_stock["Cantidad"].apply(self.clean_float)
                
                # Limpiar Material (quitar espacios y None)
                if "Material" in df_stock.columns:
                    df_stock["Material"] = df_stock["Material"].apply(lambda x: self.clean_value(x) if x is not None else "")
                    # Filtrar filas sin material
                    df_stock = df_stock[df_stock["Material"] != ""]
                
                print(f"      Stock recuperado: {len(df_stock)} registros.")
                return df_stock

            except Exception as e:
                print(f"      [ERROR] Falló procesamiento Excel: {e}")
                return pd.DataFrame()

        except Exception as e:
            print(f"Error crítico en LX02: {e}")
            return pd.DataFrame()

    def procesar_requerimientos(self, excel_path, df_stock):
        print("--- Procesando Requerimientos ---")
        
        # Leer Excel usando método robusto (COM o copia temporal)
        df_req = None
        filename = os.path.basename(excel_path)
        
        try:
            # Intentar leer desde Excel abierto (COM)
            excel_app = win32com.client.GetActiveObject("Excel.Application")
            wb_found = None
            for wb in excel_app.Workbooks:
                if wb.Name == filename:
                    wb_found = wb
                    break
            
            if wb_found:
                print("   -> Leyendo requerimientos desde Excel abierto...")
                
                # Buscar la hoja de requerimientos (NO la Lx02)
                ws = None
                # Intentar primero "Hoja1" o la primera hoja que no sea "Lx02"
                for sheet in wb_found.Worksheets:
                    if sheet.Name.lower() != "lx02":
                        ws = sheet
                        print(f"      Usando hoja: {sheet.Name}")
                        break
                
                if not ws:
                    # Fallback: usar primera hoja
                    ws = wb_found.Worksheets(1)
                
                used_range = ws.UsedRange.Value
                
                if used_range:
                    # Convertir a DataFrame
                    if isinstance(used_range[0], tuple):
                        # Limpiar headers agresivamente
                        headers = []
                        for x in used_range[0]:
                            if x is not None:
                                h = str(x).strip().replace('\n', '').replace('\r', '')
                                headers.append(h)
                            else:
                                headers.append("")
                        data = used_range[1:]
                    else:
                        # Una sola fila
                        headers = [str(used_range).strip()]
                        data = []
                    
                    # Debug: Mostrar headers leídos
                    print(f"      Headers detectados: {headers}")
                    
                    rows = []
                    for row in data:
                        if isinstance(row, tuple):
                            row_dict = {}
                            for i in range(min(len(headers), len(row))):
                                if headers[i]:  # Solo agregar si el header no está vacío
                                    row_dict[headers[i]] = row[i]
                            # Solo agregar si tiene datos válidos (no todo None/vacío)
                            if row_dict and any(v is not None and str(v).strip() not in ['', 'nan', 'None'] for v in row_dict.values()):
                                rows.append(row_dict)
                        else:
                            if headers[0] and row is not None:
                                rows.append({headers[0]: row})
                    
                    df_req = pd.DataFrame(rows)
        except:
            pass
        
        # Fallback: Leer desde copia temporal
        if df_req is None:
            print("   -> Leyendo requerimientos desde archivo (copia temporal)...")
            temp_file = excel_path + ".temp_lt01.xlsx"
            try:
                shutil.copyfile(excel_path, temp_file)
                df_req = pd.read_excel(temp_file, sheet_name=0, dtype=str)
            except Exception as e:
                print(f"Error leyendo Excel input: {e}")
                return []
            finally:
                if os.path.exists(temp_file):
                    try: os.remove(temp_file)
                    except: pass

        if df_req is None or df_req.empty:
            print("No se pudieron leer requerimientos.")
            return []

        # Normalizar columnas
        df_req.columns = df_req.columns.str.strip()
        
        # Columnas esperadas según imagen usuario: Material, Cantidad, Unidad, Alm_Dest, Ubicación
        # Mapeo flexible
        col_map = {
            "Material": "Material",
            "Cantidad": "Cantidad",
            "Unidad": "Unidad",
            "Alm_Dest": "Alm_Dest",
            "Ubicación": "Ubicacion" # Sin tilde a veces
        }
        
        # Verificar existencia
        for k, v in col_map.items():
            found = False
            for c in df_req.columns:
                if c.lower() == v.lower() or c.lower() == k.lower():
                    col_map[k] = c
                    found = True
                    break
            if not found and k in ["Material", "Cantidad"]: # Material y Cantidad son críticos
                print(f"Falta columna requerida: {k}")
                return []

        movimientos = []
        
        for idx, row in df_req.iterrows():
            material = self.clean_value(row[col_map["Material"]])
            cant_req = self.clean_float(row[col_map["Cantidad"]])
            
            # Leer Unidad
            unidad = "UN"  # Default
            if "Unidad" in col_map:
                val = self.clean_value(row[col_map["Unidad"]])
                if val and val.lower() != "nan":
                    unidad = val
            
            # Leer Destino
            tipo_dest = "999" # Default
            if "Alm_Dest" in col_map:
                val = self.clean_value(row[col_map["Alm_Dest"]])
                if val and val.lower() != "nan": tipo_dest = val
            
            ubic_dest = ""
            if "Ubicación" in col_map:
                val = self.clean_value(row[col_map["Ubicación"]])
                if val and val.lower() != "nan": 
                    ubic_dest = self.format_ubicacion(val)
            
            # Lógica de inferencia si viene vacío (Legacy support)
            if not ubic_dest:
                if tipo_dest == "999": ubic_dest = "SCHROTT"
                elif tipo_dest == "920": ubic_dest = "TRANSFER"

            print(f"Req: Mat {material} | Cant {cant_req} | Dest {tipo_dest}-{ubic_dest}")
            
            # Debug: Mostrar primeros materiales del stock
            if idx == 0:  # Solo en la primera iteración
                materiales_sample = df_stock['Material'].head(10).values.tolist()
                print(f"   [DEBUG] Primeros materiales en stock: {materiales_sample}")
            
            # Filtrar stock
            stock_mat = df_stock[df_stock["Material"] == material].copy()
            
            if stock_mat.empty:
                print(f"   [ALERTA] No hay stock para material {material}")
                print(f"   [DEBUG] Buscando '{material}' (tipo: {type(material)})")
                continue
            
            print(f"   -> Encontrados {len(stock_mat)} lotes disponibles")
            
            # Ordenar por cantidad descendente (Greedy)
            stock_mat = stock_mat.sort_values(by="Cantidad", ascending=False)
            
            cant_pendiente = cant_req
            
            for _, stock_row in stock_mat.iterrows():
                if cant_pendiente <= 0:
                    break
                
                cant_disponible = stock_row["Cantidad"]
                lote = self.clean_value(stock_row["Lote"])  # Limpiar .0 del lote
                ubic_orig = self.format_ubicacion(self.clean_value(stock_row["Ubicacion"]))
                tipo_orig = self.clean_value(stock_row["Tipo_Almacen"])
                
                tomar = min(cant_pendiente, cant_disponible)
                
                print(f"   -> Seleccionando lote {lote} | Disponible: {cant_disponible} | Tomar: {tomar}")
                
                movimientos.append({
                    "Material": material,
                    "Cantidad": tomar,
                    "Unidad": unidad,
                    "Lote": lote,
                    "Tipo_Origen": tipo_orig,
                    "Ubic_Origen": ubic_orig,
                    "Tipo_Destino": tipo_dest,
                    "Ubic_Destino": ubic_dest
                })
                
                cant_pendiente -= tomar
                
            if cant_pendiente > 0:
                print(f"   [ALERTA] Stock insuficiente para {material}. Faltaron {cant_pendiente}")

        return movimientos

    def ejecutar_lt01(self, movimientos):
        print(f"--- Ejecutando {len(movimientos)} Movimientos en LT01 ---")
        
        for i, mov in enumerate(movimientos):
            print(f"[{i+1}/{len(movimientos)}] LT01: {mov['Material']} | Lote {mov['Lote']} | {mov['Cantidad']} -> {mov['Tipo_Destino']}")
            
            try:
                self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nLT01"
                self.session.findById("wnd[0]").sendVKey(0)
                time.sleep(1.0)
                
                # Pantalla 1
                self.session.findById("wnd[0]/usr/ctxtLTAK-LGNUM").text = "NCD"
                self.session.findById("wnd[0]/usr/ctxtLTAK-BWLVS").text = "999" 
                self.session.findById("wnd[0]/usr/ctxtLTAP-MATNR").text = mov["Material"]
                self.session.findById("wnd[0]/usr/txtRL03T-ANFME").text = str(mov["Cantidad"]).replace('.', ',')
                self.session.findById("wnd[0]/usr/ctxtLTAP-ALTME").text = mov["Unidad"]
                self.session.findById("wnd[0]/usr/ctxtLTAP-WERKS").text = "SGSJ"
                self.session.findById("wnd[0]/usr/ctxtLTAP-LGORT").text = "NCD1"
                self.session.findById("wnd[0]/usr/ctxtLTAP-CHARG").text = mov["Lote"]
                
                self.session.findById("wnd[0]").sendVKey(0) 
                
                if self.session.ActiveWindow.Name == "wnd[1]":
                     pass # Ignorar popups informativos

                # Pantalla 2
                try: self.session.findById("wnd[0]/usr/chkRL03T-SQUIT").Selected = True
                except: pass
                
                # Origen
                self.session.findById("wnd[0]/usr/ctxtLTAP-VLTYP").text = mov["Tipo_Origen"]
                self.session.findById("wnd[0]/usr/txtLTAP-VLPLA").text = mov["Ubic_Origen"]
                
                # Destino
                self.session.findById("wnd[0]/usr/ctxtLTAP-NLTYP").text = mov["Tipo_Destino"]
                if mov["Ubic_Destino"]:
                    self.session.findById("wnd[0]/usr/txtLTAP-NLPLA").text = mov["Ubic_Destino"]
                
                self.session.findById("wnd[0]").sendVKey(0) 
                self.session.findById("wnd[0]").sendVKey(0) # Guardar
                
                msg = self.session.findById("wnd[0]/sbar").Text
                print(f"   SAP: {msg}")
                
            except Exception as e:
                print(f"   [ERROR] Falló movimiento: {e}")

    def run(self, excel_path):
        # 1. Descargar Stock (y pegar en Excel)
        df_stock = self.descargar_stock_lx02(excel_path)
        if df_stock.empty:
            print("No se pudo obtener stock de LX02. Abortando.")
            return
        
        # 2. Procesar lógica de negocio
        movimientos = self.procesar_requerimientos(excel_path, df_stock)
        
        if not movimientos:
            print("No hay movimientos para realizar.")
            return
            
        # 3. Ejecutar en SAP
        self.ejecutar_lt01(movimientos)
        print("--- Proceso Finalizado ---")

if __name__ == "__main__":
    pass
