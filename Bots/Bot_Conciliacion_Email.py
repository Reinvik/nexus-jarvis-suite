import win32com.client
import sys
import time
import re
import os
import shutil
import hashlib
import json
import ctypes  # [NEW] Para MessageBox
from datetime import datetime, timedelta
import pythoncom

class SapBotConciliacionEmail:
    # Tiempo de expiración del caché anti-duplicados (5 minutos)
    CACHE_EXPIRY_MINUTES = 5
    
    def __init__(self):
        self.outlook = None
        self.namespace = None
        
        # Ruta al template existente en OneDrive
        onedrive_base = os.path.join(os.path.expanduser("~"), "OneDrive - CIAL Alimentos")
        self.template_path = os.path.join(
            onedrive_base,
            "Archivos de Operación  Outbound CD - 16.-Inventario Critico",
            "carga_migo.xlsx"
        )
        
        print(f"> Template configurado: {self.template_path}")
        
        # Sistema anti-duplicados: archivo de caché
        self.cache_file = os.path.join(os.path.dirname(__file__), "transfer_cache.json")
        self.processed_transfers = self._load_cache()
        
        # Inicializar COM para este hilo
        try:
            pythoncom.CoInitialize()
        except:
            pass
    
    def _load_cache(self):
        """Carga el caché de traspasos procesados desde archivo"""
        try:
            if os.path.exists(self.cache_file):
                with open(self.cache_file, 'r') as f:
                    cache = json.load(f)
                # Limpiar entradas expiradas
                now = datetime.now()
                valid_cache = {}
                for key, timestamp_str in cache.items():
                    timestamp = datetime.fromisoformat(timestamp_str)
                    if now - timestamp < timedelta(minutes=self.CACHE_EXPIRY_MINUTES):
                        valid_cache[key] = timestamp_str
                return valid_cache
        except Exception as e:
            print(f"[WARN] Error cargando caché: {e}")
        return {}
    
    def _save_cache(self):
        """Guarda el caché de traspasos procesados"""
        try:
            with open(self.cache_file, 'w') as f:
                json.dump(self.processed_transfers, f)
        except Exception as e:
            print(f"[WARN] Error guardando caché: {e}")
    
    def _generate_transfer_hash(self, data):
        """Genera un hash único para un conjunto de datos de traspaso"""
        # Ordenar data para consistencia
        sorted_data = sorted(data, key=lambda x: (x['Material'], x['Lote']))
        # Crear string con todos los datos relevantes
        data_str = "|".join([
            f"{item['Material']}:{item['Cantidad']}:{item['Lote']}"
            for item in sorted_data
        ])
        return hashlib.md5(data_str.encode()).hexdigest()
    
    def _is_duplicate_transfer(self, data):
        """Verifica si este traspaso ya fue procesado recientemente"""
        transfer_hash = self._generate_transfer_hash(data)
        
        if transfer_hash in self.processed_transfers:
            cached_time = datetime.fromisoformat(self.processed_transfers[transfer_hash])
            elapsed = datetime.now() - cached_time
            if elapsed < timedelta(minutes=self.CACHE_EXPIRY_MINUTES):
                print(f"[WARN] DUPLICADO DETECTADO: Este traspaso fue procesado hace {elapsed.seconds // 60}m {elapsed.seconds % 60}s")
                print(f"   Hash: {transfer_hash[:8]}...")
                return True
        
        return False
    
    def _mark_transfer_processed(self, data):
        """Marca un traspaso como procesado"""
        transfer_hash = self._generate_transfer_hash(data)
        self.processed_transfers[transfer_hash] = datetime.now().isoformat()
        self._save_cache()
        print(f"[INFO] Traspaso registrado en caché (válido por {self.CACHE_EXPIRY_MINUTES} min)")
    
    def close_excel_if_open(self, filename):
        """Cierra un archivo específico si está abierto en Excel"""
        try:
            excel = win32com.client.GetObject(Class="Excel.Application")
            for wb in excel.Workbooks:
                if filename.lower() in wb.FullName.lower():
                    print(f"[WARN] Archivo abierto en Excel. Cerrando: {wb.Name}")
                    wb.Close(SaveChanges=True)
                    print("[OK] Archivo cerrado automáticamente")
                    return True
        except:
            pass  # Excel no está corriendo o el archivo no está abierto
        return False
    
    def connect_outlook(self):
        """Conecta a Outlook usando win32com"""
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            print("[OK] Conectado a Outlook exitosamente")
            return True
        except Exception as e:
            print(f"[ERROR] Error conectando a Outlook: {e}")
            print("Asegúrate de tener Outlook abierto y configurado.")
            return False
    
    def get_perdida_vacio_folder(self):
        """Obtiene la carpeta 'Perdida vacío' de Outlook"""
        try:
            inbox = self.namespace.GetDefaultFolder(6)  # 6 = Inbox
            
            # Buscar carpeta "Perdida vacío" (con tilde)
            for folder in inbox.Folders:
                # Normalizar para comparación (quitar tildes y espacios extra)
                folder_name_normalized = folder.Name.lower().replace('í', 'i').replace('ó', 'o').strip()
                if "perdida" in folder_name_normalized and "vacio" in folder_name_normalized:
                    print(f"[DIR] Carpeta encontrada: {folder.Name}")
                    return folder
            
            print("[WARN] No se encontró carpeta 'Perdida vacío'")
            print("Carpetas disponibles:")
            for folder in inbox.Folders:
                print(f"  - {folder.Name}")
            return None
            
        except Exception as e:
            print(f"[ERROR] Error accediendo a carpetas: {e}")
            return None
    
    def get_or_create_processed_folder(self, parent_folder):
        """Obtiene o crea la subcarpeta 'Procesados'"""
        try:
            for folder in parent_folder.Folders:
                if folder.Name == "Procesados":
                    return folder
            print("[DIR] Creando carpeta 'Procesados'...")
            return parent_folder.Folders.Add("Procesados")
        except Exception as e:
            print(f"[ERROR] Error gestionando carpeta Procesados: {e}")
            return None

    def get_or_create_error_folder(self, parent_folder):
        """Obtiene o crea la subcarpeta 'Errores' para emails fallidos"""
        try:
            for folder in parent_folder.Folders:
                if folder.Name == "Errores":
                    return folder
            print("[DIR] Creando carpeta 'Errores'...")
            return parent_folder.Folders.Add("Errores")
        except Exception as e:
            print(f"[ERROR] Error gestionando carpeta Errores: {e}")
            return None

    def read_pending_emails(self, folder):
        """Lee TODOS los emails en la carpeta (cola de procesamiento)"""
        try:
            # Ordenar por fecha: (True = Descendente/Más recientes primero)
            # Priorizamos los actuales como pidió el usuario
            
            # FILTRO: SOLO NO LEÍDOS
            # Esto evita procesar todo el historial si el bot falla
            messages_collection = folder.Items.Restrict("[UnRead] = True")
            messages_collection.Sort("[ReceivedTime]", True) 
            
            total_count = messages_collection.Count
            print(f"\n[DEBUG] Analizando carpeta '{folder.Name}' (Solo NO LEÍDOS)")
            print(f"   Total de items pendientes: {total_count}")
            
            if total_count == 0:
                print("[INFO] No hay correos nuevos (No Leídos).")
                return []

            # ITERACIÓN SEGURA (SNAPSHOT)
            # Evita loops infinitos si GetFirst/GetNext falla
            valid_messages = []
            
            # Limite de seguridad interno para la lectura
            READ_LIMIT = 50 
            count = 0
            
            # Convertimos a lista para congelar el estado
            # Nota: Esto puede ser lento si hay miles, pero es más seguro
            # Si hay demasiados, limitamos la lectura inicial
            
            print(f"[DEBUG] Leyendo snapshot de mensajes no leídos (Límite lectura: {READ_LIMIT})...")
            
            for message in messages_collection:
                if count >= READ_LIMIT:
                    print(f"   [INFO] Límite de lectura alcanzado ({READ_LIMIT}). Se procesarán estos primero.")
                    break
                    
                try:
                    subject = message.Subject if hasattr(message, 'Subject') else ""
                    sender = message.SenderName if hasattr(message, 'SenderName') else ""
                    
                    # Ignorar rebotes
                    subject_lower = subject.lower()
                    sender_lower = sender.lower()
                    
                    if ("no se puede entregar" in subject_lower or 
                        "undeliverable" in subject_lower or
                        "postmaster" in sender_lower or
                        "reporte bot" in subject_lower):
                        pass # Silencioso para no llenar log
                    else:
                        valid_messages.append(message)
                        count += 1
                except:
                    pass
            
            print(f"[OK] Se procesarán {len(valid_messages)} emails válidos de este lote")
            return valid_messages
            
        except Exception as e:
            print(f"[ERROR] Error leyendo emails: {e}")
            return []
    
    def parse_email_data(self, email_body, html_body=None):
        """
        Extrae datos de la tabla en el email.
        Soporta dos formatos:
        1. Inline: "1053    1    UN    ...    33883586"
        2. Multi-línea: Cada campo en una línea separada
        3. HTML Table: Si se proporciona html_body
        """
        print("[DATA] Parseando datos del email...")
        
        data = []
        if email_body:
            lines = email_body.split('\n')
            # Limpiar líneas vacías y espacios
            lines = [line.strip() for line in lines if line.strip()]
        else:
            lines = []
        
        # Método 1: Buscar patrón inline (original)
        for line in lines:
            # Patrón: número(SKU) + número(cantidad) + UN + ... + número largo(lote)
            match = re.search(r'(\d{3,5})\s+(\d+)\s+(UN)\s+.*?\s+(\d{8,})', line)
            
            if match:
                material = match.group(1).strip()
                cantidad = match.group(2).strip()
                unidad = match.group(3).strip()
                lote = match.group(4).strip()
                
                data.append({
                    'Material': material,
                    'Cantidad': cantidad,
                    'Unidad': unidad,
                    'Lote': lote
                })
                
                print(f"   [OK] {material} | {cantidad} {unidad} | Lote: {lote}")
        
        # Método 2: Si no encontró nada, buscar formato multi-línea
        if not data and lines:
            print("   [INFO] Formato inline no detectado, intentando formato multi-línea...")
            
            # Buscar el header de la tabla (SKU puede estar solo en una línea)
            header_idx = -1
            for i, line in enumerate(lines):
                # Buscar línea que contenga "SKU" (puede estar sola)
                if 'SKU' in line.upper() and len(line.strip()) < 20:  # Header corto
                    # Verificar que las siguientes líneas sean CANTIDAD, UN, LOTE
                    if i + 3 < len(lines):
                        next_lines = [lines[i+1].upper(), lines[i+2].upper(), lines[i+3].upper()]
                        if 'CANTIDAD' in next_lines[0] and ('UN' in next_lines[1] or 'MEDIDA' in next_lines[1]) and 'LOTE' in next_lines[2]:
                            header_idx = i + 3  # Empezar después de LOTE
                            print(f"   [INFO] Header encontrado en líneas {i}-{i+3}")
                            break
            
            if header_idx >= 0:
                # Procesar líneas después del header
                i = header_idx + 1
                
                while i < len(lines):
                    # Detectar separadores de email
                    line_lower = lines[i].lower()
                    if ('_____' in lines[i] or  # Separador visual
                        line_lower.startswith('de:') or 
                        line_lower.startswith('from:') or
                        line_lower.startswith('enviado:') or
                        line_lower.startswith('sent:')):
                        print(f"   [SKIP] Separador de email detectado en línea {i}. Deteniendo extracción.")
                        print(f"      (Solo se procesó el email actual, no la cadena completa)")
                        break
                    
                    # Buscar secuencia: SKU -> Cantidad -> UN -> Lote
                    if i + 3 < len(lines):
                        try:
                            # Intentar extraer 4 líneas consecutivas
                            potential_sku = lines[i].strip()
                            potential_qty = lines[i + 1].strip()
                            potential_un = lines[i + 2].strip()
                            potential_lote = lines[i + 3].strip()
                            
                            # Validar que sean los datos correctos
                            if (potential_sku.isdigit() and len(potential_sku) >= 3 and
                                potential_qty.isdigit() and
                                potential_un.upper() == 'UN' and
                                potential_lote.isdigit() and len(potential_lote) >= 8):
                                
                                data.append({
                                    'Material': potential_sku,
                                    'Cantidad': potential_qty,
                                    'Unidad': potential_un,
                                    'Lote': potential_lote
                                })
                                
                                print(f"   [OK] {potential_sku} | {potential_qty} {potential_un} | Lote: {potential_lote}")
                                
                                # Avanzar 4 líneas
                                i += 4
                                continue
                        except (IndexError, ValueError):
                            pass
                    
                    # Si no coincide, avanzar una línea
                    i += 1
        
        # Método 3: Buscar tabla HTML si existe
        if not data and html_body and '<table' in html_body.lower():
            print("   [INFO] Detectada tabla HTML, intentando extraer...")
            # Patrón para extraer datos de celdas <td>
            cells = re.findall(r'<td[^>]*>(.*?)</td>', html_body, re.IGNORECASE | re.DOTALL)
            
            # Agrupar de 4 en 4 (SKU, Cantidad, UN, Lote)
            for i in range(0, len(cells) - 3, 4):
                try:
                    sku = cells[i].strip()
                    qty = cells[i + 1].strip()
                    un = cells[i + 2].strip()
                    lote = cells[i + 3].strip()
                    
                    if sku.isdigit() and qty.isdigit() and 'UN' in un.upper() and lote.isdigit():
                        data.append({
                            'Material': sku,
                            'Cantidad': qty,
                            'Unidad': 'UN',
                            'Lote': lote
                        })
                        print(f"   [OK] {sku} | {qty} UN | Lote: {lote}")
                except:
                    continue
        
        print(f"[INFO] Total extraido: {len(data)} registros")
        
        if not data:
            print("\n[DEBUG] No se pudieron extraer datos. Mostrando primeras 30 lineas del email:")
            for i, line in enumerate(lines[:30], 1):
                print(f"   {i:2d}. {line[:80]}")
        
        return data

    def create_excel_from_template(self, data, output_path=None):
        """
        Crea el archivo Excel basándose en el template.
        Retorna: (success, path_created)
        """
        if not data:
            return False, None
            
        try:
            # RETRY LOGIC para manejar bloqueos transitorios
            max_retries = 3
            current_try = 0
            
            excel = None
            wb = None
            fixed_output = None
            
            while current_try < max_retries:
                try:
                    current_try += 1
                    
                    # Asegurar limpieza previa si falló antes
                    self.close_excel_if_open('carga_migo')
                    
                    excel = win32com.client.Dispatch("Excel.Application")
                    excel.Visible = False
                    excel.DisplayAlerts = False
                    
                    # Verificar si existe el template
                    if not os.path.exists(self.template_path):
                        print(f"[ERROR] Template no encontrado: {self.template_path}")
                        return False, None
                    
                    # Usar nombre único con timestamp para evitar bloqueos de archivo
                    filename = f"Perdida_Vacio_{int(time.time())}.xlsx"
                    fixed_output = os.path.join(
                        os.path.dirname(__file__),
                        filename
                    )
                    
                    print(f"[INFO] Creando Excel (Intento {current_try}/{max_retries}): {filename}")
                    
                    # Copiar template a destino para no modificar el original
                    shutil.copy(self.template_path, fixed_output)
                    
                    # Abrir el nuevo archivo
                    wb = excel.Workbooks.Open(fixed_output)
                    ws = wb.Sheets(1) # Asumimos Hoja 1
                    
                    # Llenar datos (partiendo de fila 2, asumiendo encabezados)
                    print(f"[DATA] Escribiendo {len(data)} registros en Excel...")
                    
                    for i, item in enumerate(data):
                        row = i + 2 # Fila base 2
                        
                        # Material (Col A / 1)
                        ws.Cells(row, 1).Value = item['Material']
                        
                        # Cantidad (Col C / 3)
                        ws.Cells(row, 3).Value = item['Cantidad']
                        
                        # Unidad (Col D / 4)
                        ws.Cells(row, 4).Value = "UN"
                        
                        # Lote_Orig (Col E / 5)
                        ws.Cells(row, 5).NumberFormat = "@" # Texto
                        ws.Cells(row, 5).Value = str(item['Lote'])
                        
                        # Alm_Orig (Col F / 6) -> NCD1
                        ws.Cells(row, 6).Value = "NCD1"
                        
                        # Centro_Orig (Col G / 7) -> SGSJ
                        ws.Cells(row, 7).Value = "SGSJ" 
                        
                        # Alm_Dest (Col H / 8) -> CDNW
                        ws.Cells(row, 8).Value = "CDNW"
                        
                        # Centro_Dest (Col I / 9) -> SGSJ
                        ws.Cells(row, 9).Value = "SGSJ"
                        
                        # Lote_Dest (Col J / 10) -> REPROCP1
                        ws.Cells(row, 10).Value = "REPROCP1"
                        
                        # Texto Cabecera (Col K / 11)
                        ws.Cells(row, 11).Value = "Perdida de vacio"
                    
                    # PASO 8: Guardar
                    print(f"\n[INFO] Guardando archivo...")
                    wb.Save()
                    
                    print(f"\n{'='*60}")
                    print(f"[OK] EXCEL CREADO EXITOSAMENTE")
                    print(f"[DIR] Ubicación: {fixed_output}")
                    print(f"[DATA] Registros: {len(data)}")
                    print(f"{'='*60}\n")
                    
                    # Cerrar Excel limpiamente
                    wb.Close()
                    excel.Quit()
                    
                    # Éxito!
                    return True, fixed_output
                    
                except PermissionError:
                    print(f"[WARN] Archivo bloqueado (PermissionError). Reintentando en 3s...")
                    try:
                        if wb: wb.Close(False)
                        if excel: excel.Quit()
                    except: pass
                    time.sleep(3)
                    
                except Exception as e:
                    print(f"[WARN] Error en intento {current_try}: {e}")
                    try:
                        if wb: wb.Close(False)
                        if excel: excel.Quit()
                    except: pass
                    time.sleep(2)
            
            # Si llegamos aquí, fallaron todos los intentos
            print(f"[ERROR] No se pudo crear Excel después de {max_retries} intentos.")
            return False, None
            
        except Exception as e:
            print(f"[ERROR] Error fatal creando Excel: {e}")
            import traceback
            traceback.print_exc()
            return False, None
    
    # ==============================================================================
    #                          LÓGICA SAP (MIGO) NATIVA
    # ==============================================================================
    
    def load_plant_mapping(self):
        """
        Carga el mapeo de SKU a Planta desde la hoja 'PLANTA' del template.
        Retorna dict: { 'SKU': 'REPROCP1'/'REPROCP2' }
        """
        print("[DATA] Cargando tabla maestra de PLANTAS...")
        mapping = {}
        try:
            import pandas as pd
            # Leer hoja PLANTA
            df = pd.read_excel(self.template_path, sheet_name='PLANTA', dtype=str)
            
            # Limpiar nombres de columnas
            df.columns = df.columns.str.strip()
            
            if 'SKU' in df.columns and 'Planta' in df.columns:
                for _, row in df.iterrows():
                    sku = str(row['SKU']).strip()
                    planta = str(row['Planta']).strip().upper()
                    
                    # Lógica de mapeo P1/P2 -> Lote
                    if planta == 'P1':
                        mapping[sku] = 'REPROCP1'
                    elif planta == 'P2':
                        mapping[sku] = 'REPROCP2'
                    else:
                        mapping[sku] = 'REPROCP1' # Default P1
                        
            print(f"[DATA] Mapeo de plantas cargado: {len(mapping)} SKUs")
        except Exception as e:
            print(f"[WARN] No se pudo cargar hoja PLANTA: {e}")
            print("       Se usará REPROCP1 por defecto para todo.")
            
        return mapping

    def connect_to_sap(self, max_retries=3):
        """Conexión robusta a SAP con reintentos para la sesión de MIGO"""
        try:
            pythoncom.CoInitialize()
        except: pass
        
        self.session = None
        for attempt in range(max_retries):
            try:
                sap_gui = win32com.client.GetObject("SAPGUI")
                application = sap_gui.GetScriptingEngine
                connection = application.Children(0)
                self.session = connection.Children(0)
                
                # Verificar sesión activa
                _ = self.session.findById("wnd[0]")
                print("[SAP] Conectado exitosamente")
                return True
            except Exception as e:
                print(f"   [SAP] Intento {attempt + 1} fallido: {e}")
                time.sleep(2)
        
        print("[SAP] Error crítico: No se puede conectar a SAP GUI.")
        return False

    def find_migo_table(self):
        """Busca la tabla de items en MIGO dinámicamente"""
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
        return None

    def map_columns(self, table):
        """Identifica índices de columnas en la tabla SAP"""
        print("[SAP] Mapeando columnas...")
        search_keys = {
            "MAT": ["MATNR", "MAKTX", "MATERIAL"], "QTY": ["ERFMG", "CANTIDAD", "QTY"], "UNIT": ["ERFME", "UNIDAD", "UNIT"],
            "PLANT_O": ["WERKS", "NAME1", "CENTRO"], "LOC_O": ["LGORT", "LGOBE", "ALMACEN"], "BATCH_O": ["CHARG", "LOTE", "BATCH"],
            "PLANT_D": ["UMWRK", "UMNAME1"], "LOC_D": ["UMLGO", "UMLGOBE"], "BATCH_D": ["UMCHA"]
        }
        cols = {k: -1 for k in search_keys}
        try:
            print(f"   [DBG] Grilla con {table.Columns.Count} columnas:")
            for i in range(table.Columns.Count):
                name = table.Columns.Item(i).Name
                title = table.Columns.Item(i).Title
                # print(f"      Col {i}: Name='{name}' Title='{title}'")
                
                for key, possibilities in search_keys.items():
                    if cols[key] == -1:
                        for p in possibilities:
                            if p in name.upper() or p in title.upper():
                                cols[key] = i
                                print(f"      => Asignada {key} a columna {i} ({title})")
                                break
        except Exception as e:
            print(f"   [WARN] Error mapeando columnas: {e}")
            pass
        return cols

    def set_val_robust(self, table, col_idx, row_vis, val):
        """Escribe un valor en una celda SAP de forma segura"""
        if val is None or str(val).strip() == "": return
        if col_idx == -1: return

        try:
            # Scroll horizontal
            if col_idx >= 15:
                try: 
                    if table.FirstVisibleColumn < (col_idx - 2):
                        table.FirstVisibleColumn = col_idx - 2
                except: pass
            else:
                try: table.FirstVisibleColumn = 0
                except: pass
            
            cell = table.GetCell(row_vis, col_idx)
            if cell.Changeable:
                cell.Text = str(val)
            else:
                # print(f"   [DBG] Celda {row_vis},{col_idx} no editable. Val: {val}")
                pass
        except Exception as e:
            pass

    def start_transaction_migo(self):
        """Limpia y reinicia MIGO con /nMIGO"""
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nMIGO"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(2.0)
            try: self.session.findById("wnd[0]").maximize()
            except: pass
            try: self.session.findById("wnd[0]/shellcont/shell/shellcont[2]/shell/shellcont[0]/shell").pressButton("CLOSE")
            except: pass
        except Exception as e:
            print(f"[SAP] Warn al iniciar TRX: {e}")

    def ask_final_success(self, processed_count):
        """
        Pregunta al usuario SI LA CARGA FUE EXITOSA para proceder a mover el email.
        """
        MB_YESNO = 4
        MB_ICONQUESTION = 0x20
        MB_TOPMOST = 0x40000
        
        msg = (
            f"El bot ha terminado de rellenar {processed_count} líneas en MIGO.\n\n"
            "1. Revise los datos en SAP.\n"
            "2. Presione 'Contabilizar' (Guardar) manualmente en SAP.\n\n"
            "¿La carga fue exitosa? (Responder 'SÍ' moverá el email a Procesados)"
        )
        
        ret = ctypes.windll.user32.MessageBoxW(0, msg, "Confirmar Éxito - JARVIS", MB_YESNO | MB_ICONQUESTION | MB_TOPMOST)
        return ret == 6 # 6 = YES

    def execute_migo_native(self, data):
        """
        Ejecución nativa de MIGO (Modo Asistido).
        Rellena datos pero NO guarda automáticamente.
        """
        print("[BOT] Ejecutando MIGO (Nativo / Asistido)...")
        if not self.connect_to_sap():
            return False

        try:
            # 1. Cargar Mapeo de Plantas
            plant_map = self.load_plant_mapping()
            
            self.start_transaction_migo()
            
            table = self.find_migo_table()
            if not table:
                print("[ERROR] No se encontró la tabla de MIGO")
                return False
            
            # Mapear columnas dinámicamente
            cols_map = self.map_columns(table)
            
            # Escribir cabecera
            try:
                self.session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0006/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0112/txtGOHEAD-BKTXT").Text = "Perdida de vacio"
            except: pass

            # PROCESAMIENTO POR BLOQUES
            import math
            BLOCK_SIZE = 15
            total_items = len(data)
            num_blocks = math.ceil(total_items / BLOCK_SIZE)
            current_sap_scroll = 0
            
            print(f"[SAP] Cargando {total_items} items en {num_blocks} bloques...")
            
            for block_idx in range(num_blocks):
                start = block_idx * BLOCK_SIZE
                end = min((block_idx + 1) * BLOCK_SIZE, total_items)
                block_data = data[start:end]
                
                print(f"   Bloque {block_idx+1}: Items {start}-{end}")
                
                # Scroll vertical inicial
                table = self.find_migo_table()
                try: table.VerticalScrollbar.Position = current_sap_scroll
                except: pass
                time.sleep(0.5)
                
                # --- ESCRITURA ---
                
                # 1. Material
                if cols_map["MAT"] != -1:
                    for i, item in enumerate(block_data):
                        self.set_val_robust(table, cols_map["MAT"], i, item['Material'])
                
                # 2. Cantidad 
                if cols_map["QTY"] != -1:
                    for i, item in enumerate(block_data):
                        self.set_val_robust(table, cols_map["QTY"], i, item['Cantidad'])

                # 3. Unidad (UN)
                if cols_map["UNIT"] != -1:
                    for i, item in enumerate(block_data):
                        self.set_val_robust(table, cols_map["UNIT"], i, "UN")
                
                # 4. Datos Fijos Origen (NCD1 / SGSJ / 920)

                
                if cols_map["PLANT_O"] != -1:
                    for i, item in enumerate(block_data):
                        self.set_val_robust(table, cols_map["PLANT_O"], i, "SGSJ")
                
                if cols_map["LOC_O"] != -1:
                    for i, item in enumerate(block_data):
                        self.set_val_robust(table, cols_map["LOC_O"], i, "NCD1")

                # VALIDAD FASE 1 (Validar Materiales)
                self.session.findById("wnd[0]").sendVKey(0) # Enter
                time.sleep(1.0)
                
                # 5. Datos Faltantes y Destino
                table = self.find_migo_table() # Refrescar referencia
                
                # [MOVED] Lote Origen (Se escribe AHORA, ya que el material fue validado)
                if cols_map["BATCH_O"] != -1:
                    for i, item in enumerate(block_data):
                        self.set_val_robust(table, cols_map["BATCH_O"], i, item['Lote'])

                # Datos Destino
                if cols_map["PLANT_D"] != -1:
                    for i, item in enumerate(block_data):
                        self.set_val_robust(table, cols_map["PLANT_D"], i, "SGSJ")

                if cols_map["LOC_D"] != -1:
                    for i, item in enumerate(block_data):
                        self.set_val_robust(table, cols_map["LOC_D"], i, "CDNW")
                
                # VALIDAD FASE 2
                self.session.findById("wnd[0]").sendVKey(0)
                time.sleep(1.0)
                
                # 6. Lote Destino (REPROCP1 / REPROCP2) - DINÁMICO
                table = self.find_migo_table()
                if cols_map["BATCH_D"] != -1:
                    for i, item in enumerate(block_data):
                        sku = str(item['Material']).strip()
                        # Buscar en el mapa, default REPROCP1 si no existe
                        lote_dest = plant_map.get(sku, "REPROCP1")
                        self.set_val_robust(table, cols_map["BATCH_D"], i, lote_dest)

                self.session.findById("wnd[0]").sendVKey(0)
                time.sleep(0.5)
                
                current_sap_scroll += len(block_data)

            # --- MODO ASISTIDO: NO GUARDAR ---
            print("[SAP] Carga finalizada.")
            print("[USER] Esperando validación manual del usuario...")
            
            # Preguntar al usuario si todo salió bien
            if self.ask_final_success(total_items):
                print("[USER] Usuario confirmó éxito.")
                return True
            else:
                print("[USER] Usuario reportó fallo.")
                return False

        except Exception as e:
            print(f"[ERROR] Error en MIGO Nativo: {e}")
            return False

    def execute_migo_transfer(self, excel_path_unused):
        """Wrapper para mantener compatibilidad"""
        print("[WARN] Método execute_migo_transfer llamado. Redirigiendo...")
        return False


    
    def send_error_email(self, error_msg, original_email=None, context=""):
        """Envía notificación de error al administrador"""
        print(f"[ERROR] Enviando reporte de error: {error_msg}")
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            mail.To = "ariel.mella@cial.cl"
            mail.Subject = f"[ALERTA] JARVIS: {context}"
            
            body = f"""Hola Ariel,

JARVIS ha detectado un problema que requiere tu atención.

CONTEXTO: {context}
ERROR: {error_msg}

"""
            if original_email:
                try:
                    body += f"""
DATOS DEL EMAIL ORIGINAL:
- De: {original_email.SenderName}
- Asunto: {original_email.Subject}
- Fecha: {original_email.ReceivedTime}
"""
                except:
                    pass
            
            body += "\nPor favor revisa el servidor."
            mail.Body = body
            mail.Send()
            print("[OK] Reporte de error enviado")
        except Exception as e:
            print(f"[ERROR] Falló el envío del reporte de error: {e}")
            
    def send_confirmation_email(self, original_email, data_summary):
        """Responde al email original (Reply All) para mantener la cadena"""
        print("[INFO] Generando respuesta (Reply All)...")
        
        try:
            # Crear respuesta manteniendo cadena
            reply = original_email.ReplyAll()
            
            # ASEGURAR DESTINATARIOS OBLIGATORIOS
            # ReplyAll pone a los originales en To/CC. Nosotros agregamos a los jefes si faltan.
            priority_recipients = ["Maicol.Pradenas@cial.cl", "christopher.aleman@cial.cl"]
            
            current_to = reply.To or ""
            current_cc = reply.CC or ""
            
            # Agregar prioritarios al TO si no están
            new_to_list = [t.strip() for t in current_to.split(';') if t.strip()]
            for recipient in priority_recipients:
                # Chequeo simple por string (no perfecto pero funcional)
                if recipient.lower() not in current_to.lower() and recipient.lower() not in current_cc.lower():
                    new_to_list.append(recipient)
            
            reply.To = "; ".join(new_to_list)
            
            # Agregarme a mí en CC si no estoy
            my_email = "ariel.mella@cial.cl"
            if my_email.lower() not in current_cc.lower() and my_email.lower() not in current_to.lower():
                reply.CC = f"{current_cc}; {my_email}".strip("; ")
            
            # CUERPO DEL MENSAJE
            # Construimos el mensaje y lo ponemos AL PRINCIPIO del HTMLBody existente
            mensaje_nuevo = f"""
            <p>Estimados,</p>
            <p>Se ha completado el traspaso al CDNW por pérdida de vacío.</p>
            <p><strong>Registros procesados:</strong> {len(data_summary)}</p>
            <pre style="font-family: Consolas, monospace;">{self._format_summary(data_summary)}</pre>
            <p>Atte.<br>
            JARVIS - Asistente de Automatización de<br>
            <br>
            Ariel Mella - Analista de Inventario</p>
            <br>
            """
            
            # Insertar antes del cuerpo original
            reply.HTMLBody = mensaje_nuevo + reply.HTMLBody
            
            reply.Save() # Guardar en borradores por seguridad
            reply.Display()  # Mostrar para revisión manual
            print("[OK] Respuesta generada y lista para enviar (Reply All)")
            return True
            
        except Exception as e:
            print(f"[ERROR] Error respondiendo email: {e}")
            self.send_error_email(str(e), original_email, "Fallo al responder confirmarción")
            return False
            
    def _format_summary(self, data):
        """Formatea resumen de datos para el email"""
        summary = "\nDetalle:\n"
        for item in data[:10]:  # Máximo 10 items
            summary += f"  - SKU {item['Material']}: {item['Cantidad']} {item['Unidad']}\n"
        
        if len(data) > 10:
            summary += f"  ... y {len(data) - 10} más.\n"
        
        return summary
            
    def ask_user_confirmation(self, excel_path):
        """
        Muestra un popup nativo de Windows para que el usuario verifique el Excel.
        Retorna True si el usuario acepta, False si cancela.
        """
        print("[USER] Solicitando confirmación manual...")
        
        # Tipo de caja: Yes/No (MB_YESNO = 4) | Icono Info (MB_ICONINFORMATION = 0x40) | TopMost (MB_TOPMOST = 0x40000)
        MB_YESNO = 4
        MB_ICONINFORMATION = 0x40
        MB_TOPMOST = 0x40000
        
        message = (
            f"Se ha generado el archivo de carga:\n{excel_path}\n\n"
            "Por favor revise que los datos (Lotes, Cantidades) sean correctos en el Excel.\n\n"
            "¿Desea continuar con la carga en SAP?"
        )
        
        # MessageBoxW retorna: 6 (Yes), 7 (No)
        response = ctypes.windll.user32.MessageBoxW(0, message, "Verificación Requerida - JARVIS", MB_YESNO | MB_ICONINFORMATION | MB_TOPMOST)
        
        if response == 6: # IDYES
            print("[USER] Usuario confirmó la operación.")
            return True
        else:
            print("[USER] Operación cancelada por el usuario.")
            return False

    def run(self):
        """Flujo principal del bot"""
        print("=" * 60)
        print("[BOT] JARVIS - CONCILIACION PERDIDA DE VACIO v2.2 (Fix Loop)")
        print("=" * 60)
        
        # 0. CHECK CRÍTICO: Template disponible?
        if not os.path.exists(self.template_path):
             print(f"❌ [CRITICAL] Template no encontrado: {self.template_path}")
             return
             
        try:
            # Intentar abrir en modo 'append' solo para ver si está bloqueado
            with open(self.template_path, 'a+'):
                pass
        except PermissionError:
            print(f"❌ [CRITICAL] EL ARCHIVO TEMPLATE ESTÁ BLOQUEADO/ABIERTO.")
            print(f"   Ruta: {self.template_path}")
            print("   Por favor cierre el archivo Excel 'carga_migo.xlsx' y vuelva a intentar.")
            # Intentar cerrar forzosamente
            self.close_excel_if_open("carga_migo")
            # return # No retornamos aquí, intentaremos seguir
        except Exception as e:
            print(f"⚠️ [WARN] No se pudo verificar acceso al template: {e}")
        
        # 1. Conectar a Outlook
        if not self.connect_outlook():
            return
        
        # 2. Obtener carpeta
        folder = self.get_perdida_vacio_folder()
        if not folder:
            return
            
        processed_folder = self.get_or_create_processed_folder(folder)
        error_folder = self.get_or_create_error_folder(folder)
        
        # 3. Leer emails pendientes (cola)
        emails = self.read_pending_emails(folder)
        if not emails:
            print("ℹ️ Cola vacía. Todo está procesado.")
            return
        
        # 4. Procesar cada email
        # SAFETY LIMIT: Max 20 emails per run to prevent hanging/loops
        MAX_EMAILS_PER_RUN = 20
        emails_to_process = emails[:MAX_EMAILS_PER_RUN]
        
        print(f"\n[INFO] Iniciando procesamiento de {len(emails_to_process)} emails (Límite por ejecución: {MAX_EMAILS_PER_RUN})...")
        
        for i, email in enumerate(emails_to_process, start=1):
            try:
                print(f"\n{'='*60}")
                print(f"[EMAIL] Procesando email {i}/{len(emails_to_process)}")
                
                # Acceso seguro a propiedades del email
                try:
                    sender = email.SenderName
                    subject = email.Subject
                    received = email.ReceivedTime
                except:
                    print("[WARN] Error leyendo cabeceras del email. Saltando.")
                    continue
                
                print(f"   De: {sender}")
                print(f"   Asunto: {subject}")
                print(f"   Fecha: {received}")
                print(f"{'='*60}\n")

                # --- FILTRO 1: ASUNTO ---
                if "traspaso" not in subject.lower():
                    print("[SKIP] Saltando email (Asunto no contiene 'Traspaso')")
                    continue

                # --- PASO 2: PARSEAR DATOS ---
                data = self.parse_email_data(email.Body, getattr(email, 'HTMLBody', ''))
                
                if not data:
                    print("[WARN] No se encontraron datos válidos en el email. Moviendo a Errores para revisión manual.")
                    # Mover a errores porque no se pudo procesar y para evitar loop
                    if error_folder:
                        email.UnRead = True # Dejar como no leído para que resalte
                        email.Move(error_folder)
                    continue
                
                # --- PASO 3: VERIFICAR DUPLICADOS ---
                if self._is_duplicate_transfer(data):
                    print("[SKIP] Saltando email por ser traspaso duplicado")
                    try:
                        email.UnRead = False
                        if processed_folder:
                            email.Move(processed_folder)
                            print("[OK] Email duplicado movido a 'Procesados'")
                    except:
                        pass
                    continue
                
                # --- CREAR EXCEL (Solo Respaldo) ---
                success, output_excel = self.create_excel_from_template(data, None)
                
                if not success:
                    print("[ERROR] Falló creación Excel (Fatality). Moviendo email a ERRORES para evitar loop.")
                    self.send_error_email("No se pudo crear el Excel de carga. El email ha sido movido a la carpeta 'Errores' para evitar bloqueos.", email, "Error Excel")
                    if error_folder:
                        email.UnRead = True
                        email.Move(error_folder)
                    continue 

                # --- MODO ASISTIDO: SIN CONFIRMACIÓN PREVIA ---
                # user_confirmed = self.ask_user_confirmation(output_excel)
                # if not user_confirmed: ...
                
                print("[INFO] Modo Asistido: Saltando confirmación previa. Iniciando carga MIGO...")

                # --- PASO 5: EJECUTAR MIGO (Nativo) ---
                # Ahora pasamos 'data' directamente, el Excel es solo respaldo
                migo_success = self.execute_migo_native(data)
                
                if not migo_success:
                    print("[WARN] Error en traspaso MIGO. Se notificará pero el email requiere atención.")
                    self.send_error_email("Falló la ejecución de la transacción MIGO en SAP (Modo Nativo)", email, "Error SAP MIGO")
                    # Mover a errores para evitar loop
                    if error_folder:
                        email.Move(error_folder)
                    continue
                
                # --- PASO 6: FINALIZAR ---
                # Registrar en caché
                self._mark_transfer_processed(data)
                
                try:
                    email.UnRead = False
                    # Enviar confirmación antes de mover
                    self.send_confirmation_email(email, data)
                    
                    if processed_folder:
                        email.Move(processed_folder)
                        print("[OK] Email procesado y movido a 'Procesados'")
                except Exception as move_err:
                    print(f"[WARN] Error finalizando (mover/confirmar): {move_err}")

                print(f"\n[OK] Email procesado exitosamente\n")
            
            except Exception as e:
                print(f"[ERROR] Error inesperado procesando email: {e}")
                # SAFETY NET: Si falla algo inesperado, mover a errores para no loop
                self.send_error_email(f"Excepción no controlada: {e}", email, "Error Crítico Loop")
                try:
                    if error_folder:
                        email.Move(error_folder)
                except: pass
        
        print("=" * 60)
        print(f"[OK] Ciclo finalizado.")
        print("=" * 60)

if __name__ == "__main__":
    try:
        print(">>> Iniciando Bot_Conciliacion_Email...")
        bot = SapBotConciliacionEmail()
        bot.run()
    except Exception as e:
        print(f"!!! CRITICAL ERROR IN MAIN: {e}")
        import traceback
        traceback.print_exc()
    input("Presione Enter para salir...")
