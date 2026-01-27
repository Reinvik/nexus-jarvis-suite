import win32com.client
import pandas as pd
import os
import shutil
import pythoncom
from datetime import datetime
import tempfile
import sys

# Configurar salida estándar a UTF-8 para soportar emojis en Windows
try:
    if sys.stdout.encoding.lower() != 'utf-8':
        sys.stdout.reconfigure(encoding='utf-8')
except:
    pass

class BotConsolidacionZonales:
    def __init__(self):
        self.outlook = None
        self.namespace = None
        
        # Rutas
        self.onedrive_base = os.path.join(os.path.expanduser("~"), "OneDrive - CIAL Alimentos")
        # Usaremos la misma carpeta de operación que los otros bots para mantener orden
        self.target_folder = os.path.join(
            self.onedrive_base,
            "Archivos de Operación  Outbound CD - 16.-Inventario Critico"
        )
        self.master_file = os.path.join(self.target_folder, "Consolidado_Zonales_Master.xlsx")
        
        # Inicializar COM
        try:
            pythoncom.CoInitialize()
        except:
            pass
        
        # Mapeo de códigos de lote a Zonales
        self.ZONAL_MAP = {
            'ARSJ': {'zonal': 'Arica', 'almacen': 'ARSJ'},
            'IQSJ': {'zonal': 'Iquique', 'almacen': 'IQSJ'},
            'CLSJ': {'zonal': 'Calama', 'almacen': 'CLSJ'},
            'ANSJ': {'zonal': 'Antofagasta', 'almacen': 'ANSJ'},
            'CPSJ': {'zonal': 'Copiapó', 'almacen': 'CPSJ'},
            'LSSJ': {'zonal': 'La Serena', 'almacen': 'LSSJ'},
            'FLSJ': {'zonal': 'San Felipe', 'almacen': 'FLSJ'},
            'VMSJ': {'zonal': 'Viña del Mar', 'almacen': 'VMSJ'},
            'SASJ': {'zonal': 'San Antonio', 'almacen': 'SASJ'},
            'RGSJ': {'zonal': 'Rancagua', 'almacen': 'RGSJ'},
            'SFSJ': {'zonal': 'San Fernando', 'almacen': 'SFSJ'},
            'TLSJ': {'zonal': 'Talca', 'almacen': 'TLSJ'},
            'CHSJ': {'zonal': 'Chillán', 'almacen': 'CHSJ'},
            'CNSJ': {'zonal': 'Concepción', 'almacen': 'CNSJ'},
            'LASJ': {'zonal': 'Los Ángeles', 'almacen': 'LASJ'},
            'TMSJ': {'zonal': 'Temuco', 'almacen': 'TMSJ'},
            'OSSJ': {'zonal': 'Osorno', 'almacen': 'OSSJ'},
            'PMSJ': {'zonal': 'Puerto Montt', 'almacen': 'PMSJ'},
            'CYSJ': {'zonal': 'Coyhaique', 'almacen': 'CYSJ'},
            'PASJ': {'zonal': 'Punta Arenas', 'almacen': 'PASJ'},
        }

    def parse_date(self, value):
        """Convierte fechas a dd/mm/yyyy con corrección inteligente de año y soporte para números seriales de Excel"""
        import re
        from datetime import datetime, timedelta, time as dt_time
        import locale
        
        # Intentar configurar locale a español para reconocer meses
        try:
            locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
        except:
            try:
                locale.setlocale(locale.LC_TIME, 'Spanish_Spain.1252')
            except:
                pass

        if pd.isna(value) or str(value).strip() == '':
            return ''
        
        # Handle time objects explicitly (e.g. 00:00:00) -> Invalid date
        if isinstance(value, dt_time):
             return ''
             
        val_str = str(value).strip()
        current_now = datetime.now()
        dt = None
        
        # 0. Manejo de números seriales de Excel (ej: 45200)
        try:
            # Si es un número (int o float)
            # A veces viene como string "45200.0"
            if val_str.replace('.','',1).isdigit():
                serial = float(val_str)
                # Rango razonable: 30000 (1982) - 60000 (2064)
                # Si es muy bajo (ej: < 1), es probablemente una hora (ej: 0.5 = 12:00 PM)
                if serial > 30000 and serial < 60000: 
                    dt = datetime(1899, 12, 30) + timedelta(days=serial)
                    return dt.strftime('%d/%m/%Y')
        except:
            pass
        
        # 1. Si ya es datetime de pandas/python
        if isinstance(value, (datetime, pd.Timestamp)):
            dt = value
        else:
            # Check for year-only strings (e.g. "2024") -> Ambiguous, treat as invalid or handle?
            # User reported "errors where date comes out 2024". Assume we want to ignore single numbers that aren't serials.
            if val_str.isdigit() and len(val_str) == 4:
                # Is it a year? "2024". dateutil parses this as current date with year 2024.
                # If we consider this invalid data for a "Format Date" column:
                return '' 

            # 2. Patrones regex comunes
            patterns = [
                (r'(\d{4})-(\d{2})-(\d{2})', "%Y-%m-%d"),
                (r'(\d{2})-(\d{2})-(\d{4})', "%d-%m-%Y"),
                (r'(\d{2})/(\d{2})/(\d{4})', "%d/%m/%Y"),
                (r'(\d{2})\.(\d{2})\.(\d{4})', "%d.%m.%Y"),
                (r'(\d{4})/(\d{2})/(\d{2})', "%Y/%m/%d"),
            ]
            
            for pat, fmt in patterns:
                match = re.search(pat, val_str)
                if match:
                    try:
                        clean_val = match.group(0)
                        dt = datetime.strptime(clean_val, fmt)
                        
                        # --- LÓGICA INTELIGENTE BASADA EN MES ACTUAL ---
                        current_month = current_now.month
                        current_year = current_now.year
                        
                        # 1. CORRECCIÓN DE AMBIGÜEDAD DÍA/MES (Prioridad Mes Actual)
                        if fmt in ["%d-%m-%Y", "%d/%m/%Y", "%d.%m.%Y"]:
                            g1, g2 = int(match.group(1)), int(match.group(2))
                            
                            if dt.month != current_month and g1 == current_month:
                                try:
                                    new_dt = dt.replace(day=g2, month=g1)
                                    print(f"      🧠 Ambigüedad (Mes Actual): {clean_val} -> Interpretado como {new_dt.strftime('%d/%m/%Y')} (Coincide con mes {current_month})")
                                    dt = new_dt
                                except:
                                    pass
                            
                            elif dt.day != 1 and g2 == 1 and g1 == current_month:
                                try:
                                    new_dt = dt.replace(day=g2, month=g1)
                                    print(f"      🧠 Regla Día 1: {clean_val} -> Forzado a {new_dt.strftime('%d/%m/%Y')}")
                                    dt = new_dt
                                except:
                                    pass

                        # 2. CORRECCIÓN DE AÑO (Error de tipeo año anterior)
                        if dt.month == current_month and dt.year == current_year - 1:
                            print(f"      🧠 Corrección Año: {dt.strftime('%d/%m/%Y')} -> Actualizado a {current_year}")
                            dt = dt.replace(year=current_year)
                        # -----------------------------------------------
                        break
                    except:
                        continue
            
            # 3. Intentar parseo inteligente con dateutil
            if dt is None:
                try:
                    from dateutil import parser
                    dt = parser.parse(val_str, dayfirst=True, fuzzy=True)
                    
                    # Sanity check: If the parsed date is just time (has default date parts), dateutil usually uses current date.
                    # If val_str "00:00:00", dateutil returns today at 00:00:00.
                    # Heuristic: If val_str looks like a time and DOES NOT have date separators...
                    if ':' in val_str and not any(sep in val_str for sep in ['/', '-', '.']):
                         # Likely time only
                         return ''

                    if dt and dt.month == current_now.month and dt.year == current_now.year - 1:
                         dt = dt.replace(year=current_now.year)
                except:
                    pass

        # Si logramos obtener una fecha válida
        if dt:
            return dt.strftime('%d/%m/%Y')
            
        return val_str

    def extract_zonal_from_lote(self, lote):
        """Extrae código de zonal del número de lote"""
        if pd.isna(lote) or str(lote).strip() == '':
            return None, None
        
        lote_str = str(lote).upper().strip()
        
        # Buscar código de 4 letras que termine en SJ
        import re
        # Buscar patrones como ARSJ, TMSJ, etc.
        for code, info in self.ZONAL_MAP.items():
            if code in lote_str:
                return info['zonal'], info['almacen']
        
        # Buscar patrón de 4 letras mayúsculas que terminen en SJ
        match = re.search(r'([A-Z]{2}SJ)', lote_str)
        if match:
            code = match.group(1)
            if code in self.ZONAL_MAP:
                return self.ZONAL_MAP[code]['zonal'], self.ZONAL_MAP[code]['almacen']
        
        return None, None


    def close_excel_if_open(self):
        """Cierra el archivo maestro si está abierto en Excel"""
        try:
            excel = win32com.client.GetObject(Class="Excel.Application")
            for wb in excel.Workbooks:
                if os.path.basename(self.master_file).lower() in wb.FullName.lower():
                    print(f"⚠️ Archivo abierto en Excel. Cerrando: {wb.Name}")
                    wb.Close(SaveChanges=True)
                    print("✅ Archivo cerrado automáticamente")
                    # Pequeña pausa para que COM se estabilice
                    import time
                    time.sleep(1)
                    return True
        except:
            pass  # Excel no está corriendo o el archivo no está abierto
        return False

    def connect_outlook(self, max_retries=3):
        """Conecta a Outlook con reintentos"""
        import time
        
        for attempt in range(max_retries):
            try:
                self.outlook = win32com.client.Dispatch("Outlook.Application")
                self.namespace = self.outlook.GetNamespace("MAPI")
                
                # Intentar forzar sincronización y login
                try:
                    self.namespace.Logon("", "", False, False)
                    # Forzar Envío/Recepción
                    self.namespace.SendAndReceive(True)
                    print("   🔄 Sincronización forzada enviada...")
                    time.sleep(5) # Esperar a que sincronice un poco
                except:
                    pass
                    
                return True
            except Exception as e:
                if attempt < max_retries - 1:
                    print(f"   ⚠️ Reintentando conexión a Outlook... ({attempt + 1}/{max_retries})")
                    time.sleep(2)
                else:
                    print(f"❌ Error conectando a Outlook: {e}")
                    return False
        return False

    def get_folders(self, max_retries=3):
        """Obtiene carpeta Zonales y asegura que exista Procesados"""
        import time
        
        target_account = "Ariel.Mella@cial.cl"
        
        for attempt in range(max_retries):
            try:
                inbox = None
                root_folder = None
                
                # 1. Intentar buscar la cuenta específica
                for folder in self.namespace.Folders:
                    if target_account.lower() in folder.Name.lower():
                        root_folder = folder
                        print(f"📁 Cuenta detectada: {folder.Name}")
                        # Buscar Bandeja de entrada / Inbox dentro de la cuenta
                        for f in folder.Folders:
                            if f.Name.lower() in ["bandeja de entrada", "inbox"]:
                                inbox = f
                                print(f"   📂 Usando Inbox: {f.Name}")
                                break
                        break
                
                # 2. Fallback: Usar carpeta por defecto
                if not inbox:
                    print("⚠️ No se encontró cuenta específica, usando Inbox por defecto.")
                    inbox = self.namespace.GetDefaultFolder(6) # Inbox default
                    if not root_folder:
                        root_folder = inbox.Parent

                # Búsqueda exhaustiva
                candidates = ["Zonales_Bot", "Zonales"]
                print(f"   🔍 Buscando carpetas candidatas: {candidates}")
                
                possible_folders = []

                # A) Buscar en Inbox
                if inbox:
                    print(f"   📂 Listando subcarpetas en {inbox.Name}:")
                    for folder in inbox.Folders:
                        # Debug: Ver qué carpetas ve el bot
                        if "zonales" in folder.Name.lower():
                             print(f"      👀 Encontrada carpeta: '{folder.Name}' - Path: {folder.FolderPath} - Items: {folder.Items.Count}")

                        if folder.Name.lower() in [c.lower() for c in candidates]:
                            try:
                                count = folder.Items.Count
                                path = folder.FolderPath
                                print(f"      📂 Candidato encontrado en Inbox: '{path}' - Items: {count}")
                                possible_folders.append((folder, count, 1)) # Prioridad 1
                            except:
                                pass

                # B) Buscar en Raíz (Root)
                if root_folder:
                    for folder in root_folder.Folders:
                        if folder.Name.lower() in [c.lower() for c in candidates]:
                            try:
                                count = folder.Items.Count
                                path = folder.FolderPath
                                print(f"      📂 Candidato encontrado en Raíz: '{path}' - Items: {count}")
                                possible_folders.append((folder, count, 2)) # Prioridad 2
                            except:
                                pass

                # Selección del mejor candidato
                selected_folder = None
                
                # 1. Filtrar los que tienen items
                with_items = [x for x in possible_folders if x[1] > 0]
                
                if with_items:
                    # Si hay con items, tomamos el primero
                    selected_folder = with_items[0][0]
                    print(f"      ✅ Seleccionada por contenido: {selected_folder.Name} (Path: {selected_folder.FolderPath})")
                elif possible_folders:
                    selected_folder = possible_folders[0][0]
                    print(f"      ⚠️ Todas las carpetas vacías. Seleccionando: {selected_folder.Name} (Path: {selected_folder.FolderPath})")
                
                if not selected_folder:
                    print(f"❌ No se encontró la carpeta 'Zonales' ni 'Zonales_Bot' en ninguna ubicación.")
                    return None, None
                    
                # Buscar/Crear subcarpeta Procesados
                procesados_folder = None
                for folder in selected_folder.Folders:
                    if folder.Name == "Procesados":
                        procesados_folder = folder
                        break
                
                if not procesados_folder:
                    try:
                        procesados_folder = selected_folder.Folders.Add("Procesados")
                        print("📁 Carpeta 'Procesados' creada.")
                    except:
                        pass
                else:
                    print(f"      📂 Subcarpeta 'Procesados' encontrada (Items: {procesados_folder.Items.Count})")
                    
                return selected_folder, procesados_folder
                
            except Exception as e:
                if attempt < max_retries - 1:
                    print(f"   ⚠️ Outlook ocupado, reintentando... ({attempt + 1}/{max_retries})")
                    time.sleep(2)
                else:
                    print(f"❌ Error gestionando carpetas: {e}")
                    return None, None
        
        return None, None

    def clean_dataframe(self, df):
        """Limpia y estandariza el dataframe"""
        try:
            # 1. Buscar la fila del encabezado
            # Buscamos una fila que contenga "SKU" y "Zonal" (o similares)
            header_idx = -1
            for i, row in df.head(15).iterrows():
                row_str = row.astype(str).str.lower().values
                if 'sku' in row_str and ('zonal' in row_str or 'zona' in row_str):
                    header_idx = i
                    break
            
            # Si encontramos header, recargamos/ajustamos
            if header_idx > -1:
                # Promover fila a header
                df.columns = df.iloc[header_idx]
                df = df.iloc[header_idx + 1:].reset_index(drop=True)
            
            # 2. Eliminar columnas sin nombre (Unnamed)
            df = df.loc[:, ~df.columns.astype(str).str.contains('^Unnamed')]
            df = df.loc[:, df.columns.notna()] # Eliminar columnas NaN
            
            # --- NORMALIZACIÓN DE COLUMNAS ---
            # Renombrar columnas para estandarizar
            new_columns = {}
            for col in df.columns:
                col_str = str(col).strip()
                col_lower = col_str.lower()
                
                # Normalizar Zonal
                if col_lower == 'zonal':
                    new_columns[col] = 'Zonal'
                
                # Normalizar Fecha Dig-
                elif 'fecha' in col_lower and 'dig' in col_lower:
                    new_columns[col] = 'Fecha Dig-'
                    
                # Normalizar Almacen
                elif col_lower in ['almacen', 'alm', 'almacén']:
                    new_columns[col] = 'Almacen'
                    
            if new_columns:
                df = df.rename(columns=new_columns)
                print(f"      ✨ Columnas normalizadas: {list(new_columns.keys())} -> {list(new_columns.values())}")
            # ---------------------------------
            
            # 3. Eliminar filas vacías (basado en SKU)
            if 'SKU' in df.columns:
                df = df.dropna(subset=['SKU'])
            else:
                df = df.dropna(how='all')
            
            # Obtener lista de columnas para referencia por índice
            col_list = list(df.columns)
            
            # 4. Estandarizar fechas en columna O (índice 14) y R (índice 17)
            # Ahora buscamos por nombre si existe, sino por índice
            # Prioridad: Buscar columna "Fecha Dig-" y "Fecha Venc." o similar
            
            # Intentar encontrar columnas de fecha por nombre estandarizado
            col_fecha_dig = 'Fecha Dig-' if 'Fecha Dig-' in df.columns else None
            
            # Si encontramos Fecha Dig-, la estandarizamos
            if col_fecha_dig:
                try:
                    df[col_fecha_dig] = df[col_fecha_dig].apply(self.parse_date)
                    print(f"      📅 Fechas estandarizadas en columna '{col_fecha_dig}'")
                except:
                    pass

            # Mantener lógica por índice para otras fechas (como Vencimiento en columna O/14)
            # Columna O = índice 14
            if 14 < len(col_list):
                col_name_14 = col_list[14]
                # Solo si no es la misma que ya procesamos
                if col_name_14 != col_fecha_dig: 
                    try:
                        df[col_name_14] = df[col_name_14].apply(self.parse_date)
                    except:
                        pass
            
            # 5. Estandarizar Zonal (Primera mayúscula)
            # Buscar columna Zonal (puede variar mayúsculas/minúsculas)
            col_zonal = next((c for c in df.columns if str(c).lower() == 'zonal'), None)
            col_almacen = next((c for c in df.columns if str(c).lower() in ['almacen', 'alm', 'almacén']), None)
            col_lote = next((c for c in df.columns if str(c).lower() == 'lote'), None)
            
            if col_zonal:
                df[col_zonal] = df[col_zonal].astype(str).str.title().str.strip()
            
            # 6. Si Zonal o Almacén están vacíos, intentar extraer del Lote
            if col_lote:
                zonal_filled = 0
                for idx in df.index:
                    lote_val = df.at[idx, col_lote]
                    
                    # Verificar si Zonal está vacío
                    zonal_empty = (col_zonal is None or 
                                   pd.isna(df.at[idx, col_zonal]) or 
                                   str(df.at[idx, col_zonal]).strip() in ['', 'Nan', 'None'])
                    
                    # Verificar si Almacen está vacío
                    almacen_empty = (col_almacen is None or 
                                     col_almacen not in df.columns or
                                     pd.isna(df.at[idx, col_almacen]) or 
                                     str(df.at[idx, col_almacen]).strip() in ['', 'Nan', 'None'])
                    
                    if zonal_empty or almacen_empty:
                        zonal_from_lote, almacen_from_lote = self.extract_zonal_from_lote(lote_val)
                        
                        if zonal_from_lote and zonal_empty and col_zonal:
                            df.at[idx, col_zonal] = zonal_from_lote
                            zonal_filled += 1
                        
                        if almacen_from_lote and almacen_empty and col_almacen and col_almacen in df.columns:
                            df.at[idx, col_almacen] = almacen_from_lote
                
                if zonal_filled > 0:
                    print(f"      🔍 {zonal_filled} zonales extraídos desde Lote")
                
            return df
        except Exception as e:
            print(f"   ⚠️ Error limpiando dataframe: {e}")
            return df

    def process_attachment(self, attachment):
        """Procesa un archivo adjunto Excel y devuelve los dataframes"""
        temp_dir = tempfile.gettempdir()
        temp_path = os.path.join(temp_dir, attachment.FileName)
        
        try:
            attachment.SaveAsFile(temp_path)
            
            # Leer Excel
            xls = pd.ExcelFile(temp_path)
            
            df_faltantes = pd.DataFrame()
            df_sobrantes = pd.DataFrame()
            df_dano_mecanico = pd.DataFrame()
            df_transporte = pd.DataFrame() # Nuevo dataframe para Transporte
            
            # Buscar pestañas (flexible con mayúsculas/minúsculas)
            sheet_names = {name.lower(): name for name in xls.sheet_names}
            
            if 'faltantes' in sheet_names:
                df_raw = pd.read_excel(xls, sheet_name=sheet_names['faltantes'], header=None)
                df_faltantes = self.clean_dataframe(df_raw)
            
            if 'sobrantes' in sheet_names:
                df_raw = pd.read_excel(xls, sheet_name=sheet_names['sobrantes'], header=None)
                df_sobrantes = self.clean_dataframe(df_raw)
            
            # Buscar Daño Mecanico (flexible: "daño mecanico", "dano mecanico", etc.)
            for sheet_key in sheet_names:
                if 'da' in sheet_key and 'mec' in sheet_key:
                    df_raw = pd.read_excel(xls, sheet_name=sheet_names[sheet_key], header=None)
                    df_dano_mecanico = self.clean_dataframe(df_raw)
                    break

            # Buscar Transporte (flexible: "transporte", "hoja1", "transportes")
            # Prioridad: nombre exacto 'transporte' -> contiene 'transp'
            for sheet_key in sheet_names:
                if 'transp' in sheet_key:
                    df_raw = pd.read_excel(xls, sheet_name=sheet_names[sheet_key], header=None)
                    df_transporte = self.clean_dataframe(df_raw)
                    break
                
            return df_faltantes, df_sobrantes, df_dano_mecanico, df_transporte
            
        except Exception as e:
            print(f"   ⚠️ Error procesando adjunto {attachment.FileName}: {e}")
            return None, None, None, None
        finally:
            # Limpieza
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except:
                    pass

    def send_report_email(self, success_count, error_count, details):
        """Envía reporte de ejecución"""
        try:
            mail = self.outlook.CreateItem(0)
            mail.To = "ariel.mella@cial.cl"
            
            if error_count > 0:
                mail.Subject = f"⚠️ Reporte JARVIS (Zonales): {success_count} procesados, {error_count} errores"
            else:
                mail.Subject = f"✅ Reporte JARVIS (Zonales): {success_count} archivos consolidados"
                
            mail.Body = f"""Resumen de ejecución:
            
Archivos procesados correctamente: {success_count}
Errores encontrados: {error_count}

Detalles:
{details}

El archivo consolidado se encuentra en:
{self.master_file}

Atte.
JARVIS - Asistente de Automatización
"""
            mail.Send()
        except Exception as e:
            print(f"❌ Error enviando reporte: {e}")

    def run_once(self):
        """Ejecuta una iteración del proceso de consolidación"""
        print(f"\n⏰ Iniciando escaneo: {datetime.now().strftime('%H:%M:%S')}")
        
        if not self.connect_outlook():
            return
            
        zonales_folder, procesados_folder = self.get_folders()
        if not zonales_folder:
            return
        
        print(f"📁 Carpeta encontrada: {zonales_folder.Name}")
            
        messages = zonales_folder.Items
        # Usar Restrict para filtrar solo no leídos y forzar actualización de la vista
        try:
            # 1. Filtrar solo NO LEÍDOS
            unread_items = messages.Restrict("[UnRead] = True")
            unread_items.Sort("[ReceivedTime]", True) # Descendente: Más nuevos primero
            
            print(f"   📬 Total de emails en carpeta: {messages.Count}")
            print(f"   📧 No leídos detectados (RESTRICT): {unread_items.Count}")
            
            # 2. Iterar usando GetFirst/GetNext para mayor fiabilidad con COM
            unread_messages = []
            msg = unread_items.GetFirst()
            while msg:
                unread_messages.append(msg)
                msg = unread_items.GetNext()
                
            # 3. Para diagnóstico, revisar algunos leídos recientes con Excel
            # (Solo para log, no para procesar)
            read_with_excel = []
            # No iteramos todo folder.Items para esto, es muy lento. 
            
        except Exception as e:
            print(f"❌ Error obteniendo mensajes: {e}")
            return

        # Diagnóstico de lo que encontró
        if unread_messages:
            print(f"   📋 Emails no leídos encontrados (Top 3):")
            for i, m in enumerate(unread_messages[:3]):
                try:
                    print(f"      {i+1}. {m.Subject[:40]}... [{m.ReceivedTime}]")
                except:
                    pass
        else:
             print("   ⚠️ No se encontraron mensajes unread en la colección restringida.")
        
        if not unread_messages:
            print("\nℹ️ No hay correos nuevos (no leídos) en 'Zonales'.")
            
            # Mostrar los últimos 3 emails leídos con Excel por si sirve de referencia
            if read_with_excel:
                print("\n📋 Últimos emails YA PROCESADOS (leídos) con Excel adjunto:")
                for i, msg in enumerate(read_with_excel[:3]):
                    try:
                        print(f"   {i+1}. {msg.Subject[:50]}... ({msg.ReceivedTime.strftime('%d/%m %H:%M')})")
                    except:
                        pass
            return

        print(f"\n📧 Encontrados {len(unread_messages)} correos sin leer.")
        
        all_faltantes = []
        all_sobrantes = []
        all_dano_mecanico = []
        all_transportes = []
        
        processed_count = 0
        error_count = 0
        log_details = ""
        
        # Lista para guardar asuntos de correos de confirmación (NC lista)
        nc_confirmations = []
        
        for email in unread_messages:
            try:
                print(f"\nProcesando: {email.Subject}")
                sender = email.SenderName
                subject_clean = email.Subject.upper().replace("RE:", "").replace("RV:", "").strip()
                
                # --- LÓGICA SALOMON (NC LISTA) ---
                # Detectar si es correo de Salomon Acevedo
                if "ACEVEDO ACEVEDO" in sender.upper() or "SALOMON IVAN" in sender.upper():
                    print(f"   👤 Detectado correo de confirmación de NC (Salomon)")
                    nc_confirmations.append(subject_clean)
                    
                    # Marcar como procesado sin buscar Excel
                    email.UnRead = False
                    email.Move(procesados_folder)
                    processed_count += 1
                    log_details += f"✅ {email.Subject}: Confirmación NC procesada\n"
                    continue
                # ---------------------------------
                
                has_excel = False
                for attachment in email.Attachments:
                    # Agregar soporte para .xlsm
                    if attachment.FileName.lower().endswith(('.xlsx', '.xls', '.xlsm')):
                        has_excel = True
                        print(f"   📎 Adjunto: {attachment.FileName}")
                        
                        df_f, df_s, df_dm, df_tr = self.process_attachment(attachment)
                        
                        if df_f is not None and not df_f.empty:
                            df_f['Origen_Archivo'] = attachment.FileName
                            df_f['Origen_Email'] = sender
                            df_f['Asunto_Email'] = subject_clean # Guardamos asunto limpio para cruce
                            # Formato limpio de fecha: DD/MM/YYYY
                            try:
                                df_f['Fecha_Email'] = email.ReceivedTime.strftime("%d/%m/%Y")
                            except:
                                df_f['Fecha_Email'] = str(email.ReceivedTime)
                                
                            all_faltantes.append(df_f)
                            print(f"      ✓ Faltantes: {len(df_f)} filas")
                            
                        if df_s is not None and not df_s.empty:
                            df_s['Origen_Archivo'] = attachment.FileName
                            df_s['Origen_Email'] = sender
                            df_s['Asunto_Email'] = subject_clean
                            # Formato limpio de fecha: DD/MM/YYYY
                            try:
                                df_s['Fecha_Email'] = email.ReceivedTime.strftime("%d/%m/%Y")
                            except:
                                df_s['Fecha_Email'] = str(email.ReceivedTime)
                                
                            all_sobrantes.append(df_s)
                            print(f"      ✓ Sobrantes: {len(df_s)} filas")
                        
                        if df_dm is not None and not df_dm.empty:
                            df_dm['Origen_Archivo'] = attachment.FileName
                            df_dm['Origen_Email'] = sender
                            df_dm['Asunto_Email'] = subject_clean
                            try:
                                df_dm['Fecha_Email'] = email.ReceivedTime.strftime("%d/%m/%Y")
                            except:
                                df_dm['Fecha_Email'] = str(email.ReceivedTime)
                                
                            all_dano_mecanico.append(df_dm)
                            print(f"      ✓ Daño Mecánico: {len(df_dm)} filas")

                        if df_tr is not None and not df_tr.empty:
                            df_tr['Origen_Archivo'] = attachment.FileName
                            df_tr['Origen_Email'] = sender
                            df_tr['Asunto_Email'] = subject_clean
                            try:
                                df_tr['Fecha_Email'] = email.ReceivedTime.strftime("%d/%m/%Y")
                            except:
                                df_tr['Fecha_Email'] = str(email.ReceivedTime)
                                
                            all_transportes.append(df_tr)
                            print(f"      ✓ Transporte: {len(df_tr)} filas")
                
                if has_excel:
                    # Mover a procesados
                    email.UnRead = False
                    email.Move(procesados_folder)
                    processed_count += 1
                    log_details += f"✅ {email.Subject}: Procesado OK\n"
                else:
                    print("   ⚠️ No se encontraron adjuntos Excel.")
                    log_details += f"⚠️ {email.Subject}: Sin Excel adjunto\n"
                    
            except Exception as e:
                print(f"❌ Error procesando email: {e}")
                error_count += 1
                log_details += f"❌ {email.Subject}: Error - {e}\n"

        # Consolidar y Guardar
        # Procesamos si hay nuevos datos O si hay confirmaciones de NC
        if all_faltantes or all_sobrantes or all_dano_mecanico or all_transportes or nc_confirmations:
            print("\n💾 Guardando Consolidado...")
            try:
                # Cerrar el archivo si está abierto en Excel
                self.close_excel_if_open()
                
                # Leer datos existentes si el archivo ya existe
                existing_faltantes = pd.DataFrame()
                existing_sobrantes = pd.DataFrame()
                existing_dano_mecanico = pd.DataFrame()
                existing_transportes = pd.DataFrame()
                
                if os.path.exists(self.master_file):
                    print("   📂 Leyendo historial existente...")
                    try:
                        existing_faltantes = pd.read_excel(self.master_file, sheet_name='Faltantes')
                        print(f"      ✓ Faltantes históricos: {len(existing_faltantes)} filas")
                    except:
                        pass
                    try:
                        existing_sobrantes = pd.read_excel(self.master_file, sheet_name='Sobrantes')
                        print(f"      ✓ Sobrantes históricos: {len(existing_sobrantes)} filas")
                    except:
                        pass
                    try:
                        existing_dano_mecanico = pd.read_excel(self.master_file, sheet_name='Daño Mecanico')
                        print(f"      ✓ Daño Mecánico histórico: {len(existing_dano_mecanico)} filas")
                    except:
                        pass
                    try:
                        existing_transportes = pd.read_excel(self.master_file, sheet_name='Transportes')
                        print(f"      ✓ Transportes históricos: {len(existing_transportes)} filas")
                    except:
                        pass
                
                # --- ACTUALIZAR ESTADO NC EN HISTORIAL ---
                if nc_confirmations and not existing_faltantes.empty:
                    print(f"   🔄 Actualizando estados de NC para {len(nc_confirmations)} asuntos...")
                    
                    # Asegurar que existe la columna
                    if 'Estado_NC' not in existing_faltantes.columns:
                        existing_faltantes['Estado_NC'] = ""
                    
                    # Asegurar que existe Asunto_Email para comparar (si es archivo antiguo puede no tenerlo)
                    if 'Asunto_Email' in existing_faltantes.columns:
                        matches_found = 0
                        for conf_subj in nc_confirmations:
                            # Buscar coincidencias parciales en el asunto
                            # Normalizamos a mayúsculas y string para comparar
                            mask = existing_faltantes['Asunto_Email'].astype(str).str.upper().str.contains(conf_subj, regex=False)
                            
                            if mask.any():
                                existing_faltantes.loc[mask, 'Estado_NC'] = "NC lista"
                                matches_found += mask.sum()
                        
                        print(f"      ✨ Se marcaron {matches_found} filas como 'NC lista'")
                    else:
                        print("      ⚠️ El archivo histórico no tiene columna 'Asunto_Email', no se pueden cruzar las NC antiguas.")
                # -----------------------------------------

                # Concatenar nuevos datos con históricos
                new_faltantes = pd.concat(all_faltantes, ignore_index=True) if all_faltantes else pd.DataFrame()
                new_sobrantes = pd.concat(all_sobrantes, ignore_index=True) if all_sobrantes else pd.DataFrame()
                new_dano_mecanico = pd.concat(all_dano_mecanico, ignore_index=True) if all_dano_mecanico else pd.DataFrame()
                new_transportes = pd.concat(all_transportes, ignore_index=True) if all_transportes else pd.DataFrame()
                
                # --- AGREGAR FECHA DE INGRESO Y COLUMNA COLOR ---
                current_date = datetime.now().strftime("%d/%m/%Y %H:%M")
                
                if not new_faltantes.empty:
                    new_faltantes['Fecha_Agregado'] = current_date
                    if 'Estado_NC' not in new_faltantes.columns:
                        new_faltantes['Estado_NC'] = ""
                    if 'NC_Manual' not in new_faltantes.columns:
                        new_faltantes['NC_Manual'] = ""
                    if 'Procesado_Color' not in new_faltantes.columns:
                        new_faltantes['Procesado_Color'] = ""
                        
                if not new_sobrantes.empty:
                    new_sobrantes['Fecha_Agregado'] = current_date
                    if 'NC_Manual' not in new_sobrantes.columns:
                        new_sobrantes['NC_Manual'] = ""
                    if 'Procesado_Color' not in new_sobrantes.columns:
                        new_sobrantes['Procesado_Color'] = ""
                    
                if not new_dano_mecanico.empty:
                    new_dano_mecanico['Fecha_Agregado'] = current_date
                    if 'NC_Manual' not in new_dano_mecanico.columns:
                        new_dano_mecanico['NC_Manual'] = ""
                    if 'Procesado_Color' not in new_dano_mecanico.columns:
                        new_dano_mecanico['Procesado_Color'] = ""
                    
                if not new_transportes.empty:
                    new_transportes['Fecha_Agregado'] = current_date
                    if 'NC_Manual' not in new_transportes.columns:
                        new_transportes['NC_Manual'] = ""
                    if 'Procesado_Color' not in new_transportes.columns:
                        new_transportes['Procesado_Color'] = ""
                # --------------------------------
                # --------------------------------
                
                final_faltantes = pd.concat([existing_faltantes, new_faltantes], ignore_index=True)
                final_sobrantes = pd.concat([existing_sobrantes, new_sobrantes], ignore_index=True)
                final_dano_mecanico = pd.concat([existing_dano_mecanico, new_dano_mecanico], ignore_index=True)
                final_transportes = pd.concat([existing_transportes, new_transportes], ignore_index=True)
                
                # Eliminar duplicados (basado en todas las columnas excepto las de origen/fecha/estado)
                # Excluimos Estado_NC, Fecha_Email y Fecha_Agregado para encontrar duplicados reales del contenido
                
                if not final_faltantes.empty:
                    before = len(final_faltantes)
                    # Usamos subset para drop_duplicates, excluyendo metadatos variables
                    # IMPORTANTE: keep='first' conserva el registro ANTIGUO (con su fecha original)
                    subset_cols = [c for c in final_faltantes.columns if c not in ['Fecha_Email', 'Estado_NC', 'NC_Manual', 'Fecha_Agregado', 'Origen_Archivo', 'Origen_Email', 'Asunto_Email', 'Procesado_Color']]
                    final_faltantes = final_faltantes.drop_duplicates(subset=subset_cols, keep='first')
                    print(f"   🧹 Faltantes: {before} → {len(final_faltantes)} (eliminados {before - len(final_faltantes)} duplicados)")
                    
                if not final_sobrantes.empty:
                    before = len(final_sobrantes)
                    subset_cols = [c for c in final_sobrantes.columns if c not in ['Fecha_Email', 'NC_Manual', 'Fecha_Agregado', 'Origen_Archivo', 'Origen_Email', 'Asunto_Email', 'Procesado_Color']]
                    final_sobrantes = final_sobrantes.drop_duplicates(subset=subset_cols, keep='first')
                    print(f"   🧹 Sobrantes: {before} → {len(final_sobrantes)} (eliminados {before - len(final_sobrantes)} duplicados)")
                    
                if not final_dano_mecanico.empty:
                    before = len(final_dano_mecanico)
                    subset_cols = [c for c in final_dano_mecanico.columns if c not in ['Fecha_Email', 'NC_Manual', 'Fecha_Agregado', 'Origen_Archivo', 'Origen_Email', 'Asunto_Email', 'Procesado_Color']]
                    final_dano_mecanico = final_dano_mecanico.drop_duplicates(subset=subset_cols, keep='first')
                    print(f"   🧹 Daño Mecánico: {before} → {len(final_dano_mecanico)} (eliminados {before - len(final_dano_mecanico)} duplicados)")
                
                if not final_transportes.empty:
                    before = len(final_transportes)
                    subset_cols = [c for c in final_transportes.columns if c not in ['Fecha_Email', 'NC_Manual', 'Fecha_Agregado', 'Origen_Archivo', 'Origen_Email', 'Asunto_Email', 'Procesado_Color']]
                    final_transportes = final_transportes.drop_duplicates(subset=subset_cols, keep='first')
                    print(f"   🧹 Transportes: {before} → {len(final_transportes)} (eliminados {before - len(final_transportes)} duplicados)")

                # Guardar
                with pd.ExcelWriter(self.master_file, engine='openpyxl') as writer:
                    final_faltantes.to_excel(writer, sheet_name='Faltantes', index=False)
                    final_sobrantes.to_excel(writer, sheet_name='Sobrantes', index=False)
                    final_dano_mecanico.to_excel(writer, sheet_name='Daño Mecanico', index=False)
                    final_transportes.to_excel(writer, sheet_name='Transportes', index=False)
                    
                print(f"✅ Archivo Maestro actualizado: {self.master_file}")
                print(f"   📊 Totales: Faltantes={len(final_faltantes)}, Sobrantes={len(final_sobrantes)}, Daño={len(final_dano_mecanico)}, Transp={len(final_transportes)}")
                self.send_report_email(processed_count, error_count, log_details)
                
            except Exception as e:
                print(f"❌ Error guardando Excel Maestro: {e}")
        else:
            print("\nℹ️ No se extrajeron datos ni confirmaciones para consolidar.")

    def open_excel_file(self):
        """Abre el archivo maestro en Excel para verificación"""
        try:
            if os.path.exists(self.master_file):
                print(f"📂 Abriendo archivo para verificación...")
                os.startfile(self.master_file)
                print(f"✅ Archivo abierto: {os.path.basename(self.master_file)}")
                return True
        except Exception as e:
            print(f"⚠️ No se pudo abrir el archivo: {e}")
        return False

    def run(self):
        """Ejecuta el proceso de consolidación una sola vez"""
        print("="*60)
        print("🤖 JARVIS - CONSOLIDADOR DE ZONALES")
        print("="*60)
        
        # Cerrar Excel si está abierto antes de procesar
        was_open = self.close_excel_if_open()
        
        try:
            self.run_once()
            print("\n✅ Proceso completado.")
            
            # Abrir el archivo para verificación
            if os.path.exists(self.master_file):
                print("\n📊 Abriendo archivo consolidado para verificación...")
                self.open_excel_file()
                
        except Exception as e:
            print(f"❌ Error en el proceso: {e}")

if __name__ == "__main__":
    bot = BotConsolidacionZonales()
    bot.run()

