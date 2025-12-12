import win32com.client
import pythoncom
import time
import os
import pyperclip 
from datetime import datetime 

class SapBotPallet:
    def run(self, ruta_excel=None):
        pythoncom.CoInitialize()
        print("--- BOT PALLET (LX02 -> Alineación Segura) ---")

        # MODO INTELIGENTE: Si no hay ruta o no existe, buscar Excel abierto
        excel = None
        wb = None
        
        if ruta_excel and os.path.exists(ruta_excel):
            # Modo tradicional: abrir archivo específico
            print(f"📂 Abriendo archivo: {ruta_excel}")
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = True
            excel.DisplayAlerts = False
            wb = excel.Workbooks.Open(ruta_excel)
        else:
            # Modo local: buscar archivo ya abierto
            print("🔍 Buscando archivo Excel abierto...")
            try:
                excel = win32com.client.GetObject(Class="Excel.Application")
                if excel.Workbooks.Count > 0:
                    # Buscar por nombre si se proporcionó
                    if ruta_excel:
                        nombre_buscar = os.path.basename(str(ruta_excel)).lower()
                        # Intento 1: Búsqueda exacta (coincidencia parcial de nombre)
                        for w in excel.Workbooks:
                            if nombre_buscar in w.Name.lower():
                                wb = w
                                print(f"✅ Encontrado (Nombre exacto): {w.Name}")
                                break
                        
                        # Intento 2: Búsqueda flexible (ignorando extensión, útil para .xlsm)
                        if not wb:
                            nombre_raiz = os.path.splitext(nombre_buscar)[0]
                            print(f"   -> Buscando variante flexible para: '{nombre_raiz}' ...")
                            for w in excel.Workbooks:
                                w_nombre_raiz = os.path.splitext(w.Name)[0].lower()
                                if nombre_raiz == w_nombre_raiz:
                                    wb = w
                                    print(f"✅ Encontrado (Nombre flexible): {w.Name}")
                                    break
                    
                    # Si no se encontró por nombre, usar el activo
                    if not wb:
                        if ruta_excel:
                            print(f"❌ Error: No se encontró ningún archivo abierto que contenga '{ruta_excel}' en el nombre.")
                            print("   -> Por seguridad, no se usará el archivo activo indiscriminadamente.")
                            return
                        else:
                            wb = excel.ActiveWorkbook
                            print(f"📊 Usando archivo activo (sin filtro de nombre): {wb.Name}")
                else:
                    print("❌ No hay archivos Excel abiertos.")
                    return
            except Exception as e:
                print(f"❌ Error: Excel no está abierto. {e}")
                return
        
        if not wb:
            print("❌ Error: No se pudo obtener archivo Excel.")
            return
            
        excel.Visible = True
        excel.DisplayAlerts = False

        # 1. CONEXIÓN SAP
        print("🔗 Conectando a SAP...")
        session = None
        try:
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
            app = SapGuiAuto.GetScriptingEngine
            connection = app.Children(0)
            session = connection.Children(0)
        except:
            print("❌ Error: No se detectó sesión SAP activa.")
            return

        try:
            # 2. EJECUCIÓN LX02
            print("🚀 Ejecutando LX02...")
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nLX02"
            session.findById("wnd[0]").sendVKey(0)
            
            print("   -> Rellenando filtros...")
            session.findById("wnd[0]/usr/ctxtS1_LGNUM").text = "ncd"
            session.findById("wnd[0]/usr/btn%_S1_LGTYP_%_APP_%-VALU_PUSH").press()
            
            # Tipos en mayúsculas y con retorno de carro Windows Standard (\r\n)
            tipos = ["PFW", "PBK", "CGO", "RCK"]
            texto_clip = "\r\n".join(tipos) 
            pyperclip.copy(texto_clip)
            print(f"   -> Tipos copiados al portapapeles: {tipos}")
            
            # Pegar desde portapapeles en SAP (Boton 24)
            session.findById("wnd[1]/tbar[0]/btn[24]").press() 
            session.findById("wnd[1]/tbar[0]/btn[8]").press()  
            
            print("   -> Ejecutando reporte (F8)...")
            session.findById("wnd[0]/tbar[1]/btn[8]").press()
            time.sleep(3)

            if "No existen datos" in session.findById("wnd[0]/sbar").text:
                print("⚠️ LX02 sin datos.")
                return

            # 3. COPIAR DATOS
            print("📋 Copiando datos...")
            try:
                session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                time.sleep(0.5)
                try: session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").select()
                except: session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").select()
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
            except:
                session.findById("wnd[0]/usr/cntlGRID1/shell").pressToolbarContextButton("&MB_COPY")
            
            time.sleep(1)
            session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            session.findById("wnd[0]").sendVKey(0)

            # (Procesamiento movido abajo para proteger portapapeles)

            # 5. PEGAR EN EXCEL (usando el workbook ya obtenido)
            print("📊 Pegando en Excel...")
            
            ws = None
            for s in wb.Sheets:
                if s.Name.lower() == "lx02":
                    ws = s
                    break
            
            if ws: 
                print("   -> Limpiando hoja existente (UsedRange)...")
                try:
                    # Usar UsedRange es más seguro que Cells (toda la hoja)
                    ws.UsedRange.ClearContents() 
                except Exception as e_clear:
                    print(f"⚠️ Advertencia al limpiar: {e_clear}")
            else: 
                print("   -> Creando nueva hoja 'lx02'...")
                ws = wb.Sheets.Add()
                ws.Name = "lx02"
            
            print("   -> Activando hoja...")
            ws.Activate()
            
            # 4. PROCESAR EN MEMORIA (PIPES A TABS) - AHORA AQUÍ
            print("🧠 Procesando en memoria...")
            try:
                raw_text = pyperclip.paste()
                clean_text = raw_text.replace('|', '\t')
                pyperclip.copy(clean_text)
            except Exception as e_clip:
                print(f"⚠️ Error procesando portapapeles: {e_clip}")

            # 5. PEGAR EN EXCEL (Con reintentos)
            print("📊 Intentando pegar...")
            for attempt in range(3):
                try:
                    # Pegar crudo en A1 (se separará por tabs automáticamente)
                    ws.Range("A1").PasteSpecial()
                    print("   -> Pegado exitoso.")
                    break
                except Exception as e:
                    print(f"⚠️ Fallo pegado (Intento {attempt+1}/3): {e}")
                    # Comprobación de errores comunes
                    if "-2147352567" in str(e) or "Source client" in str(e):
                        print("   -> Posible causa: Excel en 'Vista Protegida' o 'Edición no habilitada'.")
                    
                    time.sleep(1)
                    if attempt == 2: raise e
            
            # -----------------------------------------------------------
            # 6. ALINEACIÓN QUIRÚRGICA (MOVER DATOS SIN BORRAR FILAS)
            # -----------------------------------------------------------
            print("🧹 Alineando datos (sin romper fórmulas)...")
            
            last_row = ws.UsedRange.Rows.Count
            
            # A. ENCONTRAR DONDE ESTÁ EL ENCABEZADO REAL
            header_row = 0
            for i in range(1, 25):
                # Concatenamos texto de las primeras columnas para buscar palabras clave
                fila_txt = str(ws.Cells(i, 2).Value) + str(ws.Cells(i, 3).Value) + str(ws.Cells(i, 4).Value)
                if "Material" in fila_txt and "Lote" in fila_txt:
                    header_row = i
                    break
            
            if header_row > 0:
                print(f"   -> Encabezado encontrado en Fila {header_row}")
                
                # B. ENCONTRAR DONDE EMPIEZAN LOS DATOS (Saltar guiones)
                data_start_row = header_row + 1
                row_under_header = str(ws.Cells(data_start_row, 2).Value)
                if "---" in row_under_header or "___" in row_under_header:
                    data_start_row += 1 # Si hay guiones, los datos empiezan una fila más abajo
                
                print(f"   -> Datos reales empiezan en Fila {data_start_row}")

                # C. MOVER ENCABEZADO A FILA 1 (COPIAR Y PEGAR, NO DELETE)
                if header_row != 1:
                    # Copiar fila de encabezado encontrada -> Pegar en Fila 1
                    ws.Rows(header_row).Copy(ws.Rows(1))
                
                # D. MOVER DATOS A FILA 2 (COPIAR Y PEGAR)
                if data_start_row <= last_row:
                    # Rango origen: Desde el inicio de datos hasta el final
                    rng_datos = ws.Range(f"A{data_start_row}:XFD{last_row}")
                    rng_datos.Copy(ws.Range("A2")) # Pegar en A2
                    
                    # Calcular cuántas filas nos "sobran" abajo después de subir los datos
                    altura_datos = last_row - data_start_row + 1
                    fila_limpieza = 2 + altura_datos
                    
                    # Limpiar el residuo de abajo (datos duplicados que quedaron al fondo)
                    if fila_limpieza <= last_row:
                        print(f"   -> Limpiando residuos desde fila {fila_limpieza}...")
                        ws.Range(f"{fila_limpieza}:{last_row}").ClearContents()
                else:
                    print("   -> No hay datos debajo del encabezado.")

            # OCULTAR COLUMNA A (Para mantener estructura visual)
            ws.Columns("A:A").Hidden = True
            
            # Estampar información de ejecución en J1
            try:
                timestamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                # ws.Range("J1").Value = f"{wb.Name} - {timestamp}"
                # Ajustamos para que sea legible: Nombre archivo | Fecha Hora
                ws.Range("J1").Value = f"Fuente: {wb.Name} | Actualizado: {timestamp}"
                print(f"   -> Metadata escrita en J1: {wb.Name} | {timestamp}")
            except Exception as e_meta:
                print(f"⚠️ No se pudo escribir en J1: {e_meta}")

            ws.Columns.AutoFit()
            ws.Range("B1").Select()
            wb.Save()
            print("✅ ¡Listo! Datos alineados en Fila 1 y 2. Fórmulas seguras.")

        except Exception as e:
            print(f"❌ Error: {e}")
        
        finally:
            pythoncom.CoUninitialize()