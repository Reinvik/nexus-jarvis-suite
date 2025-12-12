import win32com.client
import os
import sys
import datetime
import time
import ctypes

# Force UTF-8 for console output
if sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

class BotReporteCambiados:
    def __init__(self):
        self.excel_name = "Reporte Desv. Zonales, prod cambiados y SDIF semana actual.xlsm"
        self.sheet_name = "Cambiado"
        self.onedrive_path = os.path.join(os.path.expanduser("~"), "OneDrive - CIAL Alimentos")
        
        # Search for the file if logic needs it, or strictly define it
        # Based on search, it's likely inthe root or a known folder.
        # We will try to find it.
        self.target_path = self.find_target_file()

    def find_target_file(self):
        """Locates the specific Excel file in OneDrive"""
        print(f"🔍 Buscando archivo: {self.excel_name}...")
        
        # 1. Check known paths
        candidates = [
            os.path.join(self.onedrive_path, self.excel_name),
            os.path.join(self.onedrive_path, "Archivos de chat de Microsoft Teams", self.excel_name),
             os.path.join(self.onedrive_path, "Escritorio", self.excel_name)
        ]
        
        for path in candidates:
            if os.path.exists(path):
                print(f"✅ Archivo encontrado: {path}")
                return path
                
        # 2. Walk search (Fallback)
        print("⚠️ No encontrado en rutas comunes. Escaneando OneDrive (esto puede tardar)...")
        for root, dirs, files in os.walk(self.onedrive_path):
            if self.excel_name in files:
                path = os.path.join(root, self.excel_name)
                print(f"✅ Archivo encontrado: {path}")
                return path
        
        print("❌ No se encontró el archivo Excel objetivo.")
        return None

    def get_excel_app(self):
        """Connects to or starts Excel"""
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
            print("📎 Conectado a instancia existente de Excel.")
        except:
            excel = win32com.client.Dispatch("Excel.Application")
            print("🚀 Iniciando nueva instancia de Excel.")
        
        excel.Visible = True # User wanted to see it
        return excel

    def run(self):
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except:
            pass

        if not self.target_path:
            ctypes.windll.user32.MessageBoxW(0, f"No se encontró el archivo:\n{self.excel_name}", "Error - Nexus Jarvis", 0x10)
            return

        excel = self.get_excel_app()
        wb = None
        
        try:
            # 1. Open/Activate Workbook
            file_open = False
            for w in excel.Workbooks:
                if w.Name == self.excel_name:
                    wb = w
                    file_open = True
                    break
            
            if not file_open:
                print(f"📂 Abriendo archivo...")
                wb = excel.Workbooks.Open(self.target_path)
            
            wb.Activate()
            
            # 2. Select Sheet
            try:
                ws = wb.Sheets(self.sheet_name)
                ws.Select()
            except Exception as e:
                print(f"❌ Error: No existe la hoja '{self.sheet_name}'.")
                ctypes.windll.user32.MessageBoxW(0, f"No se encontró la pestaña '{self.sheet_name}' en el Excel.", "Error - Nexus Jarvis", 0x10)
                return

            # 3. Filter Logic
            # User said "una vez actualizado el filtro a las 12...". 
            # We will try to filter non-blanks on Column D (SKU Faltan) as a helper
            # But the user might want to check it.
            # Let's try to verify if it's filtered.
            
            print("⚙️ Verificando filtros...")
            # Based on user image, Header is likely Row 4
            HEADER_ROW = 4
            
            if not ws.AutoFilterMode:
                 ws.Range(f"A{HEADER_ROW}").AutoFilter() 
            
            # Apply Filter: Field 4 (SKU Faltan) <> Empty
            try:
                # Field 4 is Column D ("SKU Faltar")
                ws.Range(f"A{HEADER_ROW}").AutoFilter(Field=4, Criteria1="<>")
            except:
                pass
            
            time.sleep(1)

            # 4. Copy Visible Range (Including Title at Row 2)
            last_row = ws.Cells(ws.Rows.Count, "A").End(-4162).Row
            
            # Check if there are visible data rows > HEADER_ROW
            # We assume if last_row > HEADER_ROW, there might be data.
            # But filter might hide everything.
            
            rng_data = ws.Range(f"A{HEADER_ROW+1}:A{last_row}")
            try:
                visible_cells = rng_data.SpecialCells(12).Count
            except:
                visible_cells = 0
                
            if visible_cells == 0:
                print("⚠️ No hay datos visibles para reportar.")
                # We show message but don't return, maybe user wants the empty template?
                # User said "copie lo de la foto". If photo has data, we expect data.
                ctypes.windll.user32.MessageBoxW(0, "No se encontraron datos después de filtrar.\nVerifique la hoja 'Cambiado'.", "Alerta", 0x30)
                return

            # Define Range to Export (Title + Data)
            # A2:L{last_row}
            copy_rng = ws.Range(f"A2:L{last_row}")

            # 4. SANITIZAR SELECCIÓN (Contiguous Hop Method)
            # Strategy: Filtered Range -> Paste to Temp WB (Becomes Contiguous) -> Copy Temp -> Paste to Outlook
            # This solves "Disjoint Range" issues that break standard Copy/Paste.
            
            print("Processing: Sanitizando rango en libro temporal...")
            
            # A. Copy Filtered Source
            wb.Activate()
            copy_rng.Select()
            win32com.client.Dispatch("WScript.Shell").SendKeys("^c")
            copy_rng.Copy()
            time.sleep(1)
            
            # B. Paste to Temp Workbook
            excel.DisplayAlerts = False
            temp_wb = excel.Workbooks.Add()
            temp_ws = temp_wb.Sheets(1)
            temp_ws.Activate()
            temp_ws.Range("A1").Select()
            temp_ws.Paste() # Makes it contiguous
            time.sleep(0.5)
            
            try:
                temp_ws.Range("A1").PasteSpecial(Paste=8) # Column Widths
            except:
                pass
                
            # C. Copy CLEAN Range
            # Now we have a simple contiguous block starting at A1
            clean_rng = temp_ws.UsedRange
            print(f"📋 Copiando rango limpio: {clean_rng.Address} ({clean_rng.Rows.Count} filas)")
            
            clean_rng.Copy() # Copy to clipboard for real this time
            time.sleep(1)
            
            # Close Temp (Don't need it anymore, data is in clipboard)
            temp_wb.Close(SaveChanges=False)
            excel.DisplayAlerts = True

            # 5. Create Email & Paste
            self.create_email_paste_method()

            print("✅ Proceso finalizado.")

        except Exception as e:
            print(f"❌ Error crítico: {e}")
            import traceback
            traceback.print_exc()
            ctypes.windll.user32.MessageBoxW(0, f"Ocurrió un error:\n{e}", "Error - Nexus Jarvis", 0x10)

    def create_email_paste_method(self):
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            
            today_str = datetime.datetime.now().strftime("%d-%m-%Y")
            mail.Subject = f"Informe de Revisión de productos cambiados zonales al {today_str}"
            
            # Recipients
            recipients = [
                "jefesoperacioneszonales@cialalimentos.cl",
                "cesar.esveile@cial.cl",
                "euro.velasquez@cial.cl",
                "Alejandro.Ureta@cial.cl",
                "francisco.lara@cial.cl",
                "Irene.espina@cial.cl",
                "gonzalo.tello@cial.cl",
                "controldeexistencias@cialalimentos.cl"
            ]
            mail.To = "; ".join(recipients)
            
            mail.Display()
            time.sleep(1)
            
            inspector = mail.GetInspector
            word_doc = inspector.WordEditor
            
            # Header
            header_text = "Buenas tardes Jefes de operaciones:\n\nAdjunto revisión y conciliación de los productos cambiados en los informes.\nTodos los productos concuerdan con la política de conciliación, aplica cargo.\n\n"
            word_doc.Range(0, 0).InsertBefore(header_text)
            
            # Paste Position
            doc_end = word_doc.Range()
            doc_end.Collapse(0)
            doc_end.InsertParagraphAfter()
            doc_end.Collapse(0)
            
            # PASTE (COM Strategy with HTML Format)
            print("📋 Pegando tabla (COM HTML)...")
            
            # Use PasteSpecial with HTML format (10) or WD_PasteDataType.wdPasteHTML
            # Since we cleaned the range in Temp WB, COM should handle it now.
            
            # Ensure Outlook is ready
            inspector.Activate()
            
            # Retry mechanism for Paste
            pasted = False
            try:
                # Move to end
                word_app = word_doc.Application
                word_app.Selection.EndKey(6)
                word_app.Selection.TypeParagraph()
                
                # Try standard Paste first
                word_app.Selection.Paste()
                pasted = True
                print("✅ Pegado estándar exitoso.")
            except Exception as e:
                print(f"⚠️ Falló pegado estándar: {e}")
            
            if not pasted:
                try:
                    # Try PasteSpecial HTML (10)
                    word_app.Selection.PasteSpecial(DataType=10, Placement=0, DisplayAsIcon=False)
                    pasted = True
                    print("✅ Pegado HTML exitoso.")
                except Exception as e:
                    print(f"⚠️ Falló pegado HTML: {e}")
            
            # Fallback to SendKeys if COM totally dies
            if not pasted:
                print("⚠️ Intentando fallback SendKeys...")
                shell = win32com.client.Dispatch("WScript.Shell")
                shell.SendKeys("^v")
                time.sleep(1)

            # Signature
            sig_text = "\n\nAtte.\nJARVIS - Asistente de Automatización de\n\nAriel Mella - Analista de Inventario"
            try:
                # Move to end again
                word_app.Selection.EndKey(6)
                word_app.Selection.TypeParagraph()
                word_app.Selection.TypeText(sig_text)
            except:
                pass
                
        except Exception as e:
            print(f"❌ Error creando email: {e}")

            print("✅ Proceso finalizado.")

        except Exception as e:
            print(f"❌ Error crítico: {e}")
            import traceback
            traceback.print_exc()
            ctypes.windll.user32.MessageBoxW(0, f"Ocurrió un error:\n{e}", "Error - Nexus Jarvis", 0x10)

    def create_email_with_html(self, table_html):
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            
            today_str = datetime.datetime.now().strftime("%d-%m-%Y")
            mail.Subject = f"Informe de Revisión de productos cambiados zonales al {today_str}"
            
            # Recipients
            recipients = [
                "jefesoperacioneszonales@cialalimentos.cl",
                "cesar.esveile@cial.cl",
                "euro.velasquez@cial.cl",
                "Alejandro.Ureta@cial.cl",
                "francisco.lara@cial.cl",
                "Irene.espina@cial.cl",
                "gonzalo.tello@cial.cl",
                "controldeexistencias@cialalimentos.cl"
            ]
            mail.To = "; ".join(recipients)
            
            # Construct complete HTML Body
            # Clean up Excel's HTML slightly if needed (Excel adds huge headers)
            # Usually we just take the body or the div.
            # But putting it all in is usually fine for Outlook.
            
            header = (
                "<p style='font-family:Calibri,sans-serif;font-size:11pt;color:black;'>"
                "Buenas tardes Jefes de operaciones:<br><br>"
                "Adjunto revisión y conciliación de los productos cambiados en los informes.<br>"
                "Todos los productos concuerdan con la política de conciliación, aplica cargo.<br><br>"
                "</p>"
            )
            
            signature = (
                "<p style='font-family:Calibri,sans-serif;font-size:11pt;color:black;'>"
                "<br>Atte.<br>"
                "JARVIS - Asistente de Automatización de<br><br>"
                "Ariel Mella - Analista de Inventario"
                "</p>"
            )
            
            # Combine
            mail.HTMLBody = header + table_html + signature
            
            mail.Display()
            print("✅ Email generado correctamente (Método HTML).")
            
        except Exception as e:
            print(f"❌ Error creando email: {e}")

if __name__ == "__main__":
    bot = BotReporteCambiados()
    bot.run()
