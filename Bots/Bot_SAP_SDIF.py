import win32com.client
import sys
import datetime
import time
import pythoncom

class SapBotSDIF:
    def __init__(self):
        self.session = None

    def conectar_sap(self):
        """Conexión robusta basada en Bot_Auditor"""
        try:
            pythoncom.CoInitialize()
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
            if not SapGuiAuto: raise Exception("No SAPGUI")
            
            application = SapGuiAuto.GetScriptingEngine
            connection = application.Children(0)
            session = connection.Children(0)
            self.session = session
            print("✅ Conectado a SAP")
            return True
        except Exception as e:
            print(f"❌ Error SAP: {e}")
            return False

    def obtener_excel_abierto(self):
        """Conecta al Excel que llamó al script"""
        try:
            excel = win32com.client.GetObject(Class="Excel.Application")
            return excel
        except Exception as e:
            print(f"❌ Error Excel: {e}")
            return None

    def run(self):
        if not self.conectar_sap(): return

        excel = self.obtener_excel_abierto()
        if not excel: return
        
        try:
            wb = excel.ActiveWorkbook
            print(f"📂 Trabajando en: {wb.Name}")
            
            # --- 1. MB51 ---
            print("RUNNING MB51...")
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nMB51"
            self.session.findById("wnd[0]").sendVKey(0)
            
            # LIMPIAR CAMPOS MB51
            try:
                self.session.findById("wnd[0]/usr/ctxtMATNR-LOW").text = ""
                self.session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = ""
                self.session.findById("wnd[0]/usr/ctxtLGORT-LOW").text = ""
                self.session.findById("wnd[0]/usr/ctxtCHARG-LOW").text = ""
                self.session.findById("wnd[0]/usr/ctxtBWART-LOW").text = ""
                self.session.findById("wnd[0]/usr/ctxtUSNAM-LOW").text = ""
                self.session.findById("wnd[0]/usr/ctxtBUDAT-LOW").text = ""
                self.session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").text = ""
            except: pass

            # Fechas (Ultimos 3 dias)
            hoy = datetime.date.today()
            inicio = hoy - datetime.timedelta(days=3)
            self.session.findById("wnd[0]/usr/ctxtBUDAT-LOW").text = inicio.strftime("%d.%m.%Y")
            self.session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").text = hoy.strftime("%d.%m.%Y")
            
            # Almacen SDIF
            self.session.findById("wnd[0]/usr/ctxtLGORT-LOW").text = "SDIF"
            
            # Ejecutar
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            
            # Exportar a Clipboard
            try:
                 # Intentar exportar
                self.session.findById("wnd[0]/tbar[1]/btn[45]").press() # Boton detalles/lista
                # Ojo: La exportación exacta depende de la vista. Asumimos Lista.
                # LISTA -> EXPORTAR -> FICHERO LOCAL -> PORTAPAPELES
                self.session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select() # Lista -> Exportar -> FicheroLocal
                time.sleep(1)
                self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").select() # Clipboard
                self.session.findById("wnd[1]/tbar[0]/btn[0]").press() # Check
                time.sleep(1)
            except:
                print("⚠️ Error exportando MB51 (Check script steps)")

            # Pegar en Excel MB51
            ws_mb51 = wb.Sheets("MB51")
            last_row = ws_mb51.Cells(ws_mb51.Rows.Count, "G").End(3).Row + 1 # 3=xlUp
            # Pegar en Col G (7)
            ws_mb51.Cells(last_row, 7).PasteSpecial()
            
            # ELIMINAR DUPLICADOS DESACTIVADO POR SOLICITUD DE USUARIO
            # try:
            #     tbl = ws_mb51.ListObjects("TablaSDIF")
            #     tbl.Range.RemoveDuplicates(Columns=(1,2,3,4,5,6,7), Header=1) 
            # except:
            #     pass


            # --- 2. MB52 ---
            print("RUNNING MB52...")
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nMB52"
            self.session.findById("wnd[0]").sendVKey(0)

            # LIMPIAR CAMPOS MB52
            try:
                self.session.findById("wnd[0]/usr/ctxtMATNR-LOW").text = ""
                self.session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = ""
                self.session.findById("wnd[0]/usr/ctxtLGORT-LOW").text = ""
                self.session.findById("wnd[0]/usr/ctxtCHARG-LOW").text = ""
                # self.session.findById("wnd[0]/usr/ctxtMATKL-LOW").text = "" # Opcional
            except: pass

            self.session.findById("wnd[0]/usr/ctxtLGORT-LOW").text = "SDIF"
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            
            # Exportar Clipboard
            self.session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
            time.sleep(1)
            self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").select()
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            time.sleep(1)

            # Pegar en Excel MB52
            ws_mb52 = wb.Sheets("MB52")
            last_row_52 = ws_mb52.Cells(ws_mb52.Rows.Count, "C").End(3).Row + 1
            ws_mb52.Cells(last_row_52, 3).PasteSpecial()

            # --- 3. Refresh ---
            wb.Sheets("SDIF").PivotTables(1).RefreshTable()
            print("✅ Proceso terminado")

        except Exception as e:
            print(f"❌ Error en Loop: {e}")

if __name__ == "__main__":
    bot = SapBotSDIF()
    bot.run()
