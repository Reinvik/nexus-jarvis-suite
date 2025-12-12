import win32com.client
import pandas as pd
import os
from datetime import date, timedelta, datetime
import pythoncom
import time
import tkinter as tk
from tkinter import simpledialog, messagebox

class SapBotTransporte:
    def check_file_open(self, filepath):
        if not os.path.exists(filepath): return False
        try:
            os.rename(filepath, filepath)
            return False
        except OSError: return True

    def run(self, fechas=None, enviar_correo=False):
        pythoncom.CoInitialize()
        print("--- ROBOT TRANSPORTE (VT11 -> VT03N) ---")
        
        ruta_base = os.path.join(os.path.expanduser("~"), r"OneDrive - CIAL Alimentos\Archivos de Operación  Outbound CD - 16.-Inventario Critico")
        archivo_master = os.path.join(ruta_base, "Base_Datos_Logistica.xlsx")

        if self.check_file_open(archivo_master):
            root = tk.Tk(); root.withdraw(); root.attributes("-topmost", True)
            messagebox.showerror("Error", "El archivo Excel está abierto. Ciérralo y reintenta.")
            root.destroy()
            return

        try:
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
            app = SapGuiAuto.GetScriptingEngine
            session = app.Children(0).Children(0)
        except:
            print("❌ Error: SAP no está abierto.")
            return

        # --- FECHAS ---
        fecha_hoy = date.today().strftime("%d.%m.%Y")
        fecha_ayer = (date.today() - timedelta(days=1)).strftime("%d.%m.%Y")
        
        # Si las fechas vienen como parámetro desde la interfaz web, usarlas
        if fechas:
            print(f"📅 Fechas recibidas desde interfaz: {fechas}")
            input_fechas = fechas
        else:
            # Solo pedir fechas si no vienen como parámetro
            root = tk.Tk(); root.withdraw(); root.attributes("-topmost", True)
            print("⏳ Fechas...")
            input_fechas = simpledialog.askstring("Fechas VT11", f"Rango (Inicio-Fin) o Enter para {fecha_ayer}-{fecha_hoy}", parent=root)
            root.destroy()

        f_inicio, f_fin = fecha_ayer, fecha_hoy
        if input_fechas and "-" in input_fechas:
            parts = input_fechas.split("-")
            if len(parts)==2: f_inicio, f_fin = parts[0].strip(), parts[1].strip()
        elif input_fechas and len(input_fechas)>6: f_inicio = f_fin = input_fechas.strip()

        # --- VT11 ---
        print("📥 Ejecutando VT11...")
        try:
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nVT11"
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/ctxtK_SHTYP-LOW").text = "ZTZN"
            session.findById("wnd[0]/usr/ctxtK_TPLST-LOW").text = "sgsj"
            session.findById("wnd[0]/usr/ctxtK_ERDAT-LOW").text = f_inicio
            session.findById("wnd[0]/usr/ctxtK_ERDAT-HIGH").text = f_fin
            session.findById("wnd[0]/tbar[1]/btn[8]").press()
        except Exception as e:
            print(f"❌ Error VT11: {e}")
            return

        # Lectura con scroll
        lista_tknum = []
        intentos_sin_datos = 0
        print("📊 Leyendo lista...")
        
        while intentos_sin_datos < 3:
            datos_en_pantalla = 0
            
            # Leer hasta 34 filas visibles
            for i in range(34):
                try:
                    tknum = ""
                    try: tknum = session.findById(f"wnd[0]/usr/lbl[8,{4+i}]").text.strip()
                    except: 
                        try: tknum = session.findById(f"wnd[0]/usr/lbl[1,{4+i}]").text.strip()
                        except: pass

                    if tknum and tknum.isdigit() and len(tknum) >= 7:
                        if tknum not in lista_tknum:
                            lista_tknum.append(tknum)
                            datos_en_pantalla += 1
                except: pass
            
            # Si encontramos datos, hacer scroll; si no, incrementar contador
            if datos_en_pantalla > 0:
                intentos_sin_datos = 0
                try:
                    # Hacer scroll hacia abajo (Page Down)
                    session.findById("wnd[0]").sendVKey(82)
                except: break
            else:
                intentos_sin_datos += 1

        print(f"✅ Transportes: {len(lista_tknum)}")
        if not lista_tknum: return

        # --- VT03N ---
        datos_finales = []
        timestamp = datetime.now().strftime("%d-%m-%Y %H:%M")

        for i, tknum in enumerate(lista_tknum):
            print(f"   [{i+1}/{len(lista_tknum)}] Procesando {tknum}...", end="\r")
            try:
                session.findById("wnd[0]/tbar[0]/okcd").text = "/nVT03N"
                session.findById("wnd[0]").sendVKey(0)
                session.findById("wnd[0]/usr/ctxtVTTK-TKNUM").text = tknum
                session.findById("wnd[0]").sendVKey(0)

                # 1. FECHAS (Horas)
                f_reg, h_reg = "", ""
                f_ini_carga, h_ini_carga = "", ""
                f_fin_carga, h_fin_carga = "", ""
                f_despacho, h_despacho = "", ""
                f_ini_transp, h_ini_transp = "", ""
                f_fin_transp, h_fin_transp = "", ""
                
                for idx in [2, 1, 3]:
                    try:
                        session.findById(f"wnd[0]/usr/tabsHEADER_TABSTRIP{idx}/tabpTABS_OV_DE").select()
                        base = f"wnd[0]/usr/tabsHEADER_TABSTRIP{idx}/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN{idx}:SAPMV56A:1025/"
                        
                        # Registro
                        try: f_reg = session.findById(base + "ctxtVTTK-DATREG").text; h_reg = session.findById(base + "ctxtVTTK-UAREG").text
                        except: pass
                        
                        # Inicio Carga
                        try: f_ini_carga = session.findById(base + "ctxtVTTK-DALBG").text; h_ini_carga = session.findById(base + "ctxtVTTK-UALBG").text
                        except: pass
                        
                        # Fin Carga
                        try: f_fin_carga = session.findById(base + "ctxtVTTK-DALEN").text; h_fin_carga = session.findById(base + "ctxtVTTK-UALEN").text
                        except: pass
                        
                        # Despacho (Expedición)
                        try: f_despacho = session.findById(base + "ctxtVTTK-DALAB").text; h_despacho = session.findById(base + "ctxtVTTK-UALAB").text
                        except: pass
                        
                        # Inicio Transporte
                        try: f_ini_transp = session.findById(base + "ctxtVTTK-DATBG").text; h_ini_transp = session.findById(base + "ctxtVTTK-UATBG").text
                        except: pass
                        
                        # Fin Transporte
                        try: f_fin_transp = session.findById(base + "ctxtVTTK-DATEN").text; h_fin_transp = session.findById(base + "ctxtVTTK-UATEN").text
                        except: pass

                        if f_reg or f_ini_transp: break
                    except: pass

                # 2. RESUMEN (Ruta, Signatura)
                ruta, signatura = "", ""
                for idx in [1, 2]:
                    try:
                        session.findById(f"wnd[0]/usr/tabsHEADER_TABSTRIP{idx}/tabpTABS_OV_PR").select()
                        base = f"wnd[0]/usr/tabsHEADER_TABSTRIP{idx}/tabpTABS_OV_PR/ssubG_HEADER_SUBSCREEN{idx}:SAPMV56A:1021/"
                        ruta = session.findById(base + "ctxtVTTK-ROUTE").text
                        try: signatura = session.findById(base + "ctxtVTTK-SIGNI").text
                        except:
                            try: signatura = session.findById(base + "txtVTTK-SIGNI").text
                            except: pass
                        if ruta: break
                    except: pass

                # 3. PESO (Desde Entregas)
                peso_val = ""
                try:
                    # Ir a Entregas
                    session.findById("wnd[0]/tbar[1]/btn[7]").press()
                    
                    # Leer del tree (no grid)
                    tree = session.findById("wnd[0]/usr/subPLANNING:SAPLV56I_PLAN_SCREEN:0110/cntlV56I_PLAN_SCREEN_CONTAINER/shellcont/shell/shellcont[1]/shell[1]")
                    
                    # Usar GetItemText con el node key correcto (10 espacios + "1")
                    # y el column ID correcto (C + 8 espacios + 132)
                    try: peso_val = tree.GetItemText("          1", "C        132")
                    except:
                        try: peso_val = tree.GetItemText("          1", "C        131")
                        except: pass
                    
                    # Volver atrás
                    session.findById("wnd[0]/tbar[0]/btn[3]").press()
                except: 
                    try: session.findById("wnd[0]/tbar[0]/btn[3]").press()
                    except: pass

                # Usar f_reg si existe, sino f_ini_carga, sino f_despacho, sino f_inicio
                fecha_dato = f_reg if f_reg else (f_ini_carga if f_ini_carga else (f_despacho if f_despacho else f_inicio))
                
                datos_finales.append({
                    "Fecha_Dato": fecha_dato, "Transporte": tknum, "Ruta": ruta, "Signatura": signatura, "Peso_KG": peso_val,
                    "F_Registro": f_reg, "H_Registro": h_reg,
                    "F_InicioCarga": f_ini_carga, "H_InicioCarga": h_ini_carga,
                    "F_FinCarga": f_fin_carga, "H_FinCarga": h_fin_carga,
                    "F_Despacho": f_despacho, "H_Despacho": h_despacho,
                    "F_InicioTransp": f_ini_transp, "H_InicioTransp": h_ini_transp,
                    "F_FinTransp": f_fin_transp, "H_FinTransp": h_fin_transp,
                    "Fecha_Hora_Extraccion": timestamp
                })
            except: pass

        # --- GUARDAR ---
        print("\n💾 Guardando Excel...")
        if datos_finales:
            df_nuevos = pd.DataFrame(datos_finales)
            df_nuevos['Transporte'] = pd.to_numeric(df_nuevos['Transporte'], errors='coerce')
            
            # Limpiar peso (quitar puntos de miles si es texto, o comas)
            # Esto ayuda a que Excel lo lea como numero
            df_nuevos['Peso_KG'] = df_nuevos['Peso_KG'].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False)

            if os.path.exists(archivo_master):
                try:
                    df_master = pd.read_excel(archivo_master)
                    df_master['Transporte'] = pd.to_numeric(df_master['Transporte'], errors='coerce')
                    df_total = pd.concat([df_master, df_nuevos], ignore_index=True)
                    df_total.drop_duplicates(subset=['Transporte'], keep='last', inplace=True)
                except: df_total = df_nuevos
            else: df_total = df_nuevos

            cols = ["Fecha_Dato", "Transporte", "Ruta", "Signatura", "Peso_KG", 
                    "F_Registro", "H_Registro", 
                    "F_InicioCarga", "H_InicioCarga", 
                    "F_FinCarga", "H_FinCarga", 
                    "F_Despacho", "H_Despacho", 
                    "F_InicioTransp", "H_InicioTransp", 
                    "F_FinTransp", "H_FinTransp", 
                    "Fecha_Hora_Extraccion"]
            df_total = df_total.reindex(columns=cols)

            try:
                df_total.to_excel(archivo_master, index=False)
                print(f"✅ Archivo actualizado: {os.path.basename(archivo_master)}")
                
                if enviar_correo:
                    try:
                        out = win32com.client.Dispatch("Outlook.Application")
                        m = out.CreateItem(0)
                        m.To = "ariel.mella@cialalimentos.cl"
                        m.Subject = f"Reporte Transporte {f_inicio} - {f_fin}"
                        m.HTMLBody = "Adjunto reporte actualizado."
                        m.Attachments.Add(archivo_master)
                        m.Display()
                        print("📧 Correo creado y listo para enviar.")
                    except: print("⚠️ No se pudo crear correo.")
                else:
                    print("📋 Correo omitido (opción desactivada).")
            except: print("❌ ERROR: Cierra el Excel para guardar.")
        
        pythoncom.CoUninitialize()