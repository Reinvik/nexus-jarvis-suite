import customtkinter as ctk
import threading
import sys
import os
import json
from tkinter import filedialog, messagebox

SETTINGS_FILE = "settings.json"

# --- LOGGER CENTRALIZADO ---
try:
    from nexus_logger import log_event
    log_event("LauncherGUI", "Inicio", "Abriendo Panel de Control", "OK")
except Exception as e:
    print(f"Error cargando logger: {e}")
    def log_event(*args): pass

# --- CONFIGURACIÓN PATH ---
sys.path.append(os.path.join(os.path.dirname(__file__), 'Bots'))

# --- IMPORTAR TODOS LOS BOTS (Con reporte de errores real) ---
print("--- CARGANDO SISTEMA ---")
try:
    from Tx_MIGO3 import SapMigoBotTurbo 
    print("✅ MIGO cargado.")
except Exception as e: 
    print(f"❌ Error MIGO: {e}")
    SapMigoBotTurbo = None

try:
    from Bot_Pallet import SapBotPallet
    print("✅ Pallet cargado.")
except Exception as e: 
    print(f"❌ Error Pallet: {e}")
    SapBotPallet = None

try:
    from Bot_Transporte import SapBotTransporte
    print("✅ Transporte cargado.")
except Exception as e: 
    print(f"❌ Error Transporte: {e}")
    SapBotTransporte = None

try:
    from Bot_Vision import BotVisionPizarra
    print("✅ Visión cargado.")
except Exception as e: 
    print(f"❌ Error Visión: {e}")
    BotVisionPizarra = None

try:
    from Bot_Auditor import SapBotAuditor
    print("✅ Auditor cargado.")
except Exception as e: 
    print(f"❌ Error Auditor: {e}")
    SapBotAuditor = None

try:
    from Bot_Traspaso_LT01 import SapBotTraspasoLT01
    print("✅ Traspaso LT01 cargado.")
except Exception as e:
    print(f"❌ Error Traspaso LT01: {e}")
    SapBotTraspasoLT01 = None

try:
    from Bot_Conversiones_UMV import SapBotConversiones
    print("✅ Conversiones UMV cargado.")
except Exception as e:
    print(f"❌ Error Conversiones UMV: {e}")
    SapBotConversiones = None

try:
    from Bot_Reporte_Cambiados import BotReporteCambiados
    print("✅ Resporte Cambiados cargado.")
except Exception as e:
    print(f"❌ Error Reporte Cambiados: {e}")
    BotReporteCambiados = None

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Panel de Control SAP - CIAL")
        self.geometry("900x600")
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.bot_actual_class = None

        # --- MENU LATERAL ---
        self.sidebar = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar.grid(row=0, column=0, rowspan=4, sticky="nsew")
        
        ctk.CTkLabel(self.sidebar, text="🤖 BOTS SAP", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=20)

        self.crear_boton("Transferencia MIGO", self.panel_migo)
        self.crear_boton("Auditor de Altura", self.panel_pallet)
        self.crear_boton("Auditor de Transporte", self.panel_transporte)
        # self.crear_boton("Vision IA: Panel de operación", self.panel_vision) # REMOVED
        self.crear_boton("Auditor de tránsitos pendientes", self.panel_auditor)
        self.crear_boton("Transferencias Lt01", self.panel_lt01)
        self.crear_boton("Conversiones UMV", self.panel_conversiones)
        
        # --- REPORTES ---
        ctk.CTkLabel(self.sidebar, text="📊 REPORTES", font=ctk.CTkFont(size=16, weight="bold")).pack(pady=10)
        self.crear_boton("Reporte: Cambiados", self.panel_reporte_cambiados)

        # --- PANEL CENTRAL ---
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        
        self.setup_ui_generica()

    def crear_boton(self, texto, comando):
        btn = ctk.CTkButton(self.sidebar, text=texto, command=comando, height=40)
        btn.pack(pady=5, padx=20)

    def setup_ui_generica(self):
        self.lbl_title = ctk.CTkLabel(self.main_frame, text="Selecciona un Bot", font=ctk.CTkFont(size=24))
        self.lbl_title.pack(pady=20)
        
        self.lbl_info = ctk.CTkLabel(self.main_frame, text="", justify="left")
        self.lbl_info.pack(pady=10)
        
        self.entry_file = ctk.CTkEntry(self.main_frame, placeholder_text="Ruta archivo...", width=400)
        self.btn_select = ctk.CTkButton(self.main_frame, text="Buscar Archivo", command=self.sel_archivo, fg_color="gray")
        
        self.btn_run = ctk.CTkButton(self.main_frame, text="EJECUTAR", command=self.run_migo_thread, height=50, fg_color="#2CC985", state="disabled")
        self.btn_run.pack(pady=20)
        
        self.log_box = ctk.CTkTextbox(self.main_frame, width=600, height=250)
        self.log_box.pack(fill="both", expand=True)

    def load_settings(self):
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, "r") as f:
                    return json.load(f)
            except:
                return {}
        return {}

    def save_settings(self, bot_name, path):
        settings = self.load_settings()
        settings[bot_name] = path
        try:
            with open(SETTINGS_FILE, "w") as f:
                json.dump(settings, f)
        except Exception as e:
            print(f"Error guardando settings: {e}")

    def reset_ui(self, titulo, info, necesita_archivo=False):
        self.lbl_title.configure(text=titulo)
        self.lbl_info.configure(text=info)
        self.log_box.delete("0.0", "end")
        self.entry_file.delete(0, "end")
        
        if necesita_archivo:
            self.entry_file.pack(pady=5)
            self.btn_select.pack(pady=5)
            
            # Cargar ruta guardada si existe
            if self.bot_actual_class:
                bot_name = getattr(self.bot_actual_class, "__name__", "")
                settings = self.load_settings()
                saved_path = settings.get(bot_name, "")
                if saved_path and os.path.exists(saved_path):
                    self.entry_file.delete(0, "end")
                    self.entry_file.insert(0, saved_path)
        else:
            self.entry_file.pack_forget()
            self.btn_select.pack_forget()
        
        self.btn_run.configure(state="normal")

    # --- PANELES ---
    def panel_migo(self):
        self.bot_actual_class = SapMigoBotTurbo
        self.reset_ui("Transferencia MIGO", "Selecciona el Excel para realizar la carga (Plantilla_MIGO.xlsx).", True)

    def panel_pallet(self):
        self.bot_actual_class = SapBotPallet
        self.reset_ui("Auditor de Altura", "Selecciona el Excel para pegar la hoja de LX02.", True)

    def panel_transporte(self):
        self.bot_actual_class = SapBotTransporte
        self.reset_ui("Auditor de Transporte", "Extrae reporte VT11 (Rango Fechas) -> VT03N.", False)

    def panel_vision(self):
        self.bot_actual_class = BotVisionPizarra
        self.reset_ui("Vision IA", "Baja temporal.", True)

    def panel_auditor(self):
        self.bot_actual_class = SapBotAuditor
        self.reset_ui("Auditor de tránsitos pendientes", "Al ejecutar, te pedirá el Almacén.", False)

    def panel_lt01(self):
        self.bot_actual_class = SapBotTraspasoLT01
        self.reset_ui("Transferencias Lt01",
                      "Realiza traspasos LT01 basándose en stock.\nSelecciona el Excel (Plantilla_LT01.xlsx).",
                      necesita_archivo=True)

    def panel_conversiones(self):
        self.bot_actual_class = SapBotConversiones
        self.reset_ui("Conversiones UMV",
                      "Extrae factores de conversión (UN -> UNV/CJ).\nSelecciona el Excel con materiales.",
                      necesita_archivo=True)

    def panel_reporte_cambiados(self):
        self.bot_actual_class = BotReporteCambiados
        self.reset_ui("Reporte Cambiados Zonales", 
                      "Genera el borrador del correo con la tabla de 'Cambiados'.\n\nRequiere que el Excel 'Reporte Desv. Zonales...' esté en OneDrive.", 
                      False)

    # --- LOGICA ---
    def sel_archivo(self):
        f = filedialog.askopenfilename()
        if f: 
            self.entry_file.delete(0, "end")
            self.entry_file.insert(0, f)

    def log(self, msg):
        self.log_box.insert("end", str(msg) + "\n")
        self.log_box.see("end")

    # --- FUNCIONES DE EJECUCIÓN (AHORA DENTRO DE LA CLASE) ---
    def run_migo_thread(self):
        if self.bot_actual_class is None:
            self.log("❌ ERROR: El bot seleccionado no se cargó correctamente (Faltan librerías o error de código).")
            return

        ruta = None
        arg_extra = None

        # Validar archivo para MIGO, PALLET, VISION, LT01 y CONVERSIONES
        if self.bot_actual_class in [SapMigoBotTurbo, SapBotPallet, BotVisionPizarra, SapBotTraspasoLT01, SapBotConversiones]:
            ruta = self.entry_file.get()
            if not ruta:
                self.log("❌ Error: Selecciona un archivo primero.")
                return
            
            # Guardar ruta exitosa
            bot_name = getattr(self.bot_actual_class, "__name__", "")
            self.save_settings(bot_name, ruta)

        # Input especial para Auditor
        elif self.bot_actual_class == SapBotAuditor:
            almacenes = ["SGVT", "CDNW", "SGTR", "TAVI", "SGSD", "AVAS", "SGBC", "SGVE", "SGEN", "SDIF"]
            dialog = ctk.CTkInputDialog(text=f"Almacenes: {', '.join(almacenes)}\n\nEscribe el Almacén:", title="Auditor MM")
            arg_extra = dialog.get_input()
            if not arg_extra:
                self.log("Cancelado.")
                return
            arg_extra = arg_extra.upper().strip()

        self.btn_run.configure(state="disabled", text="Ejecutando...")
        
        # Obtener nombre del bot de forma segura
        nombre_bot = getattr(self.bot_actual_class, "__name__", "Bot SAP")
        self.log(f"--- Iniciando {nombre_bot}... ---")
        
        threading.Thread(target=self.execute, args=(ruta, arg_extra)).start()

    def execute(self, ruta, arg_extra):
        try:
            bot = self.bot_actual_class()
            
            # Ejecución según tipo de bot
            if self.bot_actual_class in [SapMigoBotTurbo, SapBotPallet, BotVisionPizarra, SapBotTraspasoLT01, SapBotConversiones]:
                bot.run(ruta) 
            elif self.bot_actual_class == SapBotAuditor:
                bot.run(arg_extra)
            else: # Transporte
                bot.run()
                
            self.log("✅ PROCESO FINALIZADO CON ÉXITO")
        except Exception as e:
            self.log(f"❌ ERROR: {e}")
        finally:
            self.btn_run.configure(state="normal", text="EJECUTAR")

def main():
    app = App()
    app.mainloop()

if __name__ == "__main__":
    main()