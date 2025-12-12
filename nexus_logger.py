import os
import csv
import pandas as pd
from datetime import datetime
import json
import logging

# Configuración por defecto
DEFAULT_LOG_DIR = os.path.join(os.path.expanduser("~"), r"OneDrive - CIAL Alimentos\Nexus_System\Logs")
DEFAULT_LOG_FILE = "bitacora_operaciones.csv"
SETTINGS_FILE = "settings.json"

class NexusLogger:
    def __init__(self):
        self.log_dir = DEFAULT_LOG_DIR
        self.log_file = DEFAULT_LOG_FILE
        self.load_config()
        self.setup_logging()

    def load_config(self):
        """Carga configuración desde settings.json si existe"""
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, "r") as f:
                    settings = json.load(f)
                    # Buscar ruta configurada
                    if "OneDrivePath" in settings:
                        self.log_dir = os.path.join(settings["OneDrivePath"], "Nexus_System", "Logs")
            except:
                pass # Usar defaults

    def setup_logging(self):
        """Asegura que el directorio y archivo existan"""
        if not os.path.exists(self.log_dir):
            try:
                os.makedirs(self.log_dir)
            except Exception as e:
                # Fallback a local si falla OneDrive
                print(f"[LOGGER] Error creando dir en OneDrive: {e}. Usando local.")
                self.log_dir = os.path.join(os.getcwd(), "Logs_Local")
                os.makedirs(self.log_dir, exist_ok=True)

        self.full_path = os.path.join(self.log_dir, self.log_file)
        
        # Crear header si no existe
        if not os.path.exists(self.full_path):
            try:
                df = pd.DataFrame(columns=["Timestamp", "Bot", "Usuario", "Accion", "Detalle", "Estado"])
                df.to_csv(self.full_path, index=False)
            except Exception as e:
                print(f"[LOGGER] Error inicializando CSV: {e}")

    def log(self, bot_name, accion, detalle="", estado="INFO"):
        """Registra un evento en el log centralizado"""
        try:
            timestamp = datetime.now().isoformat()
            usuario = os.getenv('USERNAME') or "Unknown"
            
            # 1. Escribir a CSV (Append mode)
            # Usamos CSV estándar para velocidad y robustez concurrente simple
            with open(self.full_path, "a", newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow([timestamp, bot_name, usuario, accion, detalle, estado])
                
            # 2. También mostrar en consola para debug
            print(f"[{bot_name}] {accion}: {detalle} ({estado})")
            
        except Exception as e:
            print(f"[LOGGER ERROR] No se pudo escribir log: {e}")

# Instancia global
nexus_logger = NexusLogger()

def log_event(bot_name, accion, detalle="", estado="INFO"):
    nexus_logger.log(bot_name, accion, detalle, estado)

if __name__ == "__main__":
    # Prueba
    log_event("TestBot", "Inicio", "Probando logger centralizado", "OK")
    print(f"Log escrito en: {nexus_logger.full_path}")
