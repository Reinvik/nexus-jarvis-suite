import os
import sys
import shutil
import time
import requests
import subprocess
import json
from tkinter import messagebox

# Configuración
SETTINGS_FILE = "settings.json"
CURRENT_VERSION_FILE = "version.txt"
DEFAULT_DIST_DIR = os.path.join(os.path.expanduser("~"), r"OneDrive - CIAL Alimentos\Nexus_System\Dist")

class NexusUpdater:
    def __init__(self, current_version="1.0.0"):
        self.current_version = current_version
        self.dist_dir = DEFAULT_DIST_DIR
        self.load_config()

    def load_config(self):
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, "r") as f:
                    settings = json.load(f)
                    if "OneDrivePath" in settings:
                        self.dist_dir = os.path.join(settings["OneDrivePath"], "Nexus_System", "Dist")
            except: pass

    def check_for_updates(self):
        """Compara versión local con remota"""
        print(f"[UPDATER] Buscando actualizaciones en: {self.dist_dir}")
        
        remote_version_file = os.path.join(self.dist_dir, "version.txt")
        remote_exe = os.path.join(self.dist_dir, "Nexus Jarvis.exe")
        
        if not os.path.exists(remote_version_file):
            print("[UPDATER] No se encontró información de versión remota.")
            return False

        try:
            with open(remote_version_file, "r") as f:
                remote_ver = f.read().strip()
            
            print(f"[UPDATER] Versión Local: {self.current_version} | Remota: {remote_ver}")
            
            if self.is_newer(remote_ver, self.current_version):
                print("[UPDATER] ¡Nueva versión disponible!")
                if os.path.exists(remote_exe):
                    return True, remote_exe, remote_ver
                else:
                    print("[UPDATER] Aviso: Hay nueva versión pero falta el .exe en el servidor.")
        except Exception as e:
            print(f"[UPDATER] Error verificando actualización: {e}")
            
        return False, None, None

    def is_newer(self, remote, local):
        try:
            r_parts = [int(x) for x in remote.split('.')]
            l_parts = [int(x) for x in local.split('.')]
            return r_parts > l_parts
        except:
            return False

    def perform_update(self, remote_exe_path):
        """
        Descarga (copia) el nuevo EXE y ejecuta script de reemplazo.
        """
        # 1. Notificar usuario
        resp = messagebox.askyesno("Actualización Disponible", 
                                   "Existe una nueva versión de Nexus Jarvis.\n¿Desea actualizar y reiniciar?")
        if not resp:
            return

        # 2. Preparar paths
        current_exe = sys.executable
        temp_exe = "Nexus Jarvis_new.exe"
        
        print("[UPDATER] Iniciando proceso de actualización...")
        
        # 3. Copiar nuevo EXE a temp
        try:
            shutil.copy2(remote_exe_path, temp_exe)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo descargar la actualización: {e}")
            return

        # 4. Crear script BAT para el reemplazo (Self-destructing)
        bat_script = """
@echo off
timeout /t 2 /nobreak
echo Actualizando Nexus Jarvis...
del "{current}"
move "{new}" "{current}"
start "" "{current}"
del "%~f0"
        """.format(current=current_exe, new=os.path.abspath(temp_exe))

        with open("update_runner.bat", "w") as f:
            f.write(bat_script)

        # 5. Ejecutar BAT y cerrar
        print("[UPDATER] Reiniciando para aplicar cambios...")
        subprocess.Popen("update_runner.bat", shell=True)
        sys.exit(0)

def run_updater_check():
    # Intentar leer versión local
    local_ver = "1.0.0"
    if os.path.exists("version.txt"):
        with open("version.txt", "r") as f:
            local_ver = f.read().strip()
            
    updater = NexusUpdater(current_version=local_ver)
    is_avail, exe_path, ver = updater.check_for_updates()
    
    if is_avail:
        updater.perform_update(exe_path)

if __name__ == "__main__":
    run_updater_check()
