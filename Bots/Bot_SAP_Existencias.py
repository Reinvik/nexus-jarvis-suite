import os
import subprocess
import time

class SapBotExistencias:
    def run(self, mode="normal"):
        """
        Ejecuta el bot de SAP (VBScript).
        mode: "normal" o "simulation"
        """
        script_name = "sap_bot.vbs" if mode == "normal" else "sap_bot_simulation.vbs"
        
        # Obtener ruta absoluta del script VBS
        base_path = os.path.dirname(os.path.abspath(__file__))
        script_path = os.path.join(base_path, script_name)
        
        if not os.path.exists(script_path):
            raise FileNotFoundError(f"No se encontró el script: {script_path}")
            
        print(f"[BOT] Ejecutando {script_name}...")
        
        # Ejecutar VBScript usando cscript o wscript
        # Usamos cscript para capturar output si es necesario, o simplemente os.startfile
        # os.startfile es mejor para VBS que interactúan con GUI
        
        try:
            os.startfile(script_path)
            print(f"[BOT] {script_name} lanzado correctamente.")
        except Exception as e:
            raise RuntimeError(f"Error al lanzar el script: {e}")

if __name__ == "__main__":
    bot = SapBotExistencias()
    bot.run("simulation")
