import time
import sys
import os
from datetime import datetime

# Asegurar que podemos importar desde Bots
sys.path.append(os.path.join(os.path.dirname(__file__), 'Bots'))

try:
    from Bots.Bot_Consolidacion_Zonales import BotConsolidacionZonales
except ImportError:
    # Fallback si se ejecuta desde dentro de Apps/Nexus_Jarvis
    sys.path.append(os.path.join(os.path.dirname(__file__)))
    from Bots.Bot_Consolidacion_Zonales import BotConsolidacionZonales

if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')

def main():
    print("🌍 INICIANDO WORKER ZONALES (Intervalo: 60 min)")
    bot = BotConsolidacionZonales()
    
    while True:
        try:
            print(f"\n⏰ Ejecutando ciclo Zonales: {datetime.now().strftime('%H:%M:%S')}")
            bot.run() # Esto ejecuta un ciclo de escaneo y consolidación
            
            print("💤 Durmiendo 60 minutos...")
            time.sleep(3600) 
            
        except KeyboardInterrupt:
            print("🛑 Detenido por usuario.")
            break
        except Exception as e:
            print(f"❌ Error en ciclo worker zonales: {e}")
            time.sleep(60) # Esperar 1 min tras error antes de reintentar

if __name__ == "__main__":
    main()
