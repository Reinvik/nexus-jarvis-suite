import time
import os
import sys
import requests
import tempfile
import io
import json
from datetime import datetime
from dotenv import load_dotenv



# --- CONFIGURACIÓN UTF-8 PARA WINDOWS ---
# Esto evita errores de codificación con emojis en la consola de Windows
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

# --- IMPORTA TUS BOTS EXISTENTES ---
sys.path.append(os.path.join(os.path.dirname(__file__), 'Bots'))
sys.path.append(os.path.join(os.path.dirname(__file__), 'Tools'))

try:
    from Tx_MIGO3 import SapMigoBotTurbo
    from Bot_Pallet import SapBotPallet
    from Bot_Transporte import SapBotTransporte
    from Bot_Auditor import SapBotAuditor
    from Bot_Traspaso_LT01 import SapBotTraspasoLT01
    from Bot_Conversiones_UMV import SapBotConversiones
    from Bot_Conciliacion_Email import SapBotConciliacionEmail
    from Bot_Consolidacion_Zonales import BotConsolidacionZonales
    from Bot_Analisis_Zonales import BotAnalisisZonales
    from Bot_Vision import BotVisionPizarra
except ImportError as e:
    print(f"❌ Error importando bots: {e}")
    sys.exit(1)

# --- CONFIGURACIÓN ---
load_dotenv()
SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")
PC_NAME = "SANJORGE1"

# Headers globales para Supabase
HEADERS = {
    "apikey": SUPABASE_KEY,
    "Authorization": f"Bearer {SUPABASE_KEY}",
    "Content-Type": "application/json",
    "Prefer": "return=minimal"
}

def init_supabase():
    if not SUPABASE_URL or not SUPABASE_KEY:
        print("❌ Error: Faltan variables de entorno SUPABASE_URL o SUPABASE_KEY")
        sys.exit(1)
    # Ya no necesitamos crear un cliente pesado, usaremos requests



def start_worker():
    init_supabase()
    print(f"🤖 WORKER SAP INICIADO EN {PC_NAME}")
    print("📡 Escuchando órdenes desde Supabase (NexusStaging)...")
    procesar_ordenes()

def procesar_ordenes():
    print("🔍 Buscando órdenes pendientes...")
    url = f"{SUPABASE_URL}/rest/v1/ordenes_bot?status=eq.pending"
    
    while True:
        try:
            # Consultar órdenes pendientes vía REST
            response = requests.get(url, headers=HEADERS)
            if response.status_code == 200:
                ordenes = response.json()
                for datos in ordenes:
                    if datos.get('worker') != PC_NAME:
                        print(f"\n📩 NUEVA ORDEN RECIBIDA: {datos.get('tipo_bot')}")
                        ejecutar_tarea(datos.get('id'), datos)
            else:
                print(f"⚠️ Error Supabase: {response.status_code} - {response.text}")
        except Exception as e:
            print(f"⚠️ Error consultando órdenes: {e}")
        
        time.sleep(3) # Polling cada 3 segundos

# --- LOGGER SUPABASE ---
class SupabaseLogger:
    def __init__(self, doc_id):
        self.doc_id = doc_id
        self.terminal = sys.stdout
        self.url = f"{SUPABASE_URL}/rest/v1/rpc/append_execution_log"

    def write(self, message):
        self.terminal.write(message)
        self.terminal.flush()
        text = message.strip()
        if text:
            try:
                # Llamada al RPC de logs
                requests.post(self.url, headers=HEADERS, json={
                    'order_id': self.doc_id,
                    'log_line': text
                })
            except:
                pass

    def flush(self):
        self.terminal.flush()



def ejecutar_tarea(doc_id, datos):
    # 1. Avisar que empezamos
    url_order = f"{SUPABASE_URL}/rest/v1/ordenes_bot?id=eq.{doc_id}"
    requests.patch(url_order, headers=HEADERS, json={
        'status': 'running',
        'worker': PC_NAME,
        'inicio': datetime.now().isoformat()
    })



    bot_type = datos.get('tipo_bot')
    ruta_archivo = datos.get('ruta_archivo')
    
    # --- DEBUG: Ver qué llega ---
    print(f"🔍 Datos completos de la orden: {datos}")
    print(f"🔍 Parámetros extraídos: {datos.get('parametros')}")
    # ----------------------------

    # DESCARGAR ARCHIVO SI ES URL
    archivo_local = None
    if ruta_archivo and ruta_archivo.startswith("http"):
        try:
            print(f"⬇️ Descargando archivo desde: {ruta_archivo[:50]}...")
            response = requests.get(ruta_archivo)
            if response.status_code == 200:
                # Crear archivo temporal manteniendo la extensión original si es posible
                nombre_original = datos.get('nombre_archivo_original', 'archivo_temp.xlsx')
                ext = os.path.splitext(nombre_original)[1]
                if not ext: ext = ".xlsx"
                
                temp_dir = tempfile.gettempdir()
                archivo_local = os.path.join(temp_dir, f"temp_bot_{int(time.time())}{ext}")
                
                with open(archivo_local, 'wb') as f:
                    f.write(response.content)
                
                print(f"✅ Archivo descargado en: {archivo_local}")
                ruta_archivo = archivo_local
            else:
                print(f"⚠️ Error descargando archivo: Status {response.status_code}")
        except Exception as e:
            print(f"❌ Error descargando archivo: {e}")
    elif not ruta_archivo and datos.get('nombre_archivo_original'):
        # MODO ARCHIVO ABIERTO / LOCAL
        ruta_archivo = datos.get('nombre_archivo_original')
        print(f"📂 Modo Local/Abierto: Usando nombre '{ruta_archivo}'")

    # CAPTURAR LOGS
    original_stdout = sys.stdout
    sys.stdout = SupabaseLogger(doc_id)


    try:
        execution_result = None

        # --- ENRUTADOR DE BOTS ---
        if bot_type == 'MIGO':
            bot = SapMigoBotTurbo()
            bot.run(ruta_archivo)
            
        elif bot_type == 'PALLET':
            bot = SapBotPallet()
            bot.run(ruta_archivo)
            
        elif bot_type == 'TRANSPORTE':
            bot = SapBotTransporte()
            fechas = datos.get('parametros', {}).get('fechas')
            enviar_correo = datos.get('parametros', {}).get('sendEmail', False)
            print(f"🚚 Ejecutando Bot Transporte con fechas: {fechas}, enviar_correo: {enviar_correo}")
            bot.run(fechas, enviar_correo)
            
        elif bot_type == 'AUDITOR':
            bot = SapBotAuditor()
            almacen = datos.get('parametros', {}).get('almacen', 'SGVT')
            execution_result = bot.run(almacen)
            
        elif bot_type == 'LT01':
            bot = SapBotTraspasoLT01()
            bot.run(ruta_archivo)
            
        elif bot_type == 'UMV':
            bot = SapBotConversiones()
            bot.run(ruta_archivo)
            
        elif bot_type == 'CONCILIACION_EMAIL':
            bot = SapBotConciliacionEmail()
            bot.run()
            
        elif bot_type == 'ZONALES':
            bot = BotConsolidacionZonales()
            bot.run()

        elif bot_type == 'ANALISIS_ZONALES':
            bot = BotAnalisisZonales()
            bot.run()
            
        elif bot_type == 'VISION':
            bot = BotVisionPizarra()
            bot.run(ruta_archivo)
            
        elif bot_type == 'SYSTEM_RESTART':
            print("🔄 REINICIO SOLICITADO POR USUARIO")
            print("   Lanzando reiniciar.bat...")
            
            # Lanzar reiniciar.bat en una nueva consola independiente
            import subprocess
            try:
                subprocess.Popen(
                    ['cmd', '/c', 'start', 'reiniciar.bat'],
                    cwd=os.getcwd(),
                    creationflags=subprocess.CREATE_NEW_CONSOLE
                )
                print("   ✅ Script de reinicio lanzado. Cerrando worker...")
                
                # Marcar orden como completada antes de morir
                requests.patch(url_order, headers=HEADERS, json={
                    'status': 'success',
                    'worker': PC_NAME,
                    'fin': datetime.now().isoformat(),
                    'execution_logs': ["✅ Sistema reiniciando..."]
                })
                
                # Dar un momento para que Supabase sincronice
                time.sleep(2)
                sys.exit(0) # Matar este proceso


                
            except Exception as e:
                print(f"❌ Error lanzando reinicio: {e}")
                raise e
        
        else:
            raise Exception(f"Tipo de bot desconocido: {bot_type}")

        print("✅ Tarea finalizada con éxito.")
        sys.stdout = original_stdout
        
        requests.patch(url_order, headers=HEADERS, json={
            'status': 'success',
            'fin': datetime.now().isoformat(),
            'mensaje': 'Ejecución completada en SAP.',
            'result_payload': execution_result
        })

    except Exception as e:
        sys.stdout = original_stdout
        print(f"❌ Error ejecutando bot: {e}")
        requests.patch(url_order, headers=HEADERS, json={
            'status': 'error',
            'error': str(e)
        })



if __name__ == "__main__":
    # Solo procesar órdenes desde la interfaz web (botones)
    # Los bots de Email y Zonales se activan manualmente como los demás
    start_worker()