import firebase_admin
from firebase_admin import credentials, firestore
import time
import os
import sys
import requests
import tempfile
import io
from datetime import datetime

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
CREDENTIALS_FILE = "fire.json" 
PC_NAME = "SANJORGE1"

db = None

def init_firebase():
    global db
    if not firebase_admin._apps:
        # Ajuste para PyInstaller y rutas
        if getattr(sys, 'frozen', False):
            # Prioridad 1: Junto al ejecutable (carpeta dist/App)
            base_path_exe = os.path.dirname(sys.executable)
            cred_path = os.path.join(base_path_exe, CREDENTIALS_FILE)
            
            # Prioridad 2: MEIPASS (si se empaquetó dentro)
            if not os.path.exists(cred_path):
                base_path_meipass = sys._MEIPASS
                cred_path_meipass = os.path.join(base_path_meipass, CREDENTIALS_FILE)
                if os.path.exists(cred_path_meipass):
                    cred_path = cred_path_meipass
        else:
            # Modo desarrollo
            cred_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), CREDENTIALS_FILE)
            
        print(f"[DEBUG] Buscando credenciales en: {cred_path}")
        
        if not os.path.exists(cred_path):
             # Ultimo intento: Path relativo directo (confiando en CWD)
             print(f"[WARN] No encontrado en ruta absoluta. Intentando relativo: {CREDENTIALS_FILE}")
             cred_path = CREDENTIALS_FILE 

        cred = credentials.Certificate(cred_path)
        firebase_admin.initialize_app(cred)
    
    db = firestore.client()

def start_worker():
    init_firebase()
    print(f"🤖 WORKER SAP INICIADO EN {PC_NAME}")
    print("📡 Escuchando órdenes desde la Web...")
    procesar_ordenes()

def procesar_ordenes():
    # Primero procesar órdenes pendientes existentes
    print("🔍 Buscando órdenes pendientes existentes...")
    ordenes_pendientes = db.collection('ordenes_bot').where('status', '==', 'pending').get()
    
    for doc in ordenes_pendientes:
        datos = doc.to_dict()
        print(f"\n📩 ORDEN PENDIENTE ENCONTRADA: {datos.get('tipo_bot')}")
        ejecutar_tarea(doc.id, datos)
    
    if len(ordenes_pendientes) == 0:
        print("   No hay órdenes pendientes")
    
    # Luego escuchar nuevas órdenes
    coleccion = db.collection('ordenes_bot').where('status', '==', 'pending')
    
    def on_snapshot(col_snapshot, changes, read_time):
        for change in changes:
            if change.type.name == 'ADDED' or change.type.name == 'MODIFIED':
                doc = change.document
                datos = doc.to_dict()
                # Solo procesar si realmente está pending (evitar duplicados)
                if datos.get('status') == 'pending' and datos.get('worker') != PC_NAME:
                    print(f"\n📩 NUEVA ORDEN RECIBIDA: {datos.get('tipo_bot')}")
                    ejecutar_tarea(doc.id, datos)

    coleccion.on_snapshot(on_snapshot)
    
    while True:
        time.sleep(1)

# --- LOGGER CLOUD ---
class FirestoreLogger:
    def __init__(self, doc_ref):
        self.doc_ref = doc_ref
        self.terminal = sys.stdout

    def write(self, message):
        # 1. Escribir en la consola local (pantalla negra)
        self.terminal.write(message)
        self.terminal.flush()
        
        # 2. Enviar a Firebase si hay texto real
        text = message.strip()
        if text:
            try:
                self.doc_ref.update({
                    'execution_logs': firestore.ArrayUnion([text])
                })
            except Exception as e:
                # Si falla el log, no detener el bot, solo avisar en local
                sys.stderr.write(f"⚠️ Error enviando log a web: {e}\n")

    def flush(self):
        self.terminal.flush()

def ejecutar_tarea(doc_id, datos):
    doc_ref = db.collection('ordenes_bot').document(doc_id)
    
    # 1. Avisar que empezamos
    doc_ref.update({
        'status': 'running',
        'worker': PC_NAME,
        'inicio': firestore.SERVER_TIMESTAMP
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
    sys.stdout = FirestoreLogger(doc_ref)

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
                doc_ref.update({
                    'status': 'success',
                    'worker': PC_NAME,
                    'fecha_termino': firestore.SERVER_TIMESTAMP,
                    'execution_logs': ["✅ Sistema reiniciando..."]
                })
                
                # Dar un momento para que Firestore sincronice
                time.sleep(2)
                sys.exit(0) # Matar este proceso
                
            except Exception as e:
                print(f"❌ Error lanzando reinicio: {e}")
                raise e
        
        else:
            raise Exception(f"Tipo de bot desconocido: {bot_type}")

        print("✅ Tarea finalizada con éxito.")
        sys.stdout = original_stdout
        
        doc_ref.update({
            'status': 'success',
            'fin': firestore.SERVER_TIMESTAMP,
            'mensaje': 'Ejecución completada en SAP.',
            'result_payload': execution_result
        })

    except Exception as e:
        sys.stdout = original_stdout
        print(f"❌ Error ejecutando bot: {e}")
        doc_ref.update({
            'status': 'error',
            'error': str(e)
        })

if __name__ == "__main__":
    # Solo procesar órdenes desde la interfaz web (botones)
    # Los bots de Email y Zonales se activan manualmente como los demás
    start_worker()