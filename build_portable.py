import PyInstaller.__main__
import os
import shutil
import sys
import io

# Fix unicode on windows console
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


# Configuración
APP_NAME = "Nexus Jarvis"
MAIN_SCRIPT = "nexus_launcher.py" 
OUTPUT_DIR = "Nexus_Jarvis_Build_v5"

# Asegurar que existe el directorio
if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)

print(f"🔨 Iniciando compilación UNIVERSAL de {APP_NAME}...")

# Argumentos base
args = [
    MAIN_SCRIPT,
    '--name=%s' % APP_NAME,
    '--onedir',             
    # '--windowed',  # IMPORTANTE: Si usamos windowed, no veremos logs de consola de los workers. 
                     # Para "Producción", windowed es mejor, pero workers necesitan ver logs?
                     # Los workers escriben a FirestoreLogger.
                     # Si ejecutamos el EXE doble click -> GUI (Windowed es mejor).
                     # Si ejecutamos nexus_manager -> Console es mejor?
                     
                     # Opción: Compilar como Windowed. 
                     # Cuando se ejecuta por consola (manager), stdout puede no verse.
                     # Pero nexus_manager usa print.
                     # Compromiso: Windowed + alloc_console en código C++?
                     # O simplemente Console mode por ahora para debug? 
                     # El usuario pidió "exe version of interface". Interface = Windowed usually.
                     # Pero "nexus_manager" es background.
                     
                     # Vamos a usar --windowed. Si el manager corre en background, mejor.
    '--windowed',
    
    '--distpath=%s' % OUTPUT_DIR, 
    '--workpath=%s/build' % OUTPUT_DIR,
    '--specpath=%s' % OUTPUT_DIR,
    '--noconfirm',
    '--clean',
    
    # --- IMPORTS OCULTOS ---
    # Scripts en raíz que se importan dinámicamente o por nexus_launcher
    '--hidden-import=worker_sap',
    '--hidden-import=email_commander',
    '--hidden-import=worker_zonales',
    '--hidden-import=nexus_manager',
    '--hidden-import=logistic_suite',
    
    # Librerías criticas
    '--hidden-import=customtkinter',
    '--hidden-import=PIL',
    '--hidden-import=win32com.client',
    '--hidden-import=pythoncom',
    '--hidden-import=pywintypes',
    '--hidden-import=firebase_admin',
    '--hidden-import=google.generativeai',
    
    # --- RECURSOS ---
    '--collect-all=customtkinter', # Trae temas y assets de CTk
    '--collect-submodules=Bots',   # Trae todos los bots
]

# Ejecutar PyInstaller
PyInstaller.__main__.run(args)

print("✅ Compilación exitosa.")

# Copiar archivos de configuración externos Y OTROS RECURSOS
dist_folder = os.path.join(OUTPUT_DIR, APP_NAME)

files_to_copy = [
    "fire.json", 
    "settings.json", 
    "README.md"
]

# Copiar carpeta Tools si existe independiente
tools_src = "Tools"
tools_dest = os.path.join(dist_folder, "Tools")
if os.path.exists(tools_src):
    if os.path.exists(tools_dest):
        shutil.rmtree(tools_dest)
    shutil.copytree(tools_src, tools_dest)
    print("📂 Carpeta Tools copiada.")

print("📂 Copiando archivos de configuración...")
for f in files_to_copy:
    if os.path.exists(f):
        shutil.copy2(f, dist_folder)
        print(f"   -> {f} copiado.")
    else:
        print(f"   ⚠️ {f} no encontrado (se omitió).")

# Crear BAT para iniciar el Manager
manager_bat = os.path.join(dist_folder, "start_manager.bat")
with open(manager_bat, "w") as f:
    f.write('@echo off\n')
    f.write('echo [START] Iniciando Nexus Manager...\n')
    f.write('start "" "Nexus Jarvis.exe" --manager\n')

print(f"   -> Creado start_manager.bat")

print(f"\n🚀 LISTO! La aplicación está en:\n{os.path.abspath(dist_folder)}")
print(f"   Ejecutable principal: {APP_NAME}.exe")
