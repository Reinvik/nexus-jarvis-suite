import os
import shutil
import subprocess
import sys

# Definir directorios base
BUILD_DIR = "Nexus_Jarvis_Build_v5"
SRC_DIR = os.path.join(BUILD_DIR, "Nexus Jarvis")

USER_DIR = os.path.join(BUILD_DIR, "Entregable_Usuarios")
ADMIN_DIR = os.path.join(BUILD_DIR, "Setup_Completo_Ariel")

# Clean previous
print(f"[INFO] Limpiando directorios previos en {BUILD_DIR}...")
# Usar ignore_errors=True para evitar bloqueos por OneDrive/Antivirus
if os.path.exists(USER_DIR): shutil.rmtree(USER_DIR, ignore_errors=True)
if os.path.exists(ADMIN_DIR): shutil.rmtree(ADMIN_DIR, ignore_errors=True)

# Forzar creación (si rmtree falló parcialmente, esto asegura que existan)
os.makedirs(USER_DIR, exist_ok=True)
os.makedirs(ADMIN_DIR, exist_ok=True)

print(f"[INFO] Empaquetando distribucion desde {SRC_DIR}...")

# --- 1. SETUP ARIEL (Todo) ---
# Copiar todo el contenido de Nexus Jarvis a Setup_Completo_Ariel
print("   [Admin] Copiando distribución completa...")
shutil.copytree(SRC_DIR, os.path.join(ADMIN_DIR, "Nexus Jarvis"))

# --- 2. ENTREGABLE USUARIOS (Solo EXE y Configs) ---
files_for_users = [
    "Nexus Jarvis.exe",
    "README.md",
    "version.txt"
]

# Crear carpeta interna para la app del usuario
user_app_dir = os.path.join(USER_DIR, "Nexus Jarvis App")
os.makedirs(user_app_dir, exist_ok=True)

# Copiar archivos esenciales
for f in files_for_users:
    src = os.path.join(SRC_DIR, f)
    dst = os.path.join(user_app_dir, f)
    if os.path.exists(src):
        shutil.copy2(src, dst)
        print(f"   [Usuario] Copiado {f}")
    else:
        print(f"   [Usuario] ALERTA: {f} no encontrado")

# Crear settings.json limpio para el usuario
print("   [Usuario] Creando settings.json limpio...")
with open(os.path.join(user_app_dir, "settings.json"), "w") as f:
    f.write("{}")

# Copiar carpeta _internal (Necesaria para que corra el EXE en modo onedir)
if os.path.exists(os.path.join(SRC_DIR, "_internal")):
    print("   [Usuario] Copiando dependencias (_internal)...")
    shutil.copytree(os.path.join(SRC_DIR, "_internal"), os.path.join(user_app_dir, "_internal"))
else:
    print("   [ERROR] No se encontró carpeta _internal. ¿Ejecutaste el build primero?")

# --- 3. GENERAR PLANTILLAS (Para Usuario) ---
print("   [Usuario] Generando plantillas Excel...")
# create_templates.py usa: "Nexus_Jarvis_Build_v5/Entregable_Usuarios/Plantillas"
# Esto coincide con USER_DIR/Plantillas
try:
    subprocess.run([sys.executable, "create_templates.py"], check=True)
except Exception as e:
    print(f"   [ERROR] Falló la creación de plantillas: {e}")

# --- 4. DOCUMENTACIÓN FINAL ---
# Crear archivo LEEME para usuarios
readme_path = os.path.join(USER_DIR, "LEEME_PRIMERO.txt")
with open(readme_path, "w", encoding="utf-8") as f:
    f.write("=== NEXUS JARVIS - SUITE DE AUTOMATIZACIÓN ===\n\n")
    f.write("CONTENIDO:\n")
    f.write("1. Carpeta 'Nexus Jarvis App': Contiene la aplicación principal.\n")
    f.write("   -> Ejecute 'Nexus Jarvis.exe' para iniciar.\n")
    f.write("2. Carpeta 'Plantillas': Contiene los Excel necesarios para cada Bot.\n\n")
    f.write("INSTRUCCIONES:\n")
    f.write("- Puede mover la carpeta 'Nexus Jarvis App' a donde desee, pero mantenga la carpeta '_internal' junto al .exe.\n")
    f.write("- Use las plantillas provistas para cargar datos en los bots.\n")
    f.write("- No requiere instalar Python ni nada adicional.\n")

print("\n✅ Distribuición finalizada.")
print(f"   -> {USER_DIR}")
print(f"   -> {ADMIN_DIR}")
