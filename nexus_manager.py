import subprocess
import threading
import time
import sys
import os
import signal
from datetime import datetime

# Configuraciones
SERVICES = [
    {
        "name": "WORKER SAP",
        "script": "worker_sap.py",
        "color": "\033[94m", # Azul
        "enabled": True
    },
    {
        "name": "EMAIL CMDR",
        "script": "email_commander.py",
        "color": "\033[92m", # Verde
        "enabled": True
    },
    {
        "name": "ZONALES",
        "script": "worker_zonales.py",
        "color": "\033[93m", # Amarillo
        "enabled": True
    },
    {
        "name": "NEXUS API",
        "script": "nexus_server.py",
        "color": "\033[95m", # Magenta
        "enabled": True
    }
]

RESET = "\033[0m"
processes = {}
running = True

if sys.platform == 'win32':
    os.system('color') # Habilitar colores ANSI en Windows CMD

def log(service_name, message, color):
    timestamp = datetime.now().strftime("%H:%M:%S")
    # Limpiamos saltos de línea extra para que se vea compacto
    msg_clean = message.strip()
    if msg_clean:
        print(f"{color}[{timestamp}] [{service_name}] {msg_clean}{RESET}")

def stream_reader(process, service_info):
    """Lee stdout del proceso y lo imprime"""
    name = service_info['name']
    color = service_info['color']
    
    for line in iter(process.stdout.readline, ''):
        if not running: break
        if line:
            log(name, line, color)
    process.stdout.close()

def run_service(service_info):
    name = service_info['name']
    script = service_info['script']
    color = service_info['color']
    
    while running:
        log(name, "[START] Iniciando servicio...", color)
        
        # Ejecutar con python -u (unbuffered) para ver logs en tiempo real
        cmd = []
        cwd = os.path.dirname(os.path.abspath(__file__))
        
        if getattr(sys, 'frozen', False):
            # Modo Frozen (EXE)
            exe_path = sys.executable
            # Mapear script a argumento
            if script == "worker_sap.py":
                cmd = [exe_path, "--worker-sap"]
            elif script == "email_commander.py":
                cmd = [exe_path, "--email-commander"]
            elif script == "worker_zonales.py":
                cmd = [exe_path, "--worker-zonales"]
            elif script == "nexus_server.py":
                cmd = [exe_path, "--api-server"]
            else:
                log(name, f"[WARN] Script desconocido en modo EXE: {script}", color)
                time.sleep(5)
                continue
        else:
            # Modo Script Normal
            python_exe = sys.executable if sys.platform == "win32" else "python3"
            cmd = [python_exe, "-u", script]

        try:
            p = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                cwd=cwd,
                encoding='utf-8',
                errors='replace'
            )
            processes[name] = p
            
            # Hilo para leer logs
            t = threading.Thread(target=stream_reader, args=(p, service_info))
            t.daemon = True
            t.start()
            
            p.wait() # Esperar a que termine/crashee
            
            if running:
                log(name, f"[WARN] Servicio detenido (Codigo {p.returncode}). Reiniciando en 5s...", "\033[91m")
                time.sleep(5)
                
        except Exception as e:
            log(name, f"[ERROR] Error critico lanzando proceso: {e}", "\033[91m")
            time.sleep(10)

def main():
    global running
    print(f"{RESET}==================================================")
    print(f"   [N.JARVIS] NEXUS JARVIS - PROCESS MANAGER v1.0")
    print(f"==================================================\n")
    
    threads = []
    
    # Iniciar hilos de gestión
    for service_info in SERVICES:
        if service_info['enabled']:
            t = threading.Thread(target=run_service, args=(service_info,))
            t.daemon = True
            t.start()
            threads.append(t)
            time.sleep(1) # Escalonar inicios

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print(f"\n[STOP] Deteniendo NEXUS MANAGER...")
        running = False
        
        # Matar subprocesos
        for name, p in processes.items():
            try:
                print(f"   Matando {name}...")
                p.terminate()
            except:
                pass
        sys.exit(0)

if __name__ == "__main__":
    main()
