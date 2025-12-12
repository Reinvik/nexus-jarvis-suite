import sys
import os
import multiprocessing
import time
import io

if sys.platform == 'win32':
    # Fix unicode on windows console
    if sys.stdout:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    if sys.stderr:
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

# Necessary for PyInstaller multiprocessing support
multiprocessing.freeze_support()

# Import core modules lazily but we rely on hidden imports/PyInstaller analysis
# to bundle them.

# Import core modules lazily but we rely on hidden imports/PyInstaller analysis
# to bundle them.
import nexus_updater

def main():
    # --- AUTO-UPDATE CHECK (GUI MODE ONLY) ---
    if len(sys.argv) <= 1:
        # Solo chequear actualizaciones si iniciamos en modo normal (GUI)
        # para evitar bucles si lo iniciamos desde otro proceso
        nexus_updater.run_updater_check()

    # Parse arguments
    mode = "gui"
    if len(sys.argv) > 1:
        # Check argument like --worker-sap
        arg = sys.argv[1]
        
        if arg == "--worker-sap":
            mode = "worker_sap"
        elif arg == "--email-commander":
            mode = "email_commander"
        elif arg == "--worker-zonales":
            mode = "worker_zonales"
        elif arg == "--manager":
            mode = "manager"
    
    # Hide console if in GUI mode and running from console? 
    # No, PyInstaller handles console vs windowed. 
    # If compiled as windowed, stdout goes nowhere unless we redirect it.
    
    # If mode is NOT gui, we might want to attach a console if possible?
    # Or rely on logs. For now, we assume simple execution.
    
    if mode != "gui":
        print(f"[INFO] Nexus Jarvis Launcher: Mode={mode}")

    if mode == "worker_sap":
        import worker_sap
        worker_sap.start_worker()
        
    elif mode == "email_commander":
        import email_commander
        email_commander.main()
        
    elif mode == "worker_zonales":
        import worker_zonales
        worker_zonales.main()
        
    elif mode == "manager":
        import nexus_manager
        nexus_manager.main()
        
    else:
        # GUI
        try:
            import logistic_suite
            logistic_suite.main()
        except ImportError as e:
            # If standard import fails, try adding local dir?
            # Usually strict import is better.
            print(f"[ERROR] Error lanzando interfaz grafica: {e}")
            print("Asegurate de ejecutar desde la carpeta correcta.")
            time.sleep(10)
        except Exception as e:
            print(f"[ERROR] Error inesperado: {e}")
            import traceback
            traceback.print_exc()
            time.sleep(10)

if __name__ == "__main__":
    main()
