import win32com.client
import sys

# Configuración UTF-8 para consola
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

def diagnostico_outlook():
    print("🚀 Iniciando Diagnóstico de Estructura de Outlook...")
    
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        print("\n📂 LISTADO DE CUENTAS Y CARPETAS:")
        print("=" * 50)
        
        for folder in namespace.Folders:
            print(f"📧 CUENTA: {folder.Name}")
            try:
                listar_subcarpetas(folder, 1)
            except Exception as e:
                print(f"   ❌ Error leyendo cuenta: {e}")
            print("-" * 50)
            
    except Exception as e:
        print(f"❌ Error crítico conectando a Outlook: {e}")

def listar_subcarpetas(parent_folder, level):
    indent = "   " * level
    for folder in parent_folder.Folders:
        try:
            count = folder.Items.Count
            unread = folder.UnReadItemCount
            
            # Marcar visualmente si es la carpeta que buscamos
            marker = ""
            if "zonal" in folder.Name.lower():
                marker = " 🎯 <--- POSIBLE CANDIDATO"
            
            print(f"{indent}📁 {folder.Name} [Total: {count}, No leídos: {unread}]{marker}")
            
            # Si encontramos una carpeta Zonales con items, listar los primeros 3 para confirmar
            if "zonal" in folder.Name.lower() and count > 0:
                print(f"{indent}   ℹ️ Primeros 3 asuntos en '{folder.Name}':")
                items = folder.Items
                items.Sort("[ReceivedTime]", True)
                for i in range(1, min(4, count + 1)):
                    try:
                        print(f"{indent}      - {items[i].Subject}")
                    except:
                        pass

            # Recursividad limitada para no saturar
            if level < 3: 
                listar_subcarpetas(folder, level + 1)
                
        except Exception as e:
            print(f"{indent}❌ Acceso denegado a {folder.Name}")

if __name__ == "__main__":
    diagnostico_outlook()
