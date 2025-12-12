import win32com.client
import sys

# Configuración UTF-8 para consola
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

def find_real_zonales():
    print("🚀 Buscando la verdadera carpeta Zonales...")
    
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        # Recorrer todas las carpetas de todas las cuentas
        for account_folder in namespace.Folders:
            print(f"🔍 Escaneando cuenta: {account_folder.Name}")
            search_recursive(account_folder)
            
    except Exception as e:
        print(f"❌ Error: {e}")

def search_recursive(folder, level=0):
    try:
        # Verificar si esta carpeta es "Zonales"
        if "zonal" in folder.Name.lower():
            print(f"🎯 Encontrado candidato: {folder.FolderPath}")
            print(f"   Items: {folder.Items.Count}")
            
            # Verificar hijos
            children = [f.Name for f in folder.Folders]
            print(f"   Hijos: {children}")
            
            if "Licitaciones" in children or "Perdida vacío" in children:
                print("   ✅ ¡ESTA ES LA CARPETA CORRECTA! (Tiene hijos coincidentes)")
                print(f"   EntryID: {folder.EntryID}")
                
                # Listar items para confirmar
                if folder.Items.Count > 0:
                    print("   📧 Primeros correos:")
                    items = folder.Items
                    items.Sort("[ReceivedTime]", True)
                    for i in range(1, min(4, folder.Items.Count + 1)):
                        print(f"      - {items[i].Subject}")
            else:
                print("   ⚠️ No parece ser la correcta (estructura diferente).")

        # Recursividad
        for sub in folder.Folders:
            search_recursive(sub, level + 1)
            
    except Exception as e:
        pass

if __name__ == "__main__":
    find_real_zonales()
