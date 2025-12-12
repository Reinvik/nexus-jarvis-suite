import win32com.client
import sys

# Configuración UTF-8 para consola
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

def check_current_folder():
    print("🚀 Verificando carpeta seleccionada en Outlook...")
    
    try:
        try:
            outlook = win32com.client.GetObject(Class="Outlook.Application")
            print("✅ Conectado a instancia existente de Outlook.")
        except:
            outlook = win32com.client.Dispatch("Outlook.Application")
            print("⚠️ Iniciando nueva instancia de Outlook (o Dispatch).")
            
        explorer = outlook.ActiveExplorer()
        
        if not explorer:
            print("⚠️ ActiveExplorer es None. Intentando obtener el primer Explorer disponible...")
            if outlook.Explorers.Count > 0:
                explorer = outlook.Explorers.Item(1)
            else:
                print("❌ No se encontraron ventanas de explorador de Outlook abiertas.")
                return
            
        folder = explorer.CurrentFolder
        
        print("\n📂 CARPETA SELECCIONADA ACTUALMENTE:")
        print("=" * 50)
        print(f"Nombre: {folder.Name}")
        print(f"Ruta completa: {folder.FolderPath}")
        print(f"Total Elementos: {folder.Items.Count}")
        print(f"No leídos: {folder.UnReadItemCount}")
        print(f"EntryID: {folder.EntryID}")
        print(f"StoreID: {folder.StoreID}")
        print("=" * 50)
        
        if folder.Items.Count > 0:
            print("\nℹ️ Primeros 3 elementos en esta carpeta:")
            items = folder.Items
            items.Sort("[ReceivedTime]", True)
            for i in range(1, min(4, folder.Items.Count + 1)):
                try:
                    item = items.Item(i)
                    print(f"   - {item.Subject} (Recibido: {item.ReceivedTime})")
                except Exception as e:
                    print(f"   - Error leyendo item: {e}")

    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    check_current_folder()
