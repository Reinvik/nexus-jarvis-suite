import win32com.client
import sys

# Configuración UTF-8 para consola
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

def debug_mailbox():
    print("🚀 Debugging Mailbox Content...")
    
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        # Buscar cuenta
        target_account = "Ariel.Mella@cial.cl"
        root = None
        for f in namespace.Folders:
            if target_account.lower() in f.Name.lower():
                root = f
                break
        
        if not root:
            print("❌ No se encontró la cuenta.")
            return

        print(f"📧 Cuenta: {root.Name}")
        print(f"🔌 Modo Exchange Caché: {root.Store.IsCachedExchange}")
        
        # Buscar Inbox
        inbox = None
        for f in root.Folders:
            if f.Name.lower() in ["bandeja de entrada", "inbox"]:
                inbox = f
                break
        
        if not inbox:
            print("❌ No se encontró Inbox.")
            return

        print(f"📂 Inbox: {inbox.Name} (Total: {inbox.Items.Count})")

        # Check Jefa Irene
        jefa = None
        for f in inbox.Folders:
            if "irene" in f.Name.lower():
                jefa = f
                break
        
        if jefa:
            print(f"📂 Jefa Irene: {jefa.Name} (Total: {jefa.Items.Count})")
        
        # Check Zonales
        zonales = None
        for f in inbox.Folders:
            if "zonal" in f.Name.lower():
                zonales = f
                break
        
        if zonales:
            print(f"📂 Zonales: {zonales.Name} (Total: {zonales.Items.Count})")
            print(f"   EntryID: {zonales.EntryID}")
            
            # Intento alternativo: Usar GetTable (a veces bypass cache)
            try:
                table = zonales.GetTable()
                print(f"   📊 Table Row Count: {table.GetRowCount()}")
            except:
                print("   ❌ No se pudo obtener Table.")

    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    debug_mailbox()
