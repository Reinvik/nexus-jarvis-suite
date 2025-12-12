import win32com.client
import time
import sys

# Configuración UTF-8 para consola
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

def limpiar_zonales():
    print("🚀 Iniciando limpieza masiva de carpeta Zonales...")
    print("ℹ️  Objetivo: Mover todos los correos LEÍDOS a la carpeta 'Procesados'")
    
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        inbox = namespace.GetDefaultFolder(6) # Inbox
        zonales = None
        for f in inbox.Folders:
            if f.Name.lower() == "zonales":
                zonales = f
                break
        
        if not zonales:
            print("❌ No se encontró la carpeta 'Zonales'")
            return

        procesados = None
        for f in zonales.Folders:
            if f.Name == "Procesados":
                procesados = f
                break
        
        if not procesados:
            procesados = zonales.Folders.Add("Procesados")
            print("📁 Carpeta 'Procesados' creada")

        items = zonales.Items
        total_items = items.Count
        print(f"📬 Total de elementos en Zonales: {total_items}")
        
        if total_items == 0:
            print("✅ La carpeta está vacía.")
            return

        print("⏳ Iniciando movimiento... (Esto puede tomar unos minutos)")
        
        moved_count = 0
        skipped_count = 0
        
        # Iteramos hacia atrás para evitar problemas al mover elementos
        # Los índices en Outlook comienzan en 1
        for i in range(total_items, 0, -1):
            try:
                item = items.Item(i)
                
                # Solo mover si NO está No Leído (es decir, si está Leído)
                if not item.UnRead:
                    item.Move(procesados)
                    moved_count += 1
                    
                    if moved_count % 100 == 0:
                        print(f"   💨 Movidos: {moved_count}...")
                else:
                    skipped_count += 1
                    
            except Exception as e:
                print(f"   ⚠️ Error moviendo item {i}: {e}")
                continue

        print("-" * 40)
        print(f"✅ LIMPIEZA COMPLETADA")
        print(f"📦 Total movidos a Procesados: {moved_count}")
        print(f"📨 Total dejados (No Leídos): {skipped_count}")
        print("-" * 40)

    except Exception as e:
        print(f"❌ Error crítico: {e}")

if __name__ == "__main__":
    limpiar_zonales()
