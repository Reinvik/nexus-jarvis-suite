import os

onedrive_base = os.path.join(os.path.expanduser("~"), "OneDrive - CIAL Alimentos")
target_folder = os.path.join(onedrive_base, "Archivos de Operación  Outbound CD - 16.-Inventario Critico")

print(f"Searching in: {target_folder}")

if os.path.exists(target_folder):
    for root, dirs, files in os.walk(target_folder):
        for file in files:
            if "Revision" in file or "Cambiados" in file or "Zonales" in file:
                print(f"FOUND: {os.path.join(root, file)}")
else:
    print("Folder not found.")
