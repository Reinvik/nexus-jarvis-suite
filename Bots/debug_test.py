import sys
import win32com.client
print("Hello from Debug")
try:
    print("Connecting to Outlook...")
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    print("Connected.")
    folder = ns.GetDefaultFolder(6)
    print("Inbox found.")
    
    # Try finding folder
    target = None
    for f in folder.Folders:
        if "perdida" in f.Name.lower():
            target = f
            print(f"Target found: {f.Name}")
            break
            
    if target:
        print("Sorting...")
        items = target.Items
        items.Sort("[ReceivedTime]", True)
        print("Sorted.")
        item = items.GetFirst()
        if item:
            print(f"First Item: {item.Subject} - {item.ReceivedTime}")
            try:
                print(f"HTMLBody len: {len(getattr(item, 'HTMLBody', ''))}")
            except:
                print("HTMLBody failed")
                
except Exception as e:
    print(f"Error: {e}")
