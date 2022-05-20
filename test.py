from email import message
import win32com.client as client

outlook = client.Dispatch('Outlook.Application').GetNamespace("MAPI")

root_folder = outlook.Folders.Item(3)
print (root_folder.Name)

for folder in root_folder.Folders:
    print (folder.Name)

