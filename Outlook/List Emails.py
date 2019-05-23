import win32com.client

olApp = win32com.client.Dispatch('Outlook.Application')
ns = olApp.GetNamespace('MAPI')

"""
constants
https://docs.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders

"""
olInboxFolder = ns.GetDefaultFolder(6)

for olMail in olInboxFolder.Items:
    print(olMail.Subject)
