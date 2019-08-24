import win32com.client as win32

# olApp = win32.Dispatch('Outlook.Application')
olApp = win32.Dispatch('Outlook.Application')
olNS = olApp.GetNamespace('MAPI')
olAcct = olNS.Accounts(2)
# get currentuser
olNS.CurrentUser()

# https://docs.microsoft.com/en-us/office/vba/api/Outlook.OlDefaultFolders

outlook_task_item = 3
tskItem = olApp.CreateItem(outlook_task_item)
tskItem.subject = "Send report"
tskItem.DueDate = '02/02/2019'
tskItem.Categories = 'Office'
tskItem.body = "Hello World"
tskItem.save()
