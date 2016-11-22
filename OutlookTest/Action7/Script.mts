'Starting MS Outlook
Call StartOutlook()

'Test for switching between
'MS Outlook applications
Set applicationNames = appNames()

Call switchApp(applicationNames.Calendar)
Wait(0.5)
Call switchApp(applicationNames.Tasks)
Wait(0.5)
Call switchApp(applicationNames.Mail)
Wait(0.5)

'Close MS Outlook
Call StopOutlook()

