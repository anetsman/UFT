'Starting MS Outlook
Call StartOutlook()
Set applications = apps()
Set applicationNames = appNames()

apps.switchApp(applicationNames.Calendar)
Wait(0.5)
apps.switchApp(applicationNames.Tasks)
Wait(0.5)
apps.switchApp(applicationNames.Mail)
Wait(0.5)

'Close MS Outlook
Call StopOutlook()

Set objOutlook = Nothing
Set objFolder = Nothing
