'Starting MS Outlook
Call StartOutlook()

'Check the total number of reminders available
Set objAllReminders = objOutlook.Reminders
Call LogFunction("Reminders", objAllReminders.Count)
msgbox objAllReminders.Count
'Check the version of the Outlook installed
Call LogFunction("Version", objOutlook.Version)
msgbox  objOutlook.Version
'Check the Product Code - Microsoft Outlook globally unique identifier (GUID).
Call LogFunction("Product Code", objOutlook.ProductCode)
msgbox  objOutlook.ProductCode

'Close MS Outlook
Call StopOutlook()

Set objOutlook = Nothing
Set objFolder = Nothing
Set objAllReminders = Nothing
