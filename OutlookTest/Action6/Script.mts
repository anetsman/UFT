'Starting MS Outlook
Call StartOutlook()

Set tab = tabs()
Set tbNames = tabNames()

tab.clickOnTab(tbNames.View)
Wait(1)
tab.clickOnTab(tbNames.Home)
Wait(1)
tab.clickOnTab(tbNames.Folder)
Wait(1)

'Close MS Outlook
Call StopOutlook()

Set objOutlook = Nothing
Set objFolder = Nothing
