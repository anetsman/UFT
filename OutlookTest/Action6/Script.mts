'Starting MS Outlook
Call StartOutlook()

Set tbNames = tabNames()

Call clickOnTab(tbNames.View)
Wait(0.5)
Call clickOnTab(tbNames.Home)
Wait(0.5)
Call clickOnTab(tbNames.Folder)
Wait(0.5)

'Close MS Outlook
Call StopOutlook()

