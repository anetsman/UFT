'Creating the Outlook object and running the Outlook
Call StartOutlook()

'Switching to Calendar
'Set applications = apps()
Set applicationNames = appNames()
Call switchApp(applicationNames.Calendar)

'Creating New Appointment
Call createNewAppointment()

'Setting up the New Appointment
Set timeReccurence = reccurenceTime()
Call setupNewAppointment("qwerty", "qwerty", timeReccurence.Daily, "Netsman, Oleksandr <ANetsman@luxoft.com>")

Reporter.ReportEvent micPass,"Pass","Shedule an appointment"

'Close MS Outlook after waiting for sending an email
Wait(2)
Call StopOutlook()
'objOutlook.Quit()

Set objOutlook = Nothing
Set objFolder = Nothing


'UIAWindow("Календар - ANetsman@luxoft.com").UIAButton("New Appointment").Click
'UIAWindow("Календар - ANetsman@luxoft.com").UIAButton("New Appointment").Click
'UIAWindow("Untitled - Appointment").UIAEdit("Subject:").SetValue "qwe"
'UIAWindow("qwe - Appointment").UIAComboBox("Location:").Select "Location:"
'UIAWindow("qwe - Appointment").UIAButton("Recurrence...").Set "On"
'UIAWindow("qwe - Appointment").UIAWindow("Appointment Recurrence").UIARadioButton("Daily").Select
'UIAWindow("qwe - Appointment").UIAWindow("Appointment Recurrence").UIAButton("OK").Click
'UIAWindow("qwe - Meeting Series").UIAEdit("To").SetValue "Netsman, Oleksandr; "
