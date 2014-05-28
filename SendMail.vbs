Set MyApp = CreateObject("Outlook.Application")
Set MyItem = MyApp.CreateItem(0) 'MailItem
With MyItem
.To = "name@mail.com"
.Subject = "Subject"
.ReadReceiptRequested = False
.HTMLBody = "Test mail send from vbs file via outlook"
End With
MyItem.Send
