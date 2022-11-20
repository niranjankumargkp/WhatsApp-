# SendWhatsAppmessagetoexcel-

Sub WhatsAppMsg1()
Dim LastRow As Long
Dim i As Integer
Dim strip As String
Dim strPhoneNumber As String
Dim strmessage As String
Dim strPostData As String
Dim IE As Object

LastRow = Range("A" & Rows.Count).End(xlUp).Row
For i = 2 To LastRow


strPhoneNumber = Sheets("Data").Cells(i, 1).Value
strmessage = Sheets("Data").Cells(i, 2).Value

'IE.navigate "whatsapp://send?phone=phone_number&text=your_message"

strPostData = "https://web.whatsapp.com/send/?phone=" & strPhoneNumber & "&text=" & strmessage & "&type=phone_number&app_absent=0"
Set IE = CreateObject("InternetExplorer.Application")
Range("c1").Value = strPostData
'IE.navigate "https://web.whatsapp.com/send/?phone=9454006081&text=hello&type=phone_number&app_absent=0"
IE.navigate strPostData
Application.Wait Now() + TimeSerial(0,0,5)

SendKeys "~"

Next i
End Sub
