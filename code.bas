Option Explicit
Sub Send_Mails()

Dim sh As Worksheet
Set sh = ThisWorkbook.Sheets("Sheet2")

Dim OA As Object
Dim msg As Object

Set OA = CreateObject("Outlook.Application")
Dim i As Integer
Dim last_row As Integer

last_row = Application.WorksheetFunction.CountA(sh.Range("A:A"))

For i = 2 To last_row
Set msg = OA.createitem(0)

msg.to = sh.Range("E" & i).Value
msg.cc = sh.Range("F" & i).Value
msg.Subject = sh.Range("G" & i).Value
msg.Body = sh.Range("C" & i).Value & " " & sh.Range("D" & i).Value & ":" & vbCr & vbCr & sh.Range("H" & i).Value
msg.HTMLBody = msg.HTMLBody & "<img  src='C:\Users\x.x\Desktop\Captura.png'>"

If sh.Range("I" & i).Value <> "" Then
msg.attachments.Add sh.Range("I" & i).Value
End If

msg.display

Next i

MsgBox "All the mails have been sent successfully"


End Sub
