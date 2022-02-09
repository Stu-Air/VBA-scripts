Attribute VB_Name = "EmailToOutlook"
Sub Email()
Application.ScreenUpdating = False

    Dim objOutlook As Object
    Dim objMail As Object
    Dim signature As String
    Dim oWB As Workbook
    Set oWB = ActiveWorkbook

    Set objOutlook = CreateObject("Outlook.Application")
    Set objMail = objOutlook.CreateItem(0)
        
    With objMail
        .Display
    End With
        signature = objMail.HTMLbody
    With objMail
        .Subject = "Quotation"
        .HTMLbody = "<font face=" & Chr(34) & "Calibri" & Chr(34) & " size=" & Chr(34) & 4 & Chr(34) & ">" & "Hello," & "<br> <br>" & "Please see attached you Quotation" & "<br> <br>" & signature & "</font>"
        .Attachments.Add PDF_FileName    'file link here 
        .Save
        .Display
    End With

    Set objOutlook = Nothing
    Set objMail = Nothing
   Application.DisplayAlerts = False
   Application.Quit                   'Closes all excel applications

End Sub



