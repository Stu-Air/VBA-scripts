' ##################################################
' ########### attach saved pdf to email ############ 
' ##################################################

    'SaveName = Sheets("Cover Sheet").range("I17")
    'attachment = FolderLocation & "\" & SaveName & "-" & "CoverSheet.pdf"

    'Calling the sub/function "Call email(CoverSheet, SaveName)"

    'Email to outlook can change "Outlook.Application" to any application
    'attachment as string linked to above on page calling the sub/function
    'SaveName as String to above on page calling the sub/function

Sub email(attachment As String, SaveName As String)

 'Email the Quote/Cover Sheet
    Dim objOutlook As Object
    Dim objMail As Object
    Dim signature As String
    Dim oWB As Workbook
    Set oWB = ActiveWorkbook

    Set objOutlook = CreateObject("Outlook.Application") '<----- Change for application (confirm works with Thunderbird)
    Set objMail = objOutlook.CreateItem(0)
        
    With objMail
        .Display
    End With
        signature = objMail.HTMLbody
    With objMail
        .Subject = "Estimate" & " " & SaveName & " " & "From SCS"
        .HTMLbody = "<font face=" & Chr(34) & "Calibri" & Chr(34) & " size=" & Chr(34) & 4 & Chr(34) & ">" & "Hello," & "<br>" & "Please see attached your Estimate" & " " & SaveName & " " & "<br>" & signature & "</font>"
        .Attachments.Add attachment '<---- Change to suit file for emailing.
        .Save
        .Display
    End With

    Set objOutlook = Nothing
    Set objMail = Nothing

End Sub



' ####################################
' ########### Save to PDF ############ 
' ####################################

    'SheetName = "Cover Sheet"
    'fileNameLocation = FolderLocation & "\" & SaveName & "-" & "CoverSheet.pdf"

    'Calling the sub/function "Call PDF("Cover Sheet", CoverSheet)"

    'fileNameLocation as string linked to above on page calling the sub/function
    'sheetName as String to above on page calling the sub/function

Sub PDF(sheetName As String, fileNameLocation As String)

 'Save Cover Sheet
    
    Sheets(sheetName).ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
    fileNameLocation, Quality:=xlQualityStandard, IncludeDocProperties _
    :=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    
End Sub

' ###################################
' ########### Copy paste ############ 
' ###################################

    
    'Calling the sub/function "call paste("Picking sheet", "A1:J62", "Screen1", "A1:J62")"

Sub paste(sheetCopy As String, rangeCopy As String, sheetPaste As String, rangePaste As String)

Sheets(sheetCopy).range(rangeCopy).Copy
    With Sheets(sheetPaste).range(rangePaste)
    .PasteSpecial paste:=xlPasteValuesAndNumberFormats, _
                 Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    .PasteSpecial , paste:=xlPasteColumnWidths, paste:=xlPasteFormats
    End With
End Sub







