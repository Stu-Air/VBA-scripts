Attribute VB_Name = "PDF"
Sub PDF()
Application.ScreenUpdating = False

 
    Dim PDF_FileName As String
    Dim UserName As String
    Dim Filename As String
    Dim Ref As String

        
     UserName = Environ$("username")
     Filename = ActiveSheet.Range("B8")
     Ref = ActiveSheet.Range("B11")
     
     PDF_FileName = "C:\Users\" & UserName & "\OneDrive\Hard Drive\01 - Quotations\General Quotations\" & Format(Now, "YYYY.MM.DD") & "-" & Filename & "-" & Ref & ".pdf"
     
 
     ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
         PDF_FileName, Quality:=xlQualityStandard, IncludeDocProperties _
         :=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
  
   Application.DisplayAlerts = False
   Application.Quit                   'Closes all excel applications

End Sub