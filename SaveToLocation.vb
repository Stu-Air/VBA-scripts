Attribute VB_Name = "SAVE"
Sub Save()
Application.ScreenUpdating = False
    
    Dim XLSM_FileName As String
    Dim UserName As String
    Dim Filename As String
    Dim Ref As String
       
    UserName = Environ$("username")
    Filename = ActiveSheet.Range("B8")
    Ref = ActiveSheet.Range("B11")

XLSM_FileName = "C:\Users\" & UserName & "\OneDrive\Hard Drive\01 - Quotations\General Quotations\" & Format(Now, "YYYY.MM.DD") & "-" & Filename & "-" & Ref & ".xlsm"

    ActiveWorkbook.SaveCopyAs Filename:=XLSM_FileName

    
   Application.DisplayAlerts = False
   Application.Quit                   'Closes all excel applications
         
End Sub





