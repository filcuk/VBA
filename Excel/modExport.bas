Attribute VB_Name = "modExportWks"
Option Explicit

' Description:
' Copy worksheet to a new workbook and save it in a user-designated folder
'
' Usage:
Sub ExportSample()
Dim Result As Boolean
    Result = ExportWorksheet("Sample", "Sample " & Format(Now(), "yyyymmddhhMM"))
End Sub
'
Private Function ExportWorksheet(ByVal pSheetName As String, ByVal pExportName As String) As Boolean
Const cDiscardOnSaveCancel As Boolean = True    ' Discard if dialog cancelled by user
Const cDiscardOnError As Boolean = True         ' Discard on critical error
Dim Success As Boolean
Dim Wks As Worksheet
Dim Wbk As Workbook
Dim Dlg As Variant
    
    On Error GoTo Export_Err
    Success = False
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    Set Wbk = Application.Workbooks.Add
    Set Wks = ThisWorkbook.Sheets(pSheetName)
    Wks.Copy Wbk.Sheets(1)
    
    For Each Wks In Wbk.Worksheets
        If Wks.Name <> pSheetName Then
            Wks.Delete
        End If
    Next
    
    Dlg = Application.GetSaveAsFilename(pExportName, "Excel Files (*.xlsx), *.xlsx")
    If Dlg <> False Then
        Wbk.SaveAs Dlg
        Success = True
    ElseIf cDiscardOnSaveCancel Then
        Wbk.Close SaveChanges:=False
        Success = False
    End If
    
Export_Exit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    ExportWorksheet = Success
    Exit Function
Export_Err:
    Success = False
    If cDiscardOnError And Not Wbk Is Nothing Then
        Wbk.Close SaveChanges:=False
    End If
    MsgBox "Export failed.", vbCritical, "Export"
    Resume Export_Exit
End Function
