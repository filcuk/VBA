Attribute VB_Name = "modExport"
Option Explicit

' Description:
' Copy worksheet to a new workbook and save it in a user-designated folder
'
Sub Export()
Const SheetName As String = "Export"                ' Name of sheet to export
Dim FileName As String: FileName = [ExportName]     ' File name input from named range
Dim Wks As Worksheet
Dim Wbk As Workbook
Dim Dlg As Variant
    
    On Error GoTo Export_Err
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set Wbk = Application.Workbooks.Add
    Set Wks = ThisWorkbook.Sheets(SheetName)
    Wks.Copy Wbk.Sheets(1)
    
    For Each Wks In Wbk.Worksheets
        If Wks.Name <> SheetName Then
            Wks.Delete
        End If
    Next
    
    Dlg = Application.GetSaveAsFilename(FileName, "Excel Files (*.xlsx), *.xlsx")
    If Dlg <> False Then
        Wbk.SaveAs Dlg
    Else
'        Wbk.Close SaveChanges:=False
    End If
    
Export_Exit:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub
Export_Err:
    MsgBox "Export failed.", vbCritical, "Export"
    Resume Export_Exit
End Sub
