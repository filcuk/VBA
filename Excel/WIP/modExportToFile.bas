Attribute VB_Name = "modExportToFile"
Option Explicit

Private Sub Sample()
Dim Success As Boolean
    Success = SaveAsCopy("C:\Users\FilipK\Desktop\Export\Test.csv", xlCSV, "Export")
    MsgBox Success
End Sub

Function SaveAsCopy( _
    ByVal pFilePathName As String, _
    ByVal pFormat As XlFileFormat, _
    ByVal pSheets As String, _
    Optional ByVal pWorkbook As String = "" _
)

Const cDelimiter As String = ";"
Const cSaveCopy As Boolean = True
Const cCloseCopy As Boolean = True


Dim SrcWks As Variant
Dim SrcWbk As Workbook
Dim DstWbk As Workbook
Dim Success As Boolean
    
    On Error GoTo SaveAsCopy_Error
    
    If pWorkbook = "" Then
        Set SrcWbk = ThisWorkbook
    Else
        Set SrcWbk = Application.Workbooks(pWorkbook)
    End If
    SrcWks = Split(pSheets, cDelimiter)
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    SrcWbk.Worksheets(SrcWks).Copy
    Set DstWbk = ActiveWorkbook
    
    If cSaveCopy Then
        DstWbk.SaveAs Filename:=pFilePathName, FileFormat:=pFormat
        If cCloseCopy Then DstWbk.Close SaveChanges:=True
    End If
    
    Success = True
    
SaveAsCopy_Exit:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    SaveAsCopy = Success
    Exit Function
SaveAsCopy_Error:
    Success = False
    Resume SaveAsCopy_Exit
End Function
