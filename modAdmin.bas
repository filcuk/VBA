Attribute VB_Name = "modAdmin"
Option Explicit

' Admin
' Function: Show/hide administrative ranges easily
' Usage:    Use named ranges scoped to worksheet
'
Sub Admin()
Dim Rows As Range
Dim Cols As Range
Dim State As Boolean
    
    On Error Resume Next
    Set Rows = ActiveSheet.Range("admRows")
    Set Cols = ActiveSheet.Range("admCols")
    State = Not Rows.EntireRow.Hidden
    State = Not Cols.EntireColumn.Hidden
    On Error GoTo 0
    
    If Rows Is Nothing And Cols Is Nothing Then GoTo Admin_Exit
    
    On Error Resume Next
    Application.ScreenUpdating = False
    Cols.EntireColumn.Hidden = State
    Rows.EntireRow.Hidden = State
    On Error GoTo 0
    
Admin_Exit:
    Application.ScreenUpdating = True
    Exit Sub
End Sub

