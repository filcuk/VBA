Attribute VB_Name = "modAdmin"
Option Explicit

Sub AdminSwitch()
    Dim Col As Range
    Dim Row As Range
    Dim Act As Boolean
    
    On Error Resume Next
    Set Col = [AdmCol]
    Set Row = [AdmRow]
    On Error GoTo AdminSwitch_Err
    
    If Col Is Nothing And Row Is Nothing Then
        MsgBox "No admin setup on current sheet."
        GoTo AdminSwitch_Exit
    End If
    
    If Not Col Is Nothing Then
        Act = Not Col.EntireColumn.hidden
        Col.EntireColumn.hidden = Act
    End If
    
    If Not Row Is Nothing Then
        If Col Is Nothing Then
            Act = Not Row.EntireRow.hidden
        End If
        
        Row.EntireRow.hidden = Act
    End If
    
AdminSwitch_Exit:
    Exit Sub
    
AdminSwitch_Err:
    MsgBox "Unable to access admin."
    Resume AdminSwitch_Exit
End Sub
