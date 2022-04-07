Attribute VB_Name = "modFreezePanes"
Option Explicit

' Freeze panes without selecting target range
Function FreezePanes( _
    ByVal pTopLeft As Range, _
    ByVal pUnfreeze As Boolean, _
    ByRef pErrMsg As String _
    ) As Boolean

Dim Win As Window
Dim Success As Boolean
    
    On Error GoTo FreezePanes_Err
    
    For Each Win In pTopLeft.Worksheet.Parent.Windows
        If Win.ActiveSheet.Name <> pTopLeft.Parent.Name Then GoTo NextLoop
        
        With Win
            If .FreezePanes Then .FreezePanes = False
            
            If pUnfreeze Then
                If .SplitRow Then .SplitRow = False
                If .SplitColumn Then .SplitColumn = False
            Else
                If Not ((pTopLeft.Row = 1) And (pTopLeft.Column = 1)) Then
                    .ScrollRow = 1
                    .ScrollColumn = 1
                    .SplitRow = pTopLeft.Row - 1
                    .SplitColumn = pTopLeft.Column - 1
                    .FreezePanes = True
                End If
            End If
        End With
        
NextLoop:
    Next
    
    Success = True
    
FreezePanes_Exit:
    FreezePanes = Success
    Exit Function
FreezePanes_Err:
    Success = False
    pErrMsg = Err.Number & ": " & Err.Description
    Resume FreezePanes_Exit
End Function

Private Sub FreezePanes_Example()
Dim ErrOut As String
    Debug.Print FreezePanes(Selection, True, ErrOut), ErrOut
End Sub
