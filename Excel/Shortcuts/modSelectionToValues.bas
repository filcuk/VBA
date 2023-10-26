Attribute VB_Name = "modSelectionToValues"
Option Explicit

' Converts selection to values
' Bind this to a CTRL+? shortcut
Sub SelectionToValues()
Attribute SelectionToValues.VB_ProcData.VB_Invoke_Func = "r\n14"
    Dim Sel As Range
    Dim Msg As String
    Const Con As Boolean = True ' Enables confirmation
    
    On Error Resume Next
    Set Sel = Selection
    
    Sel.Copy
    If Con Then
        Msg = MsgBox("Convert " & Sel.Address(0, 0) & " to values?", vbQuestion + vbOKCancel)
        If Msg <> vbOK Then GoTo SelectionToValues_Exit
    End If
    Sel.PasteSpecial xlPasteValues
    
SelectionToValues_Exit:
    On Error GoTo 0
End Sub
