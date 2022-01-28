Attribute VB_Name = "modAdmin"
Option Explicit

Private Const mSectionName As String = "Admin"

Sub AdminSwitch()
Dim Col As Range
Dim Row As Range
Dim Act As Boolean
    
    On Error Resume Next
    Set Col = [AdmCol]
    Set Row = [AdmRow]
    On Error GoTo AdminSwitch_Err
    
    If Col Is Nothing And Row Is Nothing Then
        MsgBox "This sheet has no configured admin areas.", vbInformation, mSectionName
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
    MsgBox "Failed to access admin.", vbExclamation, mSectionName
    Resume AdminSwitch_Exit
End Sub

' Switch between development and production
Sub AdminWbkSwitch()
Const Nam As String = "AdmSht"
Dim Act As Boolean
Dim Ini As Boolean: Init = True
Dim Wks As Worksheet
    
    For Each Wks In ThisWorkbook.Worksheets
        If NamedRangeExists(Wks, Nam) Then
            ' Setup initial state
            If Ini Then
                Act = Wks.visible = xlSheetVisible
                Ini = False
            End If
            
            ' Apply changes
            Wks.visible = IIf(Act, xlSheetVeryHidden, xlSheetVisible)
        End If
    Next
    
ProdSwicth_Exit:
    Exit Sub
ProdSwicth_Err:
    MsgBox "Failed to switch admin state.", vbExclamation, mSectionName
End Sub

Private Function NamedRangeExists(ByVal pSheet As Worksheet, ByVal pName As String) As Boolean
Dim Nam As Name
    On Error GoTo NamedRangeExists_Fail
    Set Nam = pSheet.Names(pName)
    NamedRangeExists = True
    Exit Function
NamedRangeExists_Fail:
    NamedRangeExists = False
End Function
