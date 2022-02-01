Attribute VB_Name = "modPanels"
Option Explicit

Sub PanelSwitch_Top()
    Call PanelSwitch(msoBarTop)
End Sub

Sub PanelSwitch_Bottom()
    Call PanelSwitch(msoBarBottom)
End Sub

Sub PanelSwitch_Left()
    Call PanelSwitch(msoBarLeft)
End Sub

Sub PanelSwitch_Right()
    Call PanelSwitch(msoBarRight)
End Sub

Sub PanelSwitch_Menu()
    Call PanelSwitch(msoBarMenuBar)
End Sub

Sub PanelSwitch_Popup()
    Call PanelSwitch(msoBarPopup)
End Sub

Sub PanelSwitch_Floating()
    Call PanelSwitch(msoBarFloating)
End Sub

Private Sub PanelSwitch(ByVal Pos As MsoBarPosition)
Const TITLE As String = "Panel Switch"

Dim Sheet As Worksheet
Dim Target As Range
Dim Panel As String
Dim msg As String
    
    ' Select response
    On Error Resume Next
    Select Case Pos
        Case msoBarTop
            Set Target = [Panel_Top]
            Panel = "Top"
            
        Case msoBarBottom
            Set Target = [Panel_Bottom]
            Panel = "Bottom"
            
        Case msoBarLeft
            Set Target = [Panel_Left]
            Panel = "Left"
            
        Case msoBarRight
            Set Target = [Panel_Right]
            Panel = "Right"
            
        Case msoBarMenuBar
            Set Target = [Panel_Menu]
            Panel = "Menu"
            
        Case msoBarPopup
            Set Target = [Panel_Popup]
            Panel = "Popup"
            
        Case msoBarFloating
            Set Target = [Panel_Floating]
            Panel = "Floating"
            
        Case Else
            Set Target = Nothing
            Panel = "Unknown"
            Debug.Print "PanelSwitch(): Passed parameter mismatch."
            
    End Select
    On Error GoTo Err_PanelSwitch
    
    ' Incorrect range reference
    If Target Is Nothing Then
        msg = Panel & " panel is not available on Sheet [" & ActiveSheet.Name & "]."
        MsgBox msg, vbOKOnly, TITLE
        Exit Sub
    End If
    
    ' Sheeteet is protected
    Set Sheet = Target.Parent
    If Sheet.ProtectContents Then
        msg = "Sheet [" & Sheet.Name & "] is protected."
        MsgBox msg, vbExclamation, TITLE
        Exit Sub
    End If
        
    ' Get proper range dimension
    ' This assumes defined ranges use full w/h selection
    If Target.Columns.Count = Sheet.Columns.Count Then
        Set Target = Target.EntireRow
    Else
        Set Target = Target.EntireColumn
    End If
    
    ' Apply
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Target.hidden = Not Target.hidden
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    Exit Sub
    
' Error handling
Err_PanelSwitch:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox Err.Number & ": " & Err.Description, vbCritical, TITLE
End Sub
