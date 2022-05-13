VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private IsBeforeClose As Boolean
Private CancelClose As Boolean

Private Sub Workbook_Open()
    Call Landing(False)
    Call WaitFrm(False, "Loading data...")
    ThisWorkbook.Sheets("Data").ListObjects("tblData").QueryTable.Refresh   ' Set refresh target
    Call WaitFrm(True)
End Sub

' Requires [appAllowNextSave] name to point to a range
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
Dim IsSaved  As Boolean

    Cancel = True
    If Not [appAllowNextSave] = True Then
        Debug.Print "Saveing is disabled. Confirm to re-enable temporarily:"
        Debug.Print "[appAllowNextSave] = True"
    Else
        Application.EnableEvents = False
        [appAllowNextSave] = False
        Call Landing(True)
        
        If SaveAsUI Then
            IsSaved = Application.Dialogs(xlDialogSaveAs).Show
            If Not IsSaved Then CancelClose = True
        Else
            ThisWorkbook.Save
        End If
        
        Application.EnableEvents = True
        Call Landing(False)
    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Debug.Assert False
    If CancelClose Then Cancel = True
End Sub

Private Sub Landing(pEnable As Boolean)
Const cWksLanding As String = "Landing" ' Shown in place of hidden sheets
Const cWksActivate As String = "Data"   ' Sheet to land on once active content is enabled
Const cWksAbout As String = "About"     ' Sheet to always remain visible
Dim Wbk As Workbook
Dim Wks As Worksheet
Dim Lan As Worksheet
Dim Act As Worksheet
Dim Abt As Worksheet

    Set Wbk = ThisWorkbook
    Set Lan = Wbk.Sheets(cWksLanding)
    Set Act = Wbk.Sheets(cWksActivate)
    Set Abt = Wbk.Sheets(cWksAbout)
    
    ' At least one sheet must remain visible
    Lan.Visible = xlSheetVisible
    For Each Wks In ThisWorkbook.Sheets
        If Wks.Name <> Lan.Name _
        And Wks.Name <> Abt.Name Then
            Wks.Visible = IIf(pEnable, xlSheetHidden, xlSheetVisible)
        End If
    Next
    Lan.Visible = IIf(pEnable, xlSheetVisible, xlSheetHidden)
    If Not Abt.Visible Then Abt.Visible = xlSheetVisible
    
    ' Activate sheet
    If pEnable And Lan.Visible = xlSheetVisible Then
        Lan.Activate
    ElseIf Act.Visible = xlSheetVisible Then
        Act.Activate
    Else
        Abt.Activate
    End If
End Sub

