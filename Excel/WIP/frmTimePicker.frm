VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTimePicker 
   Caption         =   "Date and Time"
   ClientHeight    =   2160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6240
   OleObjectBlob   =   "frmTimePicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTimePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Dim mlngDayOffset As Long   ' used for calendar scroll
Dim mvarDays As Variant    ' date value array (1-3,1-7)
Dim mbolTimeAm As Boolean
Dim mbolValidTime As Boolean
Dim mbolValidDate As Boolean
Dim mbolClosedByX As Boolean

Private Sub btnConfirm_Click()
'    Dim strMsg As String
'    strMsg = "Real: " & txtDate.Value & vbCrLf & _
'            "Form: " & Format(txtDate.Value, "dd-mm-yyyy") & vbCrLf & _
'            vbCrLf & _
'            "Real: " & txtTime.Value & vbCrLf & _
'            "Form: " & Format(txtTime.Value, "hh:mm")
'    MsgBox strMsg, vbOKOnly, "output"
    
    If Not mbolValidTime Or Not mbolValidDate Then
        btnConfirm.Enabled = False
        Exit Sub
    End If
    
    Call Confirm
End Sub

Private Sub btnConfirm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Not mbolValidTime Or Not mbolValidDate Then Exit Sub
        Call Confirm
    ElseIf KeyCode = vbKeyEscape Then
        mbolClosedByX = True
        Call Confirm
    End If
End Sub
Private Sub txtDate_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Not mbolValidTime Or Not mbolValidDate Then Exit Sub
        Call Confirm
    ElseIf KeyCode = vbKeyEscape Then
        mbolClosedByX = True
        Call Confirm
    End If
End Sub
Private Sub txtTime_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Not mbolValidTime Or Not mbolValidDate Then Exit Sub
        Call Confirm
    ElseIf KeyCode = vbKeyEscape Then
        mbolClosedByX = True
        Call Confirm
    End If
End Sub

Private Sub lblClrDate_Click()
    txtDate.Value = ""
End Sub
Private Sub lblClrTime_Click()
    txtTime.Value = ""
End Sub

Private Sub lblCheese_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblCheese")
End Sub

Private Sub lblDay11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblDay11")
End Sub
Private Sub lblDay12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblDay12")
End Sub
Private Sub lblDay13_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblDay13")
End Sub
Private Sub lblDay14_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblDay14")
End Sub
Private Sub lblDay15_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblDay15")
End Sub
Private Sub lblDay16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblDay16")
End Sub
Private Sub lblDay17_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblDay17")
End Sub
Private Sub lblDay21_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblDay21")
End Sub
Private Sub lblDay22_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblDay22")
End Sub
Private Sub lblDay23_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblDay23")
End Sub
Private Sub lblDay24_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblDay24")
End Sub
Private Sub lblDay25_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblDay25")
End Sub
Private Sub lblDay26_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblDay26")
End Sub
Private Sub lblDay27_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblDay27")
End Sub
Private Sub lblDay31_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblDay31")
End Sub
Private Sub lblDay32_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblDay32")
End Sub
Private Sub lblDay33_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblDay33")
End Sub
Private Sub lblDay34_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblDay34")
End Sub
Private Sub lblDay35_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblDay35")
End Sub
Private Sub lblDay36_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblDay36")
End Sub
Private Sub lblDay37_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblDay37")
End Sub

Private Sub lblHour01_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblHour01")
End Sub
Private Sub lblHour02_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblHour02")
End Sub
Private Sub lblHour03_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblHour03")
End Sub
Private Sub lblHour04_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblHour04")
End Sub
Private Sub lblHour05_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblHour05")
End Sub
Private Sub lblHour06_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblHour06")
End Sub
Private Sub lblHour07_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblHour07")
End Sub
Private Sub lblHour08_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblHour08")
End Sub
Private Sub lblHour09_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblHour09")
End Sub
Private Sub lblHour10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblHour10")
End Sub
Private Sub lblHour11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblHour11")
End Sub
Private Sub lblHour12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblHour12")
End Sub

Private Sub lblMin00_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblMin00")
End Sub
Private Sub lblMin15_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblMin15")
End Sub
Private Sub lblMin30_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblMin30")
End Sub
Private Sub lblMin45_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblMin45")
End Sub

Private Sub lblAm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblAm")
End Sub
Private Sub lblPm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call EventMouseMove("lblPm")
End Sub

Private Sub lblDay11_Click()
    Call EventDate("lblDay11")
End Sub
Private Sub lblDay12_Click()
    Call EventDate("lblDay12")
End Sub
Private Sub lblDay13_Click()
    Call EventDate("lblDay13")
End Sub
Private Sub lblDay14_Click()
    Call EventDate("lblDay14")
End Sub
Private Sub lblDay15_Click()
    Call EventDate("lblDay15")
End Sub
Private Sub lblDay16_Click()
    Call EventDate("lblDay16")
End Sub
Private Sub lblDay17_Click()
    Call EventDate("lblDay17")
End Sub
Private Sub lblDay21_Click()
    Call EventDate("lblDay21")
End Sub
Private Sub lblDay22_Click()
    Call EventDate("lblDay22")
End Sub
Private Sub lblDay23_Click()
    Call EventDate("lblDay23")
End Sub
Private Sub lblDay24_Click()
    Call EventDate("lblDay24")
End Sub
Private Sub lblDay25_Click()
    Call EventDate("lblDay25")
End Sub
Private Sub lblDay26_Click()
    Call EventDate("lblDay26")
End Sub
Private Sub lblDay27_Click()
    Call EventDate("lblDay27")
End Sub
Private Sub lblDay31_Click()
    Call EventDate("lblDay31")
End Sub
Private Sub lblDay32_Click()
    Call EventDate("lblDay32")
End Sub
Private Sub lblDay33_Click()
    Call EventDate("lblDay33")
End Sub
Private Sub lblDay34_Click()
    Call EventDate("lblDay34")
End Sub
Private Sub lblDay35_Click()
    Call EventDate("lblDay35")
End Sub
Private Sub lblDay36_Click()
    Call EventDate("lblDay36")
End Sub
Private Sub lblDay37_Click()
    Call EventDate("lblDay37")
End Sub

Private Sub lblHour01_Click()
    Call EventTime("lblHour01")
End Sub
Private Sub lblHour02_Click()
    Call EventTime("lblHour02")
End Sub
Private Sub lblHour03_Click()
    Call EventTime("lblHour03")
End Sub
Private Sub lblHour04_Click()
    Call EventTime("lblHour04")
End Sub
Private Sub lblHour05_Click()
    Call EventTime("lblHour05")
End Sub
Private Sub lblHour06_Click()
    Call EventTime("lblHour06")
End Sub
Private Sub lblHour07_Click()
    Call EventTime("lblHour07")
End Sub
Private Sub lblHour08_Click()
    Call EventTime("lblHour08")
End Sub
Private Sub lblHour09_Click()
    Call EventTime("lblHour09")
End Sub
Private Sub lblHour10_Click()
    Call EventTime("lblHour10")
End Sub
Private Sub lblHour11_Click()
    Call EventTime("lblHour11")
End Sub
Private Sub lblHour12_Click()
    Call EventTime("lblHour12")
End Sub

Private Sub lblMin00_Click()
    Call EventTime("lblMin00")
End Sub
Private Sub lblMin15_Click()
    Call EventTime("lblMin15")
End Sub
Private Sub lblMin30_Click()
    Call EventTime("lblMin30")
End Sub
Private Sub lblMin45_Click()
    Call EventTime("lblMin45")
End Sub

Private Sub lblAm_Click()
    Call EventTime("lblAm")
End Sub
Private Sub lblPm_Click()
    Call EventTime("lblPm")
End Sub

Private Sub lblWeek1_Click()
    mlngDayOffset = mlngDayOffset - 7
    Call InitializeDates(Date, mlngDayOffset)
End Sub
Private Sub lblWeek2_Click()
    Call InitializeDates(Date)
End Sub
Private Sub lblWeek3_Click()
    mlngDayOffset = mlngDayOffset + 7
    Call InitializeDates(Date, mlngDayOffset)
End Sub

Private Sub txtDate_Change()
    Dim lngErr As Long

    On Error Resume Next
    txtDateValue = DateValue(txtDate.Value)
    lngErr = Err.Number
    On Error GoTo 0
    
    With txtDate
        If .Value = "" Then
            txtDateValue = ""
            
            .BorderColor = vbWindowFrame
            .BackColor = RGB(240, 240, 240)
            
            mbolValidDate = True
        ElseIf lngErr <> 0 Then
            txtDateValue = ""
            
            .BorderColor = RGB(200, 0, 0)
            .BackColor = vbWhite
            
            mbolValidDate = False
        Else
            .BorderColor = vbWindowFrame
            .BackColor = vbWhite
            
            mbolValidDate = True
        End If
    End With
    
    If Not mbolValidDate Or Not mbolValidTime Then
        btnConfirm.Enabled = False
    Else
        btnConfirm.Enabled = True
    End If
End Sub

Private Sub txtTime_Change()
    Dim lngErr As Long
    
    On Error Resume Next
    txtTimeValue = TimeValue(txtTime.Value)
    lngErr = Err.Number
    On Error GoTo 0
    
    With txtTime
        If .Value = "" Then
            txtTimeValue = ""
            
            .BorderColor = vbWindowFrame
            .BackColor = RGB(240, 240, 240)
            
            mbolValidTime = True
        ElseIf lngErr <> 0 Then
            txtTimeValue = ""
            
            .BorderColor = RGB(200, 0, 0)
            .BackColor = vbWhite
            
            mbolValidTime = False
        Else
            .BorderColor = vbWindowFrame
            .BackColor = vbWhite
            
            mbolValidTime = True
        End If
    End With
    
    If Not mbolValidDate Or Not mbolValidTime Then
        btnConfirm.Enabled = False
    Else
        btnConfirm.Enabled = True
    End If
End Sub

Private Sub UserForm_Initialize()
    ' center on caller
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
    
    ' initialize values
    mlngDayOffset = 0   ' calendar scroll
    Call InitializeDates(Date, 0)
    Call InitialValues(ActiveCell.Value)      '@ INPUT
    mbolTimeAm = True
    
    mbolClosedByX = False
End Sub

Private Sub Confirm()
    Dim dblOut As Double
    
    On Error Resume Next
    dblOut = DateValue(txtDate.Value)
    dblOut = dblOut + TimeValue(txtTime.Value)
    On Error GoTo 0
    
    If mbolClosedByX Then
        Unload Me
    Else
        ActiveCell.Value = IIf(dblOut = 0, "", dblOut)
        Unload Me
    End If
End Sub

Private Sub InitialValues(ByVal varInput As Variant)
    If IsDate(Format(varInput, "dd-mm-yyyy hh:mm")) Then
        txtDate.Value = IIf(varInput < 1, "", Format(varInput, "dd-mm-yyyy"))
        txtTime.Value = IIf(varInput - CLng(varInput) = 0, "", Format(varInput, "hh:mm"))
    Else
        txtDate.Value = ""
        txtTime.Value = ""
    End If
    
    If txtDate.Value <> "" And txtTime.Value = "" Then
        txtDate.SetFocus
    Else
        txtTime.SetFocus
    End If
End Sub

Private Sub EventDate(ByVal strCtrl As String)
    Dim ctlCtrl As MSForms.control
    
    ' get control and handle exceptions
    Set ctlCtrl = GetCtrlByName(strCtrl)
    If ctlCtrl Is Nothing Then
        MsgBox "Err: Unhandled control.", vbCritical + vbOKOnly, "Err"
        Exit Sub
    End If
    
    If Not IsNumeric(ctlCtrl.Caption) Then
        MsgBox "Err: Control value type mismatch.", vbCritical + vbOKOnly, "Err"
        Exit Sub
    End If
    
    ' get control position = array value position
    Dim lngWeek As Long, lngDay As Long
    Dim strNums As String
    
    strNums = Right(ctlCtrl.Name, 2)
    lngWeek = CLng(Left(strNums, 1))
    lngDay = CLng(Right(strNums, 1))
    
    txtDate.Text = Format(mvarDays(lngWeek, lngDay), "dd-mm-yyyy")
End Sub

Private Sub EventTime(ByVal strCtrl As String)
    Dim ctlCtrl As MSForms.control
    
    ' get control and handle exceptions
    Set ctlCtrl = GetCtrlByName(strCtrl)
    If ctlCtrl Is Nothing Then
        MsgBox "Err: Unhandled control.", vbCritical + vbOKOnly, "Err"
        Exit Sub
    End If
    
    '
    Dim strMin As String, strHour As String
    Dim strOutput As String
    
    ' get pressed button values
    If InStr(1, ctlCtrl.Name, "Am", vbTextCompare) <> 0 Then
        mbolTimeAm = True
        Call SetHours(True)
    ElseIf InStr(1, ctlCtrl.Name, "Pm", vbTextCompare) <> 0 Then
        mbolTimeAm = False
        Call SetHours(False)
    ElseIf InStr(1, ctlCtrl.Name, "Min", vbTextCompare) <> 0 Then
        strMin = Right(ctlCtrl.Name, 2)
    ElseIf InStr(1, ctlCtrl.Name, "Hour", vbTextCompare) <> 0 Then
        If mbolTimeAm Then
            strHour = Right(ctlCtrl.Name, 2)
        Else
            strHour = Hour(TimeValue(CLng(Right(ctlCtrl.Name, 2)) & ":00:00") + TimeValue("12:00:00"))
        End If
    End If
    
    ' get existing field values
    On Error Resume Next
    If strHour = "" Then strHour = Hour(TimeValue(txtTimeValue.Value))
    If strMin = "" Then strMin = Minute(TimeValue(txtTimeValue.Value))
    On Error GoTo 0
    
    '
    strOutput = IIf(Len(strHour) = 1, "0", "") & IIf(strHour = "", "00", strHour) & ":"
    strOutput = strOutput & IIf(strMin = "" Or strMin = "0", "00", strMin)
    txtTime.Text = strOutput
End Sub

Private Sub SetHours(ByVal bolDay As Boolean)
    With Me
        .lblHour01.Caption = IIf(bolDay, "1", "13")
        .lblHour02.Caption = IIf(bolDay, "2", "14")
        .lblHour03.Caption = IIf(bolDay, "3", "15")
        .lblHour04.Caption = IIf(bolDay, "4", "16")
        .lblHour05.Caption = IIf(bolDay, "5", "17")
        .lblHour06.Caption = IIf(bolDay, "6", "18")
        .lblHour07.Caption = IIf(bolDay, "7", "19")
        .lblHour08.Caption = IIf(bolDay, "8", "20")
        .lblHour09.Caption = IIf(bolDay, "9", "21")
        .lblHour10.Caption = IIf(bolDay, "10", "22")
        .lblHour11.Caption = IIf(bolDay, "11", "23")
        .lblHour12.Caption = IIf(bolDay, "12", "24")
    End With
    
    '@
    ' get existing field values
    Dim lngHour As Long, lngMin As Long
    Dim strOutput As String
    
    On Error Resume Next
    lngHour = Hour(TimeValue(txtTimeValue.Value))
    lngMin = Minute(TimeValue(txtTimeValue.Value))
    On Error GoTo 0

    '
    If lngHour = 0 Then
        lngHour = IIf(bolDay, 12, 0)
    ElseIf lngHour > 12 Then
        lngHour = IIf(bolDay, lngHour - 12, lngHour)
    Else
        lngHour = IIf(bolDay, lngHour, lngHour + 12)
    End If
    
    strOutput = IIf(lngHour < 10, "0", "") & CStr(lngHour) & ":"
    strOutput = strOutput & IIf(lngMin < 10, "0", "") & CStr(lngMin)
    txtTime.Text = strOutput
End Sub

'
Private Sub EventMouseMove(ByVal strCtrl As String)
    Dim ctlCtrl As MSForms.control
    
    ' get control and handle exceptions
    Set ctlCtrl = GetCtrlByName(strCtrl)
    If ctlCtrl Is Nothing Then
        MsgBox "Err: Unhandled control.", vbCritical + vbOKOnly, "Err"
        Exit Sub
    End If
    
    If ctlCtrl.Name = "lblCheese" Then
        Call RestoreLabel
    Else
        ctlCtrl.BackColor = RGB(200, 200, 200)
    End If
End Sub

Private Sub RestoreLabel()
    Dim ctlLbl As MSForms.control
    
    For Each ctlLbl In Me.Controls
        If TypeName(ctlLbl) = "Label" Then
            If InStr(1, ctlLbl.Name, "Day", vbTextCompare) <> 0 Or _
                InStr(1, ctlLbl.Name, "Hour", vbTextCompare) <> 0 Or _
                InStr(1, ctlLbl.Name, "Min", vbTextCompare) <> 0 Or _
                InStr(1, ctlLbl.Name, "Am", vbTextCompare) <> 0 Or _
                InStr(1, ctlLbl.Name, "Pm", vbTextCompare) <> 0 Then
                
                If ctlLbl.BackColor = RGB(200, 200, 200) Then
                    ctlLbl.BackColor = RGB(240, 240, 240)
                End If
            End If
        End If
    Next ctlLbl
End Sub

' returns control object, if found
Private Function GetCtrlByName(ByVal strName As String) As control
    Dim ctl As MSForms.control
    
    For Each ctl In Me.Controls
        If ctl.Name = strName Then
            Set GetCtrlByName = ctl
            Exit For
        End If
    Next ctl
End Function

Private Sub InitializeDates(ByVal dtmDate As Date, Optional ByVal lngDateOffset As Long = 0)
    Dim ctlCont As MSForms.control
    Dim lngWeek As Long, lngDay As Long
    
    '
    dtmDate = DateAdd("d", lngDateOffset, dtmDate)
    mvarDays = CalendarArr(dtmDate)
    
    ' mark current day, if in view
    Dim lngCWeek As Long, lngCDay As Long
    
    If lngDateOffset >= -7 And lngDateOffset < 0 Then
        lngCDay = Weekday(Date, vbMonday)
        lngCWeek = 3
    ElseIf lngDateOffset = 0 Then
        lngCDay = Weekday(Date, vbMonday)
        lngCWeek = 2
    ElseIf lngDateOffset > 0 And lngDateOffset <= 7 Then
        lngCDay = Weekday(Date, vbMonday)
        lngCWeek = 1
    End If
    
    For lngWeek = LBound(mvarDays, 1) To UBound(mvarDays, 1)
        ' week numbers
        For Each ctlCont In Me.Controls
            If ctlCont.Name = "lblWeek" & CStr(lngWeek) Then _
                ctlCont.Caption = Format(mvarDays(lngWeek, 1), "WW")
        Next ctlCont
        
        ' day numbers
        For lngDay = LBound(mvarDays, 2) To UBound(mvarDays, 2)
            For Each ctlCont In Me.Controls
                If ctlCont.Name = "lblDay" & CStr(lngWeek) & CStr(lngDay) Then
                    ctlCont.Caption = Format(mvarDays(lngWeek, lngDay), "d")
                    ctlCont.BorderStyle = fmBorderStyleNone
                    
                    ' mark current day
'                    Debug.Assert (lngDay <> lngCDay)
                    If lngWeek = lngCWeek And lngDay = lngCDay Then _
                        ctlCont.BorderStyle = fmBorderStyleSingle
                End If
            Next ctlCont
        Next lngDay
    Next lngWeek
End Sub


' returns current week, 1 past and 1 future week's dates
Private Function CalendarArr(ByVal dtmCur As Date) As Variant
    Dim aDays(1 To 3, 1 To 7) As Date
    Dim dtmCurMon As Date
    Dim lngWDay As Long
    Dim lngOffset As Long
    Dim lngWeek As Long, lngDay As Long
    
    lngWDay = Weekday(dtmCur, vbMonday)
    dtmCurMon = DateAdd("d", (lngWDay - 1) * (-1), dtmCur)  ' returns monday's date
    
    lngOffset = -7
    For lngWeek = LBound(aDays, 1) To UBound(aDays, 1)
        For lngDay = LBound(aDays, 2) To UBound(aDays, 2)
            aDays(lngWeek, lngDay) = DateAdd("d", lngOffset, dtmCurMon)
            lngOffset = lngOffset + 1
        Next lngDay
    Next lngWeek
    
    CalendarArr = aDays
End Function
