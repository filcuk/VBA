Attribute VB_Name = "HTMLEdit"

'===============================================================================
'Description: Outlook macro to edit the HTML of an email that you are composing.
'
' author : Robert Sparnaaij
' version: 1.0
' website: https://www.howto-outlook.com/howto/edit-html-source-code-email.htm
'===============================================================================

Sub EditHTML()
    HTMLEditor ("Edit")
End Sub

Function HTMLEditor(ByVal Action As String)
    Dim objMail As MailItem, oInspector As Inspector
    Dim msgResult As Integer, msgText As String, msgTitle As String
    msgText = "This is not an editable email in HTML format."
    msgTitle = "Email HTML Editor"
    
    Set oInspector = Application.ActiveInspector
    If oInspector Is Nothing Then
        msgResult = MsgBox(msgText, vbCritical, msgTitle)
    Else
        Set objMail = oInspector.CurrentItem
        With objMail
            If .Sent Then
                msgResult = MsgBox(msgText, vbCritical, msgTitle)
            Else
                If .BodyFormat = olFormatHTML Then
                    Select Case Action
                        Case "Edit"
                            HTMLEditForm.HTMLTextBox.Text = .HTMLBody
                            HTMLEditForm.Show
                        Case "Apply"
                            .HTMLBody = HTMLEditForm.HTMLTextBox.Text
                            HTMLEditForm.Hide
                        Case "ApplySend"
                            If (.Recipients.Count = 0) Or (.Recipients.ResolveAll = False) Or (.Subject = "") Then
                                msgResult = MsgBox("Please specify the recipients and/or the Subject for this message first." _
                                        & vbNewLine & "Choose Apply or Cancel instead.", vbCritical, msgTitle)
                            Else
                                .Close olSave
                                .HTMLBody = HTMLEditForm.HTMLTextBox.Text
                                .Send
                                HTMLEditForm.Hide
                            End If
                    End Select
                Else
                    msgResult = MsgBox(msgText, vbCritical, msgTitle)
                End If
            End If
        End With
	Set objMail = Nothing
    End If
    Set oInspector = Nothing
End Function
