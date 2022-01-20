VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HTMLEditForm 
   Caption         =   "Email HTML Editor"
   ClientHeight    =   9015.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14160
   OleObjectBlob   =   "HTMLEditForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "HTMLEditForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ApplyButton_Click()
    Call HTMLEdit.HTMLEditor("Apply")
End Sub

Private Sub ApplySendButton_Click()
    Call HTMLEdit.HTMLEditor("ApplySend")
End Sub

Private Sub CancelButton_Click()
    Me.Hide
End Sub
