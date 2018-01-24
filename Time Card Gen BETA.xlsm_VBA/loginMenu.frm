VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} loginMenu 
   Caption         =   "Select Job Number"
   ClientHeight    =   7305
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8280.001
   OleObjectBlob   =   "loginMenu.frx":0000
End
Attribute VB_Name = "loginMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub loginButton_Click()
    Me.Hide
End Sub

Private Sub mCancel_Click()
    mainMenu.mCancel_Click
End Sub

Private Sub reqUser_Click()
    Me.Hide
    reqMenu.Show
End Sub

Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .Label1.Caption = "Enter Password:"
        .TextBox1.SetFocus
        .Caption = "LOGIN"
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        mainMenu.mCancel_Click
    End If
End Sub
