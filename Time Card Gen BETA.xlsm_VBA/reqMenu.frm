VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} reqMenu 
   Caption         =   "Select Job Number"
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8280.001
   OleObjectBlob   =   "reqMenu.frx":0000
End
Attribute VB_Name = "reqMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mCancel_Click()
    mainMenu.mCancel_Click
End Sub

Private Sub reqSubmit_Click()
    Dim xSht As Worksheet
    Dim xOutlookObj As Object
    Dim xEmailObj As Object
    Dim send_to As String
    On Error GoTo 0
    
    Set xOutlookObj = CreateObject("Outlook.Application")
    Set xEmailObj = xOutlookObj.CreateItem(olMailItem)
    Dim name As String
    Dim pw As String
    Dim mgr As String
    name = Me.TextBox1 & " " & Me.TextBox2
    pw = encryptPassword(Me.TextBox4)
    mgr = Me.TextBox3
    With xEmailObj
        .To = "jsikorski@helixelectric.com"
        .Subject = "Time Card User Request"
        .Body = name & vbNewLine & mgr & vbNewLine & pw
        .Send
    End With
    MsgBox "Please allow 1 business day for your request to be processed", vbInformation + vbOKOnly
    Me.Hide
'    ThisWorkbook.Close False
End Sub

Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .TextBox1.SetFocus
        .Caption = "REGISTER"
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        mainMenu.mCancel_Click
    End If
End Sub
