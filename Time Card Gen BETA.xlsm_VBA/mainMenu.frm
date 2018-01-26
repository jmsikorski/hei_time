VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} mainMenu 
   Caption         =   "Select Job Number"
   ClientHeight    =   4440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8280.001
   OleObjectBlob   =   "mainMenu.frx":0000
End
Attribute VB_Name = "mainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()
    job = ComboBox1.Value
    Dim temp() As String
    temp = Split(job, " - ")
    jobNum = temp(0)
    jobName = temp(1)
    jobPath = ThisWorkbook.path & "\" & "Data.lnk"
    jobPath = Getlnkpath(jobPath)
End Sub

Public Sub mCancel_Click()
    Application.DisplayAlerts = False
    Dim unLockIn As String
    Dim ans As Integer, attempt As Integer
    Dim correct As Boolean
    correct = False
    attempt = 1
    ans = MsgBox("This file is locked" & vbNewLine & "Are you sure you want to quit?", 4147, "EXIT")
    If ans = 6 Then
        Application.DisplayAlerts = False
        ThisWorkbook.Close
    ElseIf ans = 2 Then
        If user = "jsikorski" Then
            On Error Resume Next
            If logMenu.Visible = True Then
                logMenu.Hide
            End If
            If mMenu.Visible = False Then
                mMenu.Show
            End If
            If sMenu.Visible = True Then
                sMenu.Hide
            End If
            End
        End If
        Do While correct = False And attempt > 0
            unLockIn = InputBox("This file is locked for editing" & vbNewLine & "Please enter the unlock password:", "UNLOCK FILE ATTEMPT " & attempt & "/3")
            If unLockIn = "" Then
                attempt = attempt + 1
            ElseIf unLockIn = "jms7481" Then
                On Error Resume Next
                If logMenu.Visible = True Then
                    logMenu.Hide
                End If
                If mMenu.Visible = True Then
                    mMenu.Hide
                End If
                If sMenu.Visible = True Then
                    sMenu.Hide
                End If
                If Application.WindowState = xlMinimized Then
                    Application.WindowState = xlMaximized
                End If
                On Error GoTo 0
                attempt = 0
                correct = True
            Else
                attempt = attempt + 1
            End If
            If attempt = 4 Then
                MsgBox "You have made 3 failed attempts!", 16, "FAILED UNLOCK"
                Application.DisplayAlerts = False
                ThisWorkbook.Close
            End If
        Loop
    End If

End Sub

Private Sub pjCoordinator_Click()
    MsgBox ("This feature is not implemented yet")
End Sub

Private Sub pjSuper_Click()
    If TypeName(mMenu) <> "mainMenu" Then
        job = "ERROR"
    Else
        If job = vbNullString Then
            MsgBox ("You must enter a job number")
            Exit Sub
        End If
    End If
    mMenu.Hide
    Set sMenu = New pjSuperMenu
    sMenu.Show
End Sub

Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
    Dim cJob As Range
    Dim uNum As Range
    For Each cJob In Worksheets("JOBS").Range("jobList")
        With Me.ComboBox1
        For Each uNum In Worksheets("USER").Range("A2", Worksheets("USER").Range("A2").End(xlDown))
            If uNum.Value = user Then
                If uNum.Offset(0, cJob.Row + 2) = True Then
                    .AddItem cJob.Value
                    .list(.ListCount - 1, 1) = cJob.Offset(0, 1).Value
                End If
            End If
        Next
      End With
    Next
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
       mMenu.mCancel_Click
    End If
End Sub

