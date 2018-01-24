VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} pjSuperMenu 
   Caption         =   "Superintendent Menu"
   ClientHeight    =   3255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5625
   OleObjectBlob   =   "pjSuperMenu.frx":0000
End
Attribute VB_Name = "pjSuperMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub smBuild_Click()
    Dim we As String
    Dim xlFile As String
    we = Format(week, "mm.dd.yy")
    xlFile = jobPath & "\" & jobNum & "\TimePackets\Week_" & we & "\" & jobNum & "_Week_" & we & ".xlsx"
    If testFileExist(xlFile) > 0 Then
        On Error Resume Next
        Dim ans As Integer
        ans = MsgBox("The packet already exists, Are you sure you want to overwrite it?", vbYesNo + vbQuestion)
        If ans = vbYes Then
            Kill xlFile
            Kill jobPath & "\" & jobNum & "\TimeSheets\Week_" & we & "\*.*"
        Else
            Exit Sub
        End If
        On Error GoTo 0
    End If
    sMenu.Hide
    Set lMenu = New pjSuperPkt
    lMenu.Show
End Sub

Private Sub smEdit_Click()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Dim we As String
    we = Format(week, "mm.dd.yy")
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Dim bk As Workbook
    Dim xlFile As String
    Dim aVal As Integer
    Dim bVal As Integer
    Dim i As Integer
    Dim tmp As Range
    i = 0
    xlFile = jobPath & "\" & jobNum & "\TimePackets\Week_" & we & "\" & jobNum & "_Week_" & we & ".xlsx"
    On Error GoTo 10
    Application.Workbooks.Open xlFile
    On Error GoTo 0
    Set bk = Workbooks(jobNum & "_Week_" & we & ".xlsx")
    bk.Worksheets("SAVE").Visible = xlSheetVisible
    For Each tmp In bk.Worksheets("Save").Range("A1", bk.Worksheets("SAVE").Range("A1").End(xlDown))
        If tmp.Value > aVal Then aVal = tmp.Value
        If tmp.Offset(0, 1).Value > bVal Then bVal = tmp.Offset(0, 1).Value
    Next tmp
    ReDim weekRoster(aVal, eCount)
    For Each tmp In bk.Worksheets("Save").Range("A1", bk.Worksheets("SAVE").Range("A1").End(xlDown))
        Dim xlEmp As Employee
        Set xlEmp = New Employee
        xlEmp.emClass = tmp.Offset(0, 2)
        xlEmp.elName = tmp.Offset(0, 3)
        xlEmp.efName = tmp.Offset(0, 4)
        xlEmp.emnum = tmp.Offset(0, 5)
        xlEmp.emPerDiem = tmp.Offset(0, 6)
        Set weekRoster(tmp.Offset(0, 0).Value, tmp.Offset(0, 1).Value) = xlEmp
    Next tmp
    bk.Worksheets("SAVE").Visible = xlVeryHidden
    wb.Activate
    bk.Close
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    sMenu.Hide
    Set lMenu = New pjSuperPkt
    lMenu.Show
    GoTo 20
10
    MsgBox ("Unable to Edit Packet - The file does not exist")
20
End Sub

Public Sub smExit_Click()
    sMenu.Hide
    mMenu.Show
End Sub

Private Sub smSubmut_Click()
    MsgBox ("This feature is not implemented yet")
End Sub

Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .Label2.Caption = job & vbNewLine & Format(week, "mm-dd-yy")
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Me.smExit_Click
    End If
End Sub

