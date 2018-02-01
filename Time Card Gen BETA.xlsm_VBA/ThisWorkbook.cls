VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    Application.Visible = False
    If Environ$("Username") = "jsikorski" Then
        Call ExportVBA
    End If
    For i = 1 To ThisWorkbook.Sheets.count
    ThisWorkbook.Unprotect xPass
        If Worksheets(i).name <> "HOME" Then
            Worksheets(i).Visible = xlVeryHidden
        End If
    Next i
    ThisWorkbook.Protect xPass, True, False
    Application.DisplayAlerts = False
    ThisWorkbook.SaveAs Me.path & "\" & Me.name
    Application.DisplayFullScreen = False
    Application.WindowState = xlNormal
    hiddenApp.Quit
    Set hiddenApp = Nothing
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error Resume Next
    If Environ$("Username") = "jsikorski" Then
        Call ExportVBA
    End If
    On Error GoTo 0
End Sub

Private Sub Workbook_Open()
    main
End Sub

