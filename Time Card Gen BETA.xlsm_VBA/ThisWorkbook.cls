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
    Call ExportVBA
    Call export_user_sheet
    For i = 1 To ThisWorkbook.Sheets.count
    ThisWorkbook.Unprotect "Dea504!!"
        If Worksheets(i).name <> "HOME" Then
            Worksheets(i).Visible = xlVeryHidden
        End If
    Next i
    ThisWorkbook.Protect "Dea504!!", True, False
    If user = "jsikorski" Then
        Application.DisplayAlerts = False
        ThisWorkbook.SaveAs Me.path & "\" & Me.name
    End If
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If Environ$("Username") <> "jsikorski" Then
        MsgBox "This file is locked!", 16, "FILE LOCKED"
        Cancel = True
    Else
        Call ExportVBA
    End If
End Sub

Private Sub Workbook_Open()
    openBook
End Sub
