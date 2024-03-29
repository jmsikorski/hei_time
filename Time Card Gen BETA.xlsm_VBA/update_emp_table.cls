VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "update_emp_table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Enum state
    open_phase = 1
    close_phase = 2
    update_phase = 3
End Enum

Private Const pw = ""
    
Public Sub update_emp_table()
Attribute update_emp_table.VB_Description = "Update Open phase codes from Labor Report\n"
Attribute update_emp_table.VB_ProcData.VB_Invoke_Func = "U\n14"
    Application.ScreenUpdating = False
    'on error goto 10
    Dim new_emp As Range
    Dim rng As Range
    Dim ws As Worksheet
    Dim cnt As Integer
    cnt = 1
    Set ws = ThisWorkbook.Worksheets("Roster")
    ws.Unprotect pw
    ws.ListObjects("emp_roster").DataBodyRange.Clear
1:
    Set new_emp = get_emp(cnt)
    If new_emp Is Nothing Then
        cnt = cnt + 1
        GoTo 1
    ElseIf new_emp.Cells(1, 2) = vbNullString Then
        GoTo update_done
    End If
    For Each rng In ws.ListObjects("emp_roster").ListColumns(1).DataBodyRange
        If rng.Value = vbNullString Then
            GoTo 5
        End If
        If rng.Value = new_emp.Cells(1, 1) Then
            cnt = cnt + 1
            GoTo 1
        End If
    Next rng
    ws.ListObjects("emp_roster").ListRows.Add
5:
    If insert_emp(new_emp) = -1 Then
        GoTo 20
    End If
    cnt = cnt + 1
    GoTo 1
update_done:
    On Error GoTo 0
    Set rng = ws.Range(ws.ListObjects("emp_roster").Range(1, 1), ws.ListObjects("emp_roster").Range(1, ws.ListObjects("emp_roster").ListColumns.count).End(xlDown))
    ws.ListObjects("emp_roster").Resize rng
    Workbooks("Attendance Tracking.xlsx").Close
    With ThisWorkbook.Worksheets("ROSTER")
        .Unprotect
        .Range("emp_table_updated") = Now()
        .Protect
    End With
    Application.ScreenUpdating = True
    ws.Protect pw
    Exit Sub
10:
    Dim ans As Integer
    Dim xFile As String
    With Application.FileDialog(msoFileDialogOpen)
        .title = "Find Attendance Tracking Roster"
        .Filters.Add "Excel Files", "*.xls*", 1
        .InitialFileName = ThisWorkbook.path & "\"
        ans = .Show
        If ans = 0 Then
            Exit Sub
        Else
            Workbooks.Open .SelectedItems(1)
            xFile = .SelectedItems(1)
            SetAttr xFile, vbNormal
        End If
    End With
    Set mb = Workbooks(xFile)
    Resume Next
    Exit Sub
20:
    MsgBox "ERROR: Unable to update roster", vbCritical, "ERROR!"
    On Error GoTo 0
    ws.Protect pw
    Application.ScreenUpdating = True
End Sub

Private Function get_emp(Optional cnt As Integer = 1) As Range
    Dim new_emp As Range
    Dim ans As Integer
    Dim datFile As String
    datFile = "Attendance Tracking.xlsx"
1:
    Dim mb As Workbook
    Dim xlFile As String
    On Error GoTo 30
    Set mb = Workbooks(datFile)
    On Error GoTo 0
    Dim rng As Range
    Set rng = mb.Worksheets(1).ListObjects("emp_roster").ListRows(cnt).Range
    If rng.Interior.Color = 255 Then
        get_emp = Nothing
        Exit Function
    End If
    Set new_emp = mb.Worksheets(1).ListObjects("emp_roster").ListRows(cnt).Range
    Set get_emp = new_emp
    Exit Function
30:
    If timeCard.testFileExist(ThisWorkbook.path & "\Attendance Tracking.xlsx") <> -1 Then
        xlFile = ThisWorkbook.path & "\Attendance Tracking.xlsx"
    Else
        GoTo 40
    End If
    Set mb = Workbooks(datFile)
    Resume Next
40:
    timeCard.open_data_file datFile
    Set mb = Workbooks(datFile)
    Resume Next
End Function


Private Function insert_emp(emp As Range) As Integer
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("Roster")
    
    For Each rng In ws.ListObjects("emp_roster").ListColumns(1).DataBodyRange 'ws.Range("A2", ws.Range("A1").End(xlDown))
        If rng.Value = emp.Cells(1, 1) Then
            GoTo 10
        End If
        If rng.Value = vbNullString Then
            GoTo 1
        Else
            GoTo 5
        End If
1:
        Set rng = Range(rng, rng.Offset(0, emp.Columns.count - 1))
        rng = emp.Value
        With rng
            For i = 1 To .Columns.count
                With .Cells(1, i)
                    .Font.name = "Helvetica"
                    .Font.Bold = False
                    .Font.Size = 12
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).Weight = xlThin
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Borders(xlEdgeTop).Weight = xlThin
                    .Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Borders(xlEdgeLeft).Weight = xlThin
                    .Borders(xlEdgeRight).LineStyle = xlContinuous
                    .Borders(xlEdgeRight).Weight = xlThin
                End With
            Next i
        End With
        Exit For
5:
    Next rng
    insert_code = 1
    Exit Function
10:
    insert_code = -1
    On Error GoTo 0
End Function


Private Function resize_name_range(name As String, ws As Worksheet, c1 As Range, c2 As Range) As Integer
    'on error goto 10
    Dim wb As Workbook
    Dim nr As name
    Dim rng As Range
    Set wb = ThisWorkbook
    Set nr = wb.Names.Item(name)
    Set rng = ws.Range(c1, c2)
    nr.RefersTo = rng
    resize_name_range = 1
    Exit Function
10:
    resize_name_range = -1
    On Error GoTo 0
End Function
