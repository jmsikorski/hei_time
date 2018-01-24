Attribute VB_Name = "user_form"
Public Sub get_user_list()
    Dim auth As String
    Dim user As String
    On Error GoTo err_tag
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Dim URL As String
    Dim qt As QueryTable
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("USER")
    
    URL = "https://github.com/jmsikorski/hei_misc/blob/master/Modules/Time_Card_User.csv"
    Set qt = ws.QueryTables.Add( _
        Connection:="URL;" & URL, _
        Destination:=ws.Range("A1"))
     
    With qt
        .RefreshOnFileOpen = False
        .name = "Users"
        .FieldNames = True
        .WebSelectionType = xlAllTables
        .Refresh
    End With
    ws.Range("A1").EntireColumn.Delete
    ws.Range("D1:F1").EntireColumn.Clear
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub
err_tag:
    MsgBox "ERROR LOADING LICENSE!", vbCritical + vbOKOnly
    With ws.UsedRange
        .Offset(1, 0).Clear
        .Value = "X"
    End With
    ThisWorkbook.Close False
 End Sub

Public Sub extract_users()
Attribute extract_users.VB_ProcData.VB_Invoke_Func = "E\n14"
    If Environ$("Username") = "jsikorski" Then
        Application.DisplayAlerts = False
        Application.ScreenUpdating = False
        showBooks
        Dim wb As Workbook
        Dim dwb As Workbook
        Set wb = ThisWorkbook
        Set dwb = Workbooks.Add
        Debug.Print wb.name
        Debug.Print dwb.name
        wb.Worksheets("USER").Copy after:=dwb.Worksheets(1)
        dwb.Worksheets(1).Delete
        dwb.SaveAs "C:\Users\jsikorski\Documents\GitHub\hei_misc\Modules\Time_Card_User.csv", xlCSV
        dwb.Close
        hideBooks
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
    End If
End Sub
