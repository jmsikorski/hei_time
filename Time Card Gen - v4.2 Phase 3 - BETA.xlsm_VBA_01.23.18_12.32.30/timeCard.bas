Attribute VB_Name = "timeCard"
Public week As Date
Public job As String
Public user As String
Public lCnt As Integer
Public lNum As Integer
Public eList() As String
Public menuList() As Object
Public empRoster() As Employee
Public leadRoster() As Employee
Public jobPath As String
Public jobNum As String
Public jobName As String
Public weekRoster() As Employee
Public mMenu As mainMenu
Public sMenu As pjSuperMenu
Public lMenu As pjSuperPkt
Public tReview As teamReview
Public eCount As Integer
Private Const xPass = "Dea504!!"

Public Enum mType
    mainMenu = 1
    pjSuperMenu = 2
    pjSuperPkt = 3
    pjSuperPktEmp = 4
End Enum

Public Sub addMenu(mType As Integer)
    Dim tmp As Object
    Dim added As Boolean
    added = False
    Select Case mType
        Case 1
            Set tmp = New mainMenu
        Case 2
            Set tmp = New pjSuperMenu
        Case 3
            Set tmp = New pjSuperPkt
        Case 4
            Set tmp = New pjSuperPktEmp
        Case Default
            MsgBox ("ERROR: " & mType & " is not a valid menu")
    End Select
    For i = 0 To UBound(menuList)
        If menuList(i) Is Nothing Then
            Set menuList(i) = tmp
            added = True
            Exit For
        End If
    Next i
    If added = False Then
        ReDim Preserve menuList(UBound(menuList) + 1)
        Set menuList(UBound(menuList)) = tmp
    End If
End Sub

Private Sub copy_tables(ByRef wb As Workbook)
    Dim ws As Worksheet
    Set ws = wb.Worksheets("ROSTER")
    ws.Unprotect
    ws.ListObjects("Monday").DataBodyRange.Copy
    ws.Range("Tuesday").PasteSpecial xlPasteValues
    ws.Range("Wednesday").PasteSpecial xlPasteValues
    ws.Range("Thursday").PasteSpecial xlPasteValues
    ws.Range("Friday").PasteSpecial xlPasteValues
    ws.Range("Saturday").PasteSpecial xlPasteValues
    ws.Range("Sunday").PasteSpecial xlPasteValues
    ws.Activate
    ws.Protect
    Application.CutCopyMode = False
End Sub

Private Sub open_data_file(name As String)
    On Error GoTo share_err
    Dim xPath As String
    Dim xFile As String
    Dim wb As Workbook
    xPath = ThisWorkbook.path & "\" & "Data.lnk"
    xPath = Getlnkpath(xPath)
    xFile = xPath & "\" & name
    Workbooks.Open xFile
    Exit Sub
share_err:
    xPath = ThisWorkbook.path & "\" & "Data Files"
    xFile = xPath & "\" & name
    Workbooks.Open xFile
End Sub

Public Function Getlnkpath(ByVal Lnk As String) As String
   On Error Resume Next
   With CreateObject("Wscript.Shell").CreateShortcut(Lnk)
       Getlnkpath = .TargetPath
       .Close
   End With
End Function
Public Sub genLeadSheets()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Dim bks As Collection
    Set bks = New Collection
    Dim done As Boolean
    done = False
    Dim wb As Workbook
    Dim new_path() As String
    Set wb = ThisWorkbook
    ThisWorkbook.Unprotect xPass
    Dim xlPath As String
    Dim we As String
    we = Format(week, "mm.dd.yy")
    xlPath = jobPath & "\" & jobNum & "\TimeSheets\" & "Week_" & we & "\"
    On Error Resume Next
    new_path = Split(xlPath, "\")
    Dim i As Integer
    i = 0
    xlPath = vbNullString
    Do While new_path(i) <> jobNum
        xlPath = xlPath & new_path(i) & "\"
        i = i + 1
    Loop
    Do While i < UBound(new_path)
        xlPath = xlPath & new_path(i) & "\"
        i = i + 1
        MkDir xlPath
    Loop
    Dim e_cnt As Integer
    On Error GoTo 0
    Dim r_size As Integer
    For i = 0 To UBound(weekRoster)
        e_cnt = 1
        Dim iTemp As Employee
        Set iTemp = weekRoster(i, 0)
        Dim lsPath As String
        lsPath = iTemp.getLName & "_Week_" & we & ".xlsx"
        lsPath = xlPath + lsPath
        Dim ls As Workbook
        open_data_file "Lead Card - Office.xlsm"
        Set ls = Workbooks("Lead Card - Office.xlsm")
        SetAttr ls.path, vbNormal
        Application.DisplayAlerts = False
        Application.EnableEvents = False
        ls.SaveAs lsPath, 51
        Application.EnableEvents = True
        With ls.Worksheets("ROSTER").Range("Monday").Cells(1, 1)
            ls.Worksheets("ROSTER").Unprotect
            .Value = iTemp.getClass
            .Offset(0, 1).Value = iTemp.getFName & " " & iTemp.getLName
            .Offset(0, 2).Value = iTemp.getNum
            ls.Worksheets("ROSTER").Protect
        End With
        bks.Add ls
        For x = 1 To UBound(weekRoster, 2)
            Dim xTemp As Employee
            Set xTemp = weekRoster(i, x)
            If xTemp Is Nothing Then
            Else
                e_cnt = e_cnt + 1
                With ls.Worksheets("ROSTER").Range("Monday").Cells(x + 1, 1)
                    ls.Worksheets("ROSTER").Unprotect
                    .Value = iTemp.getClass
                    .Offset(0, 1).Value = xTemp.getFName & " " & xTemp.getLName
                    .Offset(0, 2).Value = xTemp.getNum
                    ls.Worksheets("ROSTER").Protect
                End With
            End If
        Next x
        ls.Worksheets("ROSTER").Unprotect
        For n = 1 To 7
            For p = e_cnt + 1 To 15
                ls.Worksheets("ROSTER").ListObjects(n).ListRows(e_cnt + 1).Delete
            Next p
        Next n
        ls.Worksheets("ROSTER").Protect
        copy_tables ls
    Next i
    If jobPath = vbNullString Then
        MsgBox ("ERROR!")
        Exit Sub
    End If

    For Each ls In bks
        ls.Save
        ls.Close
    Next ls
'    wb.Worksheets("LEAD").Visible = False
    ThisWorkbook.Protect xPass
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Public Sub showBooks()
Attribute showBooks.VB_ProcData.VB_Invoke_Func = "S\n14"
    If user = vbNullString Then
        user = Environ$("Username")
    End If
    If user = "jsikorski" Then
        ThisWorkbook.Unprotect xPass
        For i = 1 To ThisWorkbook.Sheets.count
            If Worksheets(i).Visible = xlVeryHidden Then
                Worksheets(i).Visible = True
            End If
        Next i
        Worksheets("KEY").Visible = xlVeryHidden
    Else
        MsgBox ("Sorry this toy is not for you to play with")
    End If
End Sub

Public Sub openBook()
    ThisWorkbook.Unprotect "Dea504!!"
    For i = 1 To ThisWorkbook.Sheets.count
        If Worksheets(i).name <> "HOME" Then
            Worksheets(i).Visible = xlVeryHidden
        End If
    Next i
    eCount = 15
    ReDim menuList(0)
    ReDim empRoster(0, 0)
    ReDim leadRoster(0, 0)
    ReDim weekRoster(0, eCount)
    Dim ld As Boolean 'True to load mainMenu false to skip
    ld = True
    lCnt = 1
    i = 0
    Dim rg As Range
    Dim auth As Boolean
    Dim attempt As Integer
    attempt = 0
    auth = False
    Dim uNum As Integer
    uNum = 2
    
    get_user_list
    auth = file_auth
    If auth = False Then
        ThisWorkbook.Close False
    End If
    
    
    For i = 1 To ThisWorkbook.Sheets.count
        If Worksheets(i).name <> "HOME" Then
            If Worksheets(i).name <> "KEY" Then
                Worksheets(i).Visible = xlHidden
                End If
        End If
    Next i

    week = calcWeek(Date)
    jobPath = vbNullString
    job = vbNullString
    Dim lst As Range
    Set lst = ThisWorkbook.Worksheets("Jobs").UsedRange
    lst.name = "jobList"
    Set lst = ThisWorkbook.Worksheets("ROSTER").UsedRange
    lst.name = "empList"
    Set mMenu = New mainMenu
    ThisWorkbook.Protect "Dea504!!", True, False
    If user <> "jsikorski" Then
        mMenu.Show
    ElseIf ld = True Then
        mMenu.Show
    End If
End Sub
 Public Sub hideBooks()
Attribute hideBooks.VB_ProcData.VB_Invoke_Func = "H\n14"
    For i = 1 To ThisWorkbook.Sheets.count
        If Worksheets(i).name <> "HOME" Then
            If Worksheets(i).name <> "KEY" Then
                Worksheets(i).Visible = False
                End If
        End If
    Next i
End Sub
Private Function file_auth() As Boolean
    Dim rg As Range
    Set rg = Worksheets("USER").Range("A" & 2)
    Dim logMenu As loginMenu
    Set logMenu = New loginMenu
    Dim pw As String
    Dim auth As Boolean
    aut = False
    user = Environ$("Username")
'    If user = "jsikorski" Then
'        file_auth = True
'    End If
    
    logMenu.TextBox2.Value = user
    logMenu.Show
    pw = logMenu.TextBox1.Value
    user = logMenu.TextBox2.Value
    Unload logMenu
    Do While rg.Offset(i, 0) <> vbNullString
        If user = rg.Offset(i, 0) Then
            auth = True
            uNum = i
        End If
        i = i + 1
    Loop
    If auth = False Then
        MsgBox ("YOU ARE NOT AUTHORIZED TO VIEW THIS FILE!")
        ThisWorkbook.Close
    End If
    
    pw = encryptPassword(pw)
    Do While testPW(rg.Offset(uNum, 1).Value, pw) = False
        If attempt < 3 Then
            Set logMenu = New loginMenu
            logMenu.TextBox2.Value = user
            logMenu.Show
            pw = logMenu.TextBox1.Value
            user = logMenu.TextBox2.Value
            Unload logMenu
            Do While rg.Offset(i, 0) <> vbNullString
                If user = rg.Offset(i, 0) Then
                    auth = True
                    uNum = i
                End If
                i = i + 1
            Loop
            attempt = attempt + 1
        Else
            MsgBox "You have made 3 failed attempts!", 16, "FAILED UNLOCK"
            If user <> "jsikorski" Then
                file_auth = False
            Else
                Exit Do
            End If
        End If
    Loop
    file_auth = True
End Function

Public Function saveWeekRoster(ByRef ws As Worksheet) As Integer
    ws.name = "SAVE"
    Dim cnt As Integer, x As Integer
    cnt = 0
    x = 0
    Dim done As Boolean
    Dim temp As Employee
    Set temp = New Employee
    With ws.Range("A1")
        For i = 0 To UBound(weekRoster)
            done = False
            Do While done = False
                If weekRoster(i, x) Is Nothing Then
                    done = True
                Else
                    .Offset(cnt, 0).Value = i
                    .Offset(cnt, 1).Value = x
                    .Offset(cnt, 2).Value = weekRoster(i, x).getClass
                    .Offset(cnt, 3).Value = weekRoster(i, x).getLName
                    .Offset(cnt, 4).Value = weekRoster(i, x).getFName
                    .Offset(cnt, 5).Value = weekRoster(i, x).getNum
                    .Offset(cnt, 6).Value = weekRoster(i, x).getPerDiem
                    cnt = cnt + 1
                End If
                x = x + 1
            Loop
            x = 0
        Next i
    End With

    saveWeekRoster = 1
End Function

Public Sub savePacket()
    Dim time As Date
    On Error Resume Next
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Dim bk As Workbook
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Dim xlPath As String
    Dim xlFile As String
    Dim we As String
    we = Format(week, "mm.dd.yy")
    xlPath = jobPath & "\" & jobNum & "\TimePackets\"
    MkDir xlPath
    xlPath = xlPath & "Week_" & we & "\"
    MkDir xlPath
    On Error GoTo 0
    xlFile = xlPath & jobNum & "_Week_" & we & ".xlsx"
    Set bk = Workbooks.Add
    saveWeekRoster bk.Sheets(1)
    If genRoster(bk, wb.Worksheets("ROSTER TEMPLATE")) = -1 Then
        MsgBox ("ERROR PRINTING ROSTER")
    End If
    moveRoster wb, bk
    bk.Worksheets("SAVE").Visible = xlVeryHidden
    If testFileExist(xlFile) = 1 Then
        Kill xlFile
    End If
    bk.SaveAs xlFile
    bk.Close
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

Public Function genRoster(ByRef wb As Workbook, ByRef ws As Worksheet) As Integer
    On Error GoTo 10
    Application.DisplayAlerts = False
    wb.Worksheets("SAVE").Activate
    Dim we As String
    Dim tmp As Range
    we = calcWeek(Date)
    Dim cnt As Integer
    cnt = 0
    With ws
        .Range("job_num").Value = jobNum
        .Range("job_name").Value = jobName
        .Range("week_ending").Value = we
        .Range("emp").Copy
        For Each tmp In wb.Worksheets("SAVE").Range("A1", wb.Worksheets("SAVE").Range("A1").End(xlDown))
            .Range("emp_count").Offset(cnt, 0).Value = cnt + 1
            .Range("emp_class").Offset(cnt, 0).Value = tmp.Offset(0, 2).Value
            .Range("emp_name").Offset(cnt, 0).Value = tmp.Offset(0, 4).Value & " " & tmp.Offset(0, 3).Value
            .Range("emp_num").Offset(cnt, 0).Value = tmp.Offset(0, 5).Value
            If (tmp.Offset(0, 6)) Then
                .Range("emp_phaseCode").Offset(cnt, 0).Value = "88070-08 Per Diem"
            Else
                .Range("emp_phaseCode").Offset(cnt, 0).Value = "N/A"
            End If
            .Range("emp").Offset(cnt, 0).PasteSpecial Paste:=xlPasteFormats
            cnt = cnt + 1
        Next tmp
        .Range("emp").Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range("emp").Borders(xlEdgeTop).Weight = xlThick
        
    End With
    Application.DisplayAlerts = True
    On Error GoTo 0
    genRoster = 1
    Exit Function
10
    genRoster = -1
    On Error GoTo 0
End Function

Public Sub moveRoster(wb As Workbook, bk As Workbook)
    wb.Unprotect xPass
    wb.Worksheets("ROSTER TEMPLATE").Visible = xlSheetVisible
    wb.Worksheets("ROSTER TEMPLATE").Copy after:=bk.Worksheets(bk.Sheets.count)
    wb.Worksheets("6-WEEK SCHEDULE").Visible = xlSheetVisible
    wb.Worksheets("6-WEEK SCHEDULE").Copy after:=bk.Worksheets(bk.Sheets.count)
    bk.Worksheets("ROSTER TEMPLATE").name = "ROSTER"
    With wb.Worksheets("ROSTER TEMPLATE").Range("emp")
        wb.Worksheets("ROSTER TEMPLATE").Range(.Offset(1, 0), .End(xlDown)).Clear
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThin
        .Value = vbNullString
    End With
    'CODE FOR CLEARING JOB INFO AND WE DATE
    changeNamedRange bk, "emp"
    changeNamedRange bk, "emp_class"
    changeNamedRange bk, "emp_comments"
    changeNamedRange bk, "emp_count"
    changeNamedRange bk, "emp_name"
    changeNamedRange bk, "emp_num"
    changeNamedRange bk, "emp_perdiem"
    changeNamedRange bk, "emp_phaseCode"
    changeNamedRange bk, "job_name"
    changeNamedRange bk, "job_num"
    changeNamedRange bk, "week_ending"
    With bk.Worksheets("ROSTER")
        .Range("job_num") = jobNum
        .Range("job_name") = jobName
        .Range("week_ending") = week
    End With
    wb.Worksheets("6-WEEK SCHEDULE").Visible = xlSheetHidden
    wb.Worksheets("ROSTER TEMPLATE").Visible = xlSheetHidden
    wb.Protect xPass

End Sub

Private Sub changeNamedRange(wb As Workbook, rng As String)
    Dim nr As name
    Set nr = wb.Names.Item(rng)
    Select Case rng
        Case "emp"
            nr.RefersTo = "=ROSTER!$A$9:$G$9"
        Case "emp_class"
            nr.RefersTo = "=ROSTER!$B$9"
        Case "emp_comments"
            nr.RefersTo = "=ROSTER!$G$9"
        Case "emp_count"
            nr.RefersTo = "=ROSTER!$A$9"
        Case "emp_name"
            nr.RefersTo = "=ROSTER!$C$9"
        Case "emp_num"
            nr.RefersTo = "=ROSTER!$D$9"
        Case "emp_perdiem"
            nr.RefersTo = "=ROSTER!$F$9"
        Case "emp_phaseCode"
            nr.RefersTo = "=ROSTER!$E$9"
        Case "job_name"
            nr.RefersTo = "=ROSTER!$E$1"
        Case "job_num"
            nr.RefersTo = "=ROSTER!$E$2"
        Case "week_ending"
            nr.RefersTo = "=ROSTER!$E$4"
        Case Else
            MsgBox ("Invalid Range")
    End Select
End Sub

Public Sub testCNR()
    Dim wb As Workbook
    Set wb = Workbooks("46XXXX _Week_12.10.17")
    changeNamedRange wb, "emp"
    changeNamedRange wb, "emp_class"
    changeNamedRange wb, "emp_comments"
    changeNamedRange wb, "emp_count"
    changeNamedRange wb, "emp_name"
    changeNamedRange wb, "emp_num"
    changeNamedRange wb, "emp_perdiem"
    changeNamedRange wb, "emp_phaseCode"
    changeNamedRange wb, "job_name"
    changeNamedRange wb, "job_num"
    changeNamedRange wb, "week_ending"
End Sub
Public Sub printRoster()
    Dim temp As Employee
    For i = 0 To UBound(weekRoster)
        For x = 0 To UBound(weekRoster, 2)
            If weekRoster(i, x) Is Nothing Then
            Else
                Set temp = weekRoster(i, x)
                MsgBox ("LD: " & i & vbNewLine & "EMP: " & x & _
                vbNewLine & temp.getFName & " " & temp.getLName)
            End If
        Next x
    Next i
            
End Sub

Public Function isSave() As Integer
    Application.ScreenUpdating = False
    
    Dim xlFile As String
    Dim we As String
    Dim tmp() As String
    we = Format(week, "mm.dd.yy")
    xlFile = jobPath & "\" & jobNum & "\TimePackets\Week_" & we & "\" & jobNum & "_Week_" & we & ".xlsx"
    If testFileExist(xlFile) > 0 Then
        isSave = 1
    Else
        isSave = -1
    End If
End Function

Public Function testFileExist(FilePath As String) As Integer

    Dim TestStr As String
    TestStr = ""
    On Error Resume Next
    TestStr = Dir(FilePath)
    On Error GoTo 0
    If TestStr = "" Then
        testFileExist = -1
    Else
        testFileExist = 1
    End If

End Function

Public Sub resizeRoster(l As Integer, e As Integer)
    
    Dim newRoster() As Employee
    ReDim newRoster(l, e)
    Dim temp As Employee
    For i = 0 To l
        For x = 0 To e
            On Error Resume Next
            Set temp = weekRoster(i, x)
'            If temp Is Nothing Then
'            Else
                Set newRoster(i, x) = temp
'            End If
        Next x
    Next i
    On Error GoTo 0
    ReDim weekRoster(l, e)
    For i = 0 To l
        For x = 0 To e
            Set weekRoster(i, x) = newRoster(i, x)
        Next x
    Next i
    
    
End Sub

Public Sub insertRoster(index As Integer)
    Dim x As Integer
    Dim tmp As Employee
    Dim tmpRoster() As Employee
    ReDim tmpRoster(UBound(weekRoster), eCount)
    For x = 0 To index - 1
        For i = 0 To eCount
            Set tmp = weekRoster(x, i)
            If tmp Is Nothing Then
            Else
                Set tmpRoster(x, i) = tmp
            End If
        Next i
    Next x
    For x = index + 1 To UBound(weekRoster)
        For i = 0 To eCount
            Set tmp = weekRoster(x - 1, i)
            If tmp Is Nothing Then
            Else
                Set tmpRoster(x, i) = tmp
            End If
        Next i
    Next x
    For x = 0 To UBound(weekRoster)
        For i = 0 To eCount
            Set weekRoster(x, i) = tmpRoster(x, i)
        Next i
    Next x
End Sub

Private Sub loadMenu() 'ws As Worksheet)
    Dim ws As Worksheet
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Set ws = Workbooks("46XXXX _Week_12.10.17").Worksheets("SAVE")
    Dim rng As Range
    Set rng = ws.Range("A1")
    Dim cnt As Integer
    cnt = 0
    cnt = rng.End(xlDown).Value
    ReDim weekRoster(cnt, 15)
    cnt = 0
    For Each rng In ws.Range(rng, rng.End(xlDown))
        Dim tmp As Employee
        Set tmp = New Employee
        tmp.efName = rng.Offset(0, 4).Value
        tmp.elName = rng.Offset(0, 3).Value
        tmp.emnum = rng.Offset(0, 5).Value
        tmp.emClass = rng.Offset(0, 2).Value
        tmp.emPerDiem = rng.Offset(0, 6).Value
        Set weekRoster(rng.Offset(0, 0).Value, rng.Offset(0, 1).Value) = tmp
        cnt = cnt + 1
    Next rng
    wb.Activate
End Sub

Public Sub testPacket()
Attribute testPacket.VB_ProcData.VB_Invoke_Func = "r\n14"
    loadMenu
    savePacket
    MsgBox ("Complete")
End Sub
