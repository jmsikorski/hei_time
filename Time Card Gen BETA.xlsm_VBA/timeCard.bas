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
Public Const eCount = 15
Public xPass As String

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
    Set ws = wb.Worksheets("LEAD")
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

Public Sub open_data_file(name As String, Optional pw As String)
    On Error GoTo share_err
    Dim xPath As String
    Dim xFile As String
    Dim wb As Workbook
    xPath = ThisWorkbook.path & "\" & "Data.lnk"
    xPath = Getlnkpath(xPath)
    xFile = xPath & "\" & name
    If pw = vbNullString Then
        Workbooks.Open xFile
    Else
        Workbooks.Open xFile, Password:=pw
    End If
    Exit Sub
share_err:
    xPath = ThisWorkbook.path & "\" & "Data Files"
    xFile = xPath & "\" & name
    If pw = vbNullString Then
        Workbooks.Open xFile
    Else
        Workbooks.Open xFile, Password:=pw
    End If
End Sub

Public Function Getlnkpath(ByVal Lnk As String) As String
   On Error Resume Next
   With CreateObject("Wscript.Shell").CreateShortcut(Lnk)
       Getlnkpath = .TargetPath
       .Close
   End With
End Function

Private Function getLeadSheets(xStrPath As String) As String
'UpdateByExtendoffice20160623
    Dim xFile As String
    On Error Resume Next
    If xStrPath = "" Then
        getLeadSheets = "-1"
        Exit Function
    End If
    xFile = Dir(xStrPath & "\*.xlsx")
    Do While xFile <> ""
        getLeadSheets = getLeadSheets & xFile & ","
        xFile = Dir
    Loop
    getLeadSheets = Left(getLeadSheets, Len(getLeadSheets) - 1)

End Function

Public Function loadShifts(Optional test As Boolean) As Integer
    On Error GoTo shift_err
    Dim wb_arr() As String
    Dim lead_arr As String
    Dim xlPath As String
    Dim hiddenApp As New Excel.Application
    If test Then
        jobNum = "461705"
        week = calcWeek(Date)
        we = Format(week, "mm.dd.yy")
    End If
    xlPath = timeCard.Getlnkpath(ThisWorkbook.path & "\Data.lnk") & "\" & jobNum & "\Week_" & we & "\TimeSheets\"
    lead_arr = getLeadSheets(xlPath)
    wb_arr = Split(lead_arr, ",")
    For i = 0 To UBound(wb_arr)
        xlFile = xlPath & wb_arr(i)
        hiddenApp.Workbooks.Open xlFile
    Next
    Dim n As Integer
    Dim rng As Range
    Dim tRng As Range
    n = 0
    For l = 0 To UBound(weekRoster)
        For e = 0 To UBound(weekRoster, 2)
            n = 0
            If weekRoster(l, e) Is Nothing Then
                Exit For
            End If
            Do While Left(wb_arr(n), Len(wb_arr(n)) - 19) <> weekRoster(l, 0).getLName
                n = n + 1
            Loop
            Set rng = hiddenApp.Workbooks(wb_arr(n)).Worksheets("DATA").Range("D1", hiddenApp.Workbooks(wb_arr(n)).Worksheets("DATA").Range("D1").End(xlDown))
            For Each tRng In rng
                If tRng.Value = weekRoster(l, e).getNum Then
                    Dim shft As shift
                    Set shft = New shift
                    shft.setDay = tRng.Offset(0, -3)
                    shft.setHrs = tRng.Offset(0, 1)
                    shft.setPhase = Val(Left(tRng.Offset(0, 2), 5))
                    weekRoster(l, e).addShift shft
                End If
            Next tRng
        Next e
        n = 0
    Next l
    For wb = 0 To UBound(wb_arr)
        hiddenApp.Workbooks(wb_arr(wb)).Close False
    Next
    hiddenApp.Quit
    Set hiddenApp = Nothing
    loadShifts = 1
    Exit Function
shift_err:
    loadShifts = -1
    For wb = 0 To UBound(wb_arr)
        hiddenApp.Workbooks(wb_arr(wb)).Close False
    Next
    hiddenApp.Quit
    Set hiddenApp = Nothing
    
End Function

Public Sub test_gen()
    timeCard.genTimeCard True
End Sub
Sub showsave()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("SAVE")
    ws.Visible = True
    Set ws = ActiveWorkbook.Worksheets("DATA")
    ws.Visible = True
    Set ws = ActiveWorkbook.Worksheets("ROSTER")
    ws.Visible = True
    Stop
    ws.Visible = False


End Sub
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
    xlPath = jobPath & "\" & jobNum & "\Week_" & we & "\TimeSheets\"
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
    Dim bk As Workbook
    Set bk = Workbooks(jobNum & "_Week_" & we & ".xlsx")
    For i = 0 To UBound(weekRoster)
        e_cnt = 1
        Dim iTemp As Employee
        Set iTemp = weekRoster(i, 0)
        Dim lsPath As String
        Dim ls As Workbook
        lsPath = iTemp.getLName & "_Week_" & we & ".xlsx"
        lsPath = xlPath + lsPath
        open_data_file "Lead Card - Office.xlsm"
        Set ls = Workbooks("Lead Card - Office.xlsm")
        SetAttr ls.path, vbNormal
        Application.DisplayAlerts = False
        Application.EnableEvents = False
        ls.SaveAs lsPath, 51
        Application.EnableEvents = True
        With ls.Worksheets("LEAD").Range("Monday").Cells(1, 1)
            ls.Worksheets("LEAD").Unprotect
            .Value = iTemp.getClass
            .Offset(0, 1).Value = iTemp.getFName & " " & iTemp.getLName
            .Offset(0, 2).Value = iTemp.getNum
            ls.Worksheets("LEAD").Protect
        End With
        bks.Add ls
        For x = 1 To UBound(weekRoster, 2)
            Dim xTemp As Employee
            Set xTemp = weekRoster(i, x)
            If xTemp Is Nothing Then
            Else
                e_cnt = e_cnt + 1
                With ls.Worksheets("LEAD").Range("Monday").Cells(x + 1, 1)
                    ls.Worksheets("LEAD").Unprotect
                    .Value = iTemp.getClass
                    .Offset(0, 1).Value = xTemp.getFName & " " & xTemp.getLName
                    .Offset(0, 2).Value = xTemp.getNum
                    ls.Worksheets("LEAD").Protect
                End With
            End If
        Next x
        ls.Worksheets("LEAD").Unprotect
        For n = 1 To 7
            For p = e_cnt + 1 To 15
                ls.Worksheets("LEAD").ListObjects(n).ListRows(e_cnt + 1).Delete
            Next p
        Next n
        ls.Worksheets("LEAD").Protect
        copy_tables ls
        If genRoster(bk, ls.Worksheets("ROSTER"), i + 1) = -1 Then
            MsgBox ("ERROR PRINTING ROSTER")
        End If
        bk.Worksheets("SAVE").Visible = xlVeryHidden
        ls.Worksheets("ROSTER").Visible = xlVeryHidden
        ls.Worksheets("DATA").Visible = xlVeryHidden
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
    Next i
    If jobPath = vbNullString Then
        MsgBox ("ERROR!")
        Exit Sub
    End If

    For Each ls In bks
        ls.Save
        ls.Close
    Next ls
    bk.Close False
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
            If ThisWorkbook.Worksheets(i).Visible = xlVeryHidden Then
                ThisWorkbook.Worksheets(i).Visible = True
            End If
        Next i
        ThisWorkbook.Worksheets("KEY").Visible = xlVeryHidden
    Else
        MsgBox ("Sorry this toy is not for you to play with")
    End If
End Sub

Public Sub openBook()
    xPass = encryptPassword("A~™þ»›Ûæ")
    ThisWorkbook.Unprotect xPass
    For i = 1 To ThisWorkbook.Sheets.count
        If Worksheets(i).name <> "HOME" Then
            Worksheets(i).Visible = xlVeryHidden
        End If
    Next i
    ReDim menuList(0)
    ReDim empRoster(0, 0)
    ReDim leadRoster(0, 0)
    ReDim weekRoster(0, eCount)
    Dim ld As Boolean 'True to load mainMenu false to skip
    ld = True
    lCnt = 1
    i = 0
    Dim rg As Range
    Dim auth As Integer
    Dim attempt As Integer
    attempt = 0
    auth = 0
    Dim uNum As Integer
    uNum = 2
    Debug.Print DateDiff("m", ThisWorkbook.Worksheets("USER").Range("user_updated").Value, Now())
    If DateDiff("n", ThisWorkbook.Worksheets("USER").Range("user_updated").Value, Now()) > 0 Then
        get_user_list
        ThisWorkbook.Worksheets("USER").Range("user_updated").Value = Now()
    End If
auth_retry:
    auth = file_auth
    If auth = -1 Then
        Dim ans As Integer
        ans = MsgBox("This program is not licensed!", vbCritical + vbAbortRetryIgnore)
        If ans = vbIgnore Then
            ThisWorkbook.Close False
        ElseIf ans = vbRetry Then
            GoTo auth_retry
        ElseIf ans = vbAbort Then
            If Environ$("username") = "jsikorski" Then
                Exit Sub
            Else
                ThisWorkbook.Close False
            End If
        Else
            ThisWorkbook.Close , False
        End If
    ElseIf auth = -2 Then
        ThisWorkbook.Close
    ElseIf auth = -3 Then
        MsgBox "YOU ARE NOT AUTHORIZED TO VIEW THIS FILE!", vbCritical + vbOKOnly, "EXIT!"
        ThisWorkbook.Close
    End If
    
    
    For i = 1 To ThisWorkbook.Sheets.count
        If Worksheets(i).name <> "HOME" Then
            If Worksheets(i).name <> "KEY" Then
                Worksheets(i).Visible = xlHidden
                End If
        End If
    Next i
    If DateDiff("h", ThisWorkbook.Worksheets("ROSTER").Range("emp_table_updated").Value, Now()) > 1 Then
        update_emp_table.update_emp_table
    End If
    week = calcWeek(Date)
    jobPath = vbNullString
    job = vbNullString
    Dim lst As Range
    Set lst = ThisWorkbook.Worksheets("Jobs").UsedRange
    lst.name = "jobList"
    Set lst = ThisWorkbook.Worksheets("ROSTER").UsedRange
    lst.name = "empList"
    Set mMenu = New mainMenu
    ThisWorkbook.Protect xPass, True, False
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

Private Function get_lic(url As String) As Boolean
    
    get_lic = False
    Dim winhttp As New WinHttpRequest
    winhttp.Open "get", url, False
    winhttp.Send
    If winhttp.responseText = "True" Then get_lic = True
End Function

Private Sub show_key()
    ThisWorkbook.Worksheets("KEY").Visible = True
End Sub
Private Sub hide_key()
    ThisWorkbook.Worksheets("KEY").Visible = False
End Sub

Public Function publicEncryptPassword(pw As String) As String
    If Environ$("username") <> "jsikorski" Then
        If InputBox("Authorization code:", "RESTRICED") <> 12292018 Then
            publicEncryptPassword = "ERROR"
            Exit Function
        End If
    End If
    Dim pwi As Long
    Dim test As String
    Dim epw As String
    Dim key As Long
    epw = vbnullStrig
    For i = 0 To Len(pw) - 1
        test = Left(pw, 1)
        pwi = Asc(test)
        pw = Right(pw, Len(pw) - 1)
        key = ThisWorkbook.Worksheets("KEY").Range("A" & i + 1).Value
        If key = pwi Then key = key + 128
        pwi = pwi Xor key
        If pwi = key + 128 Then
            pwi = key
        End If
        epw = epw & Chr(pwi)
    Next i
    publicEncryptPassword = epw
End Function

Private Function encryptPassword(pw As String) As String
    Dim pwi As Long
    Dim test As String
    Dim epw As String
    Dim key As Long
    epw = vbnullStrig
    For i = 0 To Len(pw) - 1
        test = Left(pw, 1)
        pwi = Asc(test)
        pw = Right(pw, Len(pw) - 1)
        key = ThisWorkbook.Worksheets("KEY").Range("A" & i + 1).Value
        If key = pwi Then key = key + 128
        pwi = pwi Xor key
        If pwi = key + 128 Then
            pwi = key
        End If
        epw = epw & Chr(pwi)
    Next i
    encryptPassword = epw
End Function


Private Function file_auth() As Integer
    Dim rg As Range
    Set rg = Worksheets("USER").Range("A" & 2)
    Dim logMenu As loginMenu
    Dim pw As String
    Dim auth As Integer
login_retry:
    Set logMenu = New loginMenu
    auth = 0
    If get_lic("https://raw.githubusercontent.com/jmsikorski/hei_misc/master/Licence.txt") Then
        user = Environ$("Username")
'        If user = "jsikorski" Then
'            file_auth = True
'            Exit Function
'        End If
        logMenu.TextBox2.Value = user
        logMenu.Show
        pw = logMenu.TextBox1.Value
        user = logMenu.TextBox2.Value
        Unload logMenu
        Do While rg.Offset(i, 0) <> vbNullString
            If user = rg.Offset(i, 0) Then
                If rg.Offset(i, 2) = "YES" Then
                    auth = 1
                    uNum = i
                    Exit Do
                Else
                    If MsgBox("User pending authorization", vbwarning + vbRetryCancel, "INVALID USERNAME") = vbRetry Then
                        GoTo login_retry
                    Else
                        file_auth = -2
                        Exit Function
                    End If
                End If
            End If
            i = i + 1
        Loop
        If auth = False Then
            file_auth = -3
            Exit Function
        End If
        Do While encryptPassword(rg.Offset(uNum, 1).Value) <> pw
            If attempt < 2 Then
                attempt = attempt + 1
                Dim pw_ans As Integer
                pw_ans = MsgBox("Invalid Password" & vbNewLine & "Attempt " & attempt & " of 3", vbExclamation + vbRetryCancel, "ERROR")
                If pw_ans = vbCancel Then
                    file_auth = -3
                End If
                Set logMenu = New loginMenu
                logMenu.TextBox2.Value = user
                logMenu.Show
                pw = logMenu.TextBox1.Value
                user = logMenu.TextBox2.Value
                Unload logMenu
                Do While rg.Offset(i, 0) <> vbNullString
                    If user = rg.Offset(i, 0) Then
                        auth = 1
                        uNum = i
                    End If
                    i = i + 1
                Loop
            Else
                MsgBox "You have made 3 failed attempts!", 16, "FAILED UNLOCK"
                If user <> "jsikorski" Then
                    ThisWorkbook.Close False
                Else
                    Exit Do
                End If
            End If
        Loop
        file_auth = 1
    Else
        file_auth = -1
    End If
End Function

Public Function saveWeekRoster(ByRef ws As Worksheet) As Integer
    ws.name = "SAVE"
    Dim cnt As Integer, x As Integer
    cnt = 0
    x = 0
    Dim done As Boolean
    Dim tEmp As Employee
    Set tEmp = New Employee
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
    xlPath = jobPath & "\" & jobNum & "\Week_" & we
    MkDir xlPath
    xlPath = xlPath & "\TimePackets\"
    MkDir xlPath
    On Error GoTo 0
    xlFile = xlPath & jobNum & "_Week_" & we & ".xlsx"
    open_data_file "Packet Template.xlsx"
    Set bk = Workbooks("Packet Template.xlsx")
    saveWeekRoster bk.Sheets("SAVE")
    If genRoster(bk, bk.Worksheets("ROSTER")) = -1 Then
        MsgBox ("ERROR PRINTING ROSTER")
    End If
'    moveRoster wb, bk
    bk.Worksheets("SAVE").Visible = xlVeryHidden
    If testFileExist(xlFile) = 1 Then
        Kill xlFile
    End If
    bk.SaveAs xlFile
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

Public Function genRoster(ByRef wb As Workbook, ByRef ws As Worksheet, Optional lead As Integer) As Integer
    On Error GoTo 10
    Application.DisplayAlerts = False
    wb.Worksheets("SAVE").Activate
    Dim we As String
    Dim tmp As Range
    we = calcWeek(Date)
    Dim cnt As Integer
    cnt = 0
    If lead = 0 Then
        With ws
            .Range("job_num").Value = jobNum
            .Range("job_name").Value = jobName
            .Range("week_ending").Value = we
            .Range("emp").Offset(1, 0).Copy
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
                If cnt > 1 Then
                    .Range("emp").Offset(cnt, 0).PasteSpecial Paste:=xlPasteFormats
                End If
                cnt = cnt + 1
            Next tmp
            .Range("emp").Borders(xlEdgeTop).LineStyle = xlContinuous
            .Range("emp").Borders(xlEdgeTop).Weight = xlThick
            
        End With
    Else
        With ws
            .Range("job_num").Value = jobNum
            .Range("job_name").Value = jobName
            .Range("week_ending").Value = we
            .Range("emp").Copy
            ws.Activate
            For Each tmp In wb.Worksheets("SAVE").Range("A1", wb.Worksheets("SAVE").Range("A1").End(xlDown))
                If tmp.Value = lead - 1 Then
                    .Range("emp_count").Offset(cnt, 0).Value = cnt + 1
                    .Range("emp_class").Offset(cnt, 0).Value = tmp.Offset(0, 2).Value
                    .Range("emp_name").Offset(cnt, 0).Value = tmp.Offset(0, 4).Value & " " & tmp.Offset(0, 3).Value
                    .Range("emp_num").Offset(cnt, 0).Value = tmp.Offset(0, 5).Value
                    If (tmp.Offset(0, 6)) Then
                        .Range("emp_phaseCode").Offset(cnt, 0).Value = "88070-08 Per Diem"
                    Else
                        .Range("emp_phaseCode").Offset(cnt, 0).Value = "N/A"
                    End If
                    If cnt > 1 Then
                        .Range("emp").Offset(cnt, 0).PasteSpecial Paste:=xlPasteFormats
                    End If
                    cnt = cnt + 1
                ElseIf tmp.Value > lead Then
                    Exit For
                End If
            Next tmp
'            .Range("emp").Borders(xlEdgeTop).LineStyle = xlContinuous
'            .Range("emp").Borders(xlEdgeTop).Weight = xlThick
        End With
    End If
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
    Dim tEmp As Employee
    For i = 0 To UBound(weekRoster)
        For x = 0 To UBound(weekRoster, 2)
            If weekRoster(i, x) Is Nothing Then
            Else
                Set tEmp = weekRoster(i, x)
                MsgBox ("LD: " & i & vbNewLine & "EMP: " & x & _
                vbNewLine & tEmp.getFName & " " & tEmp.getLName)
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
    xlFile = jobPath & "\" & jobNum & "\Week_" & we & "\TimePackets\" & jobNum & "_Week_" & we & ".xlsx"
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
    Dim tEmp As Employee
    For i = 0 To l
        For x = 0 To e
            On Error Resume Next
            Set tEmp = weekRoster(i, x)
'            If temp Is Nothing Then
'            Else
                Set newRoster(i, x) = tEmp
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

Public Sub genTimeCard(Optional test As Boolean)
    Dim hiddenApp As New Excel.Application
    Dim xlPath As String
    Dim xlFile As String
    If test Then
        jobNum = "461705"
        week = calcWeek(Date)
        we = Format(week, "mm.dd.yy")
        jobPath = ThisWorkbook.path & "\" & "Data.lnk"
        jobPath = Getlnkpath(jobPath)
    End If
    xlPath = jobPath & "\" & jobNum & "\Week_" & we & "\TimePackets\"
    xlFile = jobNum & "_Week_" & we & "_TimeCards.xlsx"
    If loadRoster = -1 Then GoTo load_err
    If timeCard.loadShifts(test) = -1 Then
        Stop
    End If
    Stop
    hiddenApp.Visible = True
    hiddenApp.Workbooks.Open jobPath & "\Master TC.xlsx", False
    Set wb_tc = hiddenApp.Workbooks("Master TC.xlsx")
    Dim cnt As Integer
    cnt = 1
    Dim tEmp As Variant
    ThisWorkbook.Unprotect xPass
    For Each tEmp In weekRoster
        If tEmp Is Nothing Then
            Exit For
        Else
            wb_tc.Worksheets(1).Copy after:=wb_tc.Worksheets(wb_tc.Sheets.count)
            With wb_tc.Worksheets(wb_tc.Sheets.count)
                .name = "TAG " & cnt
                .Range("e_name") = tEmp.getFullname
                .Range("e_num") = tEmp.getNum
                .Range("we_date") = calcWeek(Date)
                .Range("job_desc") = jobNum & " - " & jobName
                Dim tshft As Variant
                For Each tshft In tEmp.getShifts
                    Dim i As Integer
                    i = 0
rep_add:
                    If tshft.getPhase <> 0 Then
                        If .Range("COST_CODE").Offset(i, 0) = vbNullString Then
                            .Range("COST_CODE").Offset(i, 0) = tshft.getPhase
                            .Range("COST_CODE").Offset(i, tshft.getDay + 3) = tshft.getHrs
                        ElseIf .Range("COST_CODE").Offset(i, 0) = tshft.getPhase Then
                            .Range("COST_CODE").Offset(i, tshft.getDay + 3).Value = tshft.getHrs
                        Else
                            i = i + 1
                            GoTo rep_add
                        End If
                    End If
                Next
            End With
        End If
        cnt = cnt + 1
    Next
    ThisWorkbook.Protect xPass
    Stop
    hiddenApp.DisplayAlerts = False
    wb_tc.Worksheets(1).Delete
    wb_tc.Close True, xlPath & xlFile
    hiddenApp.Quit
    Set hiddenApp = Nothing
    Exit Sub
load_err:
    MsgBox "No Packet Found!"
    
End Sub

Public Sub showHiddenBooks()
    Dim oXLApp As Object

    '~~> Get an existing instance of an EXCEL application object
    On Error Resume Next
    Set oXLApp = GetObject(, "Excel.Application")
    On Error GoTo 0
    oXLApp.Visible = True

    Set oXLApp = Nothing
End Sub

Public Function loadRoster() As Integer
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
    ReDim weekRoster(0, eCount)
    Dim hiddenApp As New Excel.Application
    i = 0
    xlFile = jobPath & "\" & jobNum & "\Week_" & we & "\TimePackets\" & jobNum & "_Week_" & we & ".xlsx"
'    On Error GoTo 10
    hiddenApp.Workbooks.Open xlFile
    SetAttr xlFile, vbNormal
    On Error GoTo 0
    Set bk = hiddenApp.Workbooks(jobNum & "_Week_" & we & ".xlsx")
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
    bk.Worksheets("SAVE").Visible = False
    wb.Activate
    bk.Close False
    hiddenApp.Quit
    Set hiddenApp = Nothing
    loadRoster = 1
    Exit Function
10:
    loadRoster = -1
    hiddenApp.Visible = True
    On Error Resume Next
    For i = 1 To hiddenApp.Workbooks.count
        hiddenApp.Workbooks(i).Close False
    Next
    hiddenApp.Quit
    Set hiddenApp = Nothing
End Function

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

Public Function get_job_value(Optional c As Range) As Integer
    If c Is Nothing Then
        Set c = Application.Caller
    End If
    Dim tmp As Double
    tmp = 0
    Dim rng As Range
    Dim job_cnt As Integer
    Set rng = ThisWorkbook.Worksheets("USER").Range("D" & c.Row)
    job_cnt = c.Column - rng.Column - 1
    For i = 0 To job_cnt
        If rng.Offset(0, i).Value = True Then
            tmp = tmp + Application.WorksheetFunction.Power(2, i)
        End If
    Next i
    get_job_value = tmp
End Function
Public Sub testPacket()
Attribute testPacket.VB_ProcData.VB_Invoke_Func = "r\n14"
    loadMenu
    savePacket
    MsgBox ("Complete")
End Sub
