Attribute VB_Name = "gen_TimeCard"
Public eRoster As Collection
Public leads As Collection
Enum day
    Err = 0
    Monday = 1
    Tuesday = 2
    Wednesday = 3
    Thursday = 4
    Friday = 5
    Saturday = 6
    Sunday = 7
End Enum

Sub getSheet(book As String, sheet As String, start As Integer, finish As Integer)

    Dim wb As Workbook
    Dim temp As Integer
    temp = start
    Set wb = ThisWorkbook
    On Error GoTo 10
    Set bk = Workbooks(book)
    On Error GoTo 0
    For i = start To finish
        bk.Worksheets(i).Copy before:=wb.Worksheets(1)
        On Error GoTo 20
        wb.Worksheets(1).name = sheet & " " & temp
        wb.Worksheets(1).Move after:=wb.Sheets(wb.Sheets.count)
        temp = temp + 1
    Next i
    GoTo 30

10: Workbooks.Open Filename:=book, UpdateLinks:=0
    Set bk = ActiveWorkbook
    Resume Next
    
20: temp = temp + 1
    Resume
    
30:
    bk.Close
    wb.Worksheets(1).Activate
End Sub
Public Sub showBooks()
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
Sub openFiles()
'UpdateByExtendoffice20160623
    Dim list As Range
    Dim wb As Workbook
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    Set list = wb.Worksheets("Lead Files").Range("A1")
    Dim xStrPath As String
    Dim xFileDialog As FileDialog
    Dim xFile As String
'    On Error Resume Next
    xStrPath = timeCard.Getlnkpath(ThisWorkbook.path & "\Data.lnk")
    xStrPath = xStrPath & timeCard.jobNum & "\TimeSheets\Week_" & Format(timeCard.week, "mm.dd.yy") & "\"
    Debug.Print xStrPath
    If xStrPath = "" Then Exit Sub
    xFile = Dir(xStrPath & "\*.xlsx")
    Do While xFile <> ""
        Workbooks.Open xStrPath & "\" & xFile
        wb.Activate
        list.Value = xStrPath & "\" & xFile
        Set list = list.Offset(1, 0)
        xFile = Dir
                
    Loop
End Sub

Function getBook() As String
    Dim book As String
    book = Application.GetOpenFilename(title:="Please choose a file to open", _
        FileFilter:="Excel Files *.xls* (*.xls*),")
    Dim x As Integer
    x = 1
    If IsEmpty(ThisWorkbook.Worksheets("Lead Files").Range("A" & x)) Then
        ThisWorkbook.Worksheets("Lead Files").Range("A" & x).Value = book
    Else
        x = x + 1
    End If
    getBook = book
End Function

Public Sub genLead(Optional num As Integer)
    Application.DisplayAlerts = False
    If num = 0 Then num = 1
    Set wb = ThisWorkbook
    Dim file As String
    Dim y As Integer
    y = 1
    Dim list(100) As String
    For st = 0 To 100
        list(st) = "BLANK"
    Next st
    Dim tag As Integer
    tag = 100
5:
    If IsEmpty(ThisWorkbook.Worksheets("Lead Files").Range("A1")) Then
        Call openFiles
'        GoTo 5
    Else
        For ii = 1 To tag
            If IsEmpty(ThisWorkbook.Worksheets("Lead Files").Range("A" & ii)) Then
                Exit For
            Else
                list(ii - 1) = ThisWorkbook.Worksheets("Lead Files").Range("A" & ii).Value
            End If
            If ii = 100 Then
                MsgBox ("YOU MADE ERROR - I GO BOOM!")
                GoTo 30
            End If
        Next ii
    End If
    '5:
    '    If IsEmpty(wb.Worksheets("Lead Files").Range("A" & y)) And y = 1 Then
    '        file = getBook
    '        y = y + 1
    '    Else
    '        If IsEmpty(wb.Worksheets("Lead Files").Range("A" & y)) Then
    '            GoTo 100
    '        Else
    '            file = wb.Worksheets("Lead Files").Range("A" & y).Value
    '            y = y + 1
    '        End If
    '    End If
    ii = 0
    Do While list(ii) <> "BLANK"
        file = list(ii)
        Dim bk As Workbook
        On Error GoTo 10
        Set bk = Workbooks(file)
        On Error Resume Next
        wb.Worksheets("LEAD").Copy before:=wb.Worksheets(1)
        Dim sht As Worksheet
        Set sht = wb.Worksheets(1)
        sht.name = "LEAD " & num
        sht.Visible = xlSheetVisible
        If num = 1 Then
            wb.Worksheets("LEAD " & num).Move after:=wb.Worksheets("LEAD")
        Else
            wb.Worksheets("LEAD " & num).Move after:=wb.Worksheets("LEAD " & num - 1)
        End If
            
        
        Dim eRoster2 As Collection
        Set eRoster2 = New Collection
        Dim tNum As Integer
        
        tNum = 3
            
        For i = 1 To bk.Worksheets(2).UsedRange.Rows.count
            If bk.Worksheets(2).Range("B" & i).Value = True Then
                Dim emp As Employee
                Dim i2 As Integer
                Set emp = New Employee
                Let emp.efName = bk.Worksheets(2).Range("D" & i)
                Let emp.elName = bk.Worksheets(2).Range("C" & i)
                Let emp.emnum = bk.Worksheets(2).Range("E" & i)
                Let emp.emClass = getClass(emp)
                For i2 = 0 To 13
                    Dim sa As shift
                    Dim sb As shift
                    Set sa = New shift
                    Set sb = New shift
                    Dim temp As Integer
                    temp = CInt(i2 \ 2) + 1
                    sa.setDay = temp
                    sa.setPhase = bk.Worksheets(CInt(i2 \ 2) + 3).Range("B" & tNum).Value
                    sa.setHrs = bk.Worksheets(CInt(i2 \ 2) + 3).Range("C" & tNum).Value
                    Call emp.addShift(sa)
                    i2 = i2 + 1
                    sb.setDay = CInt(i2 \ 2) + 1
                    sb.setPhase = bk.Worksheets(CInt(i2 \ 2) + 3).Range("D" & tNum).Value
                    sb.setHrs = bk.Worksheets(CInt(i2 \ 2) + 3).Range("E" & tNum).Value
                    Call emp.addShift(sb)
                Next i2
                
                eRoster2.Add emp
                tNum = tNum + 1
            End If
        Next i
        
        tNum = 3
        For Each emp In eRoster2
            If eRoster Is Nothing Then
                Set eRoster = New Collection
            End If
            eRoster.Add emp
            wb.Worksheets("LEAD " & num).Activate
            ActiveSheet.Cells(tNum, 1).Select
            Dim os As Integer, x As Integer
            Dim s2 As Collection
            Dim dayOS(6) As Integer
            For i = 0 To 6
                dayOS(i) = 1
            Next i
            Dim shft As shift
            With ActiveCell
                .Value = emp.getNum
                .Offset(0, 1).Value = emp.getFName & " " & emp.getLName
                For Each shft In emp.getShifts
                    Dim today As Integer
                    today = shft.getDay
                    If shft.getHrs > 0 And today > 0 Then
                        If dayOS(today - 1) = 1 Then
                            .Offset(0, (today * 6) - (4 * dayOS(today - 1))).Value = shft.getHrs
                            .Offset(0, today * 6 - 3 * dayOS(today - 1)).Value = shft.getPhase
                            dayOS(today - 1) = 0
                        ElseIf dayOS(today - 1) = 0 Then
                            .Offset(0, (today * 6) - 1).Value = shft.getHrs
                            .Offset(0, (today * 6)).Value = shft.getPhase
                            dayOS(today - 1) = -1
                        Else
                            MsgBox ("TOO MANY PHASES IN ONE DAY")
                            End
                        End If
                    End If
                Next shft
                    
    '            For x = 0 To 13
    '                If emp.getDayHrs(x) > 0 Then
    '                    .Offset(0, os).Value = emp.getDayHrs(x)
    '                    os = os + 1
    '                    .Offset(0, os).Value = emp.getDayPhase(x)
    '                    os = os + 2
    '                Else
    '                    os = os + 3
    '                End If
    '            Next x
            End With
            tNum = tNum + 2
        Next emp
    bk.Close
    num = num + 1
    ii = ii + 1
    Loop
    GoTo 30

10:
    Workbooks.Open Filename:=file, UpdateLinks:=0
    Set bk = ActiveWorkbook
    Resume Next
    
30:
    wb.Worksheets(1).Activate
    Application.DisplayAlerts = True
End Sub

Sub showlist(e As Collection)
    Dim emp As Employee
    For Each emp In e
        MsgBox ("First Name: " & emp.getFName & " " & emp.getLName)
    Next emp
        
End Sub

Function sheetExists(sheetToFind As String) As Boolean
    sheetExists = False
    For Each sheet In Worksheets
        If sheetToFind = sheet.name Then
            sheetExists = True
            Exit Function
        End If
    Next sheet
End Function

Sub genRoster()
    Dim tNum As Integer
    Dim emp As Employee
    Dim x As Integer
    Set emp = New Employee
    x = 1
    For Each emp In eRoster
            ThisWorkbook.Worksheets("ROSTER").Activate
            ActiveSheet.Cells(8 + x, 1).Select
            With ActiveCell
                .Value = x
                .BorderAround Weight:=xlThin
                .Offset(0, 1).Value = emp.getClass
                .Offset(0, 1).BorderAround Weight:=xlThin
                .Offset(0, 2).Value = emp.getFName & " " & emp.getLName
                .Offset(0, 2).BorderAround Weight:=xlThin
                .Offset(0, 3).Value = emp.getNum
                .Offset(0, 3).BorderAround Weight:=xlThin
                .Offset(0, 4).Value = "88070-80 Per Diem"
                .Offset(0, 4).BorderAround Weight:=xlThin
                
            End With
            x = x + 1
            Next emp
End Sub

Sub genEmpFromLead()
    Application.DisplayAlerts = Flase
Set wb = ThisWorkbook
    Dim file As String
    Dim y As Integer
    y = 1
    Dim list(100) As String
    For st = 0 To 100
        list(st) = "BLANK"
    Next st
    Dim tag As Integer
    tag = 100
5:
    If IsEmpty(ThisWorkbook.Worksheets("Lead Files").Range("A1")) Then
        Call openFiles
'        GoTo 5
    Else
        For ii = 1 To tag
            If IsEmpty(ThisWorkbook.Worksheets("Lead Files").Range("A" & ii)) Then
                Exit For
            Else
                list(ii - 1) = ThisWorkbook.Worksheets("Lead Files").Range("A" & ii).Value
            End If
            If ii = 100 Then
                MsgBox ("YOU MADE ERROR - I GO BOOM!")
                GoTo 30
            End If
        Next ii
    End If
    '5:
    '    If IsEmpty(wb.Worksheets("Lead Files").Range("A" & y)) And y = 1 Then
    '        file = getBook
    '        y = y + 1
    '    Else
    '        If IsEmpty(wb.Worksheets("Lead Files").Range("A" & y)) Then
    '            GoTo 100
    '        Else
    '            file = wb.Worksheets("Lead Files").Range("A" & y).Value
    '            y = y + 1
    '        End If
    '    End If
    ii = 0
    num = 1
    Do While list(ii) <> "BLANK"
        file = list(ii)
        Dim bk As Workbook
        On Error GoTo 10
        Set bk = Workbooks(file)
        On Error Resume Next
        wb.Worksheets("LEAD").Copy before:=wb.Worksheets(1)
        Dim sht As Worksheet
        Set sht = wb.Worksheets(1)
        sht.name = "LEAD " & num
        sht.Visible = xlSheetVisible
        If num = 1 Then
            wb.Worksheets("LEAD " & num).Move after:=wb.Worksheets("LEAD")
        Else
            wb.Worksheets("LEAD " & num).Move after:=wb.Worksheets("LEAD " & num - 1)
        End If
            
        Dim eRoster2 As Collection
        Set eRoster2 = New Collection
        Dim tNum As Integer
        
        tNum = 2
        Dim done As Boolean
        done = False
        Dim i As Integer
        i = 1
        Do While done = False
            If IsEmpty(bk.Worksheets(1).Range("B" & i + tNum)) Then
                done = True
            Else
                Dim emp As Employee
                Dim i2 As Integer
                Set emp = New Employee
                Dim fullName As Variant
                fullName = Split(bk.Worksheets(1).Range("B" & i + tNum), " ")
                Let emp.efName = fullName(0)
                Let emp.elName = fullName(1)
                Let emp.emnum = bk.Worksheets(1).Range("A" & i + tNum)
                Let emp.emClass = getClass(emp)
                For i2 = 0 To 13
                    Dim sa As shift
                    Dim sb As shift
                    Set sa = New shift
                    Set sb = New shift
                    Dim temp As Integer
                    temp = CInt(i2 \ 2) + 1
                    If IsEmpty(bk.Worksheets(1).Cells(i + tNum, ((temp) * 6) - 3).Value) Then
                        sa.setHrs = 0
                    Else
                        sa.setHrs = bk.Worksheets(1).Cells(i + tNum, ((temp) * 6) - 3).Value
                    End If
                    If IsEmpty(bk.Worksheets(1).Cells(i + tNum, ((temp) * 6) - 2).Value) Then
                        sa.setPhase = " "
                    Else
                        sa.setPhase = bk.Worksheets(1).Cells(i + tNum, ((temp) * 6) - 2).Value
                    End If
                    sa.setDay = temp
                    Call emp.addShift(sa)
                    i2 = i2 + 1
                    temp = CInt(i2 \ 2) + 1
                    If IsEmpty(bk.Worksheets(1).Cells(i + tNum, ((temp) * 6)).Value) Then
                        sb.setHrs = 0
                    Else
                        sb.setHrs = bk.Worksheets(1).Cells(i + tNum, ((temp) * 6)).Value
                    End If
                    If IsEmpty(bk.Worksheets(1).Cells(i + tNum, ((temp) * 6) + 1).Value) Then
                        sb.setPhase = " "
                    Else
                        sb.setHrs = bk.Worksheets(1).Cells(i + tNum, ((temp) * 6) + 1).Value
                    End If
                    sa.setDay = temp

'                    sb.setDay = temp
'                    sb.setPhase = bk.Worksheets(1).Cells(i + tNum, ((temp) * 6) - 2).Value
'                    sb.setHrs = bk.Worksheets(1).Cells(i + tNum, ((temp) * 6) - 3).Value
                    Call emp.addShift(sb)
                Next i2
                
                eRoster2.Add emp
                tNum = tNum + 1
            End If
        i = i + 1
        Loop
        
        tNum = 3
        For Each emp In eRoster2
            If eRoster Is Nothing Then
                Set eRoster = New Collection
            End If
            eRoster.Add emp
            wb.Worksheets("LEAD " & num).Activate
            ActiveSheet.Cells(tNum, 1).Select
            Dim os As Integer, x As Integer
            Dim s2 As Collection
            Dim dayOS(6) As Integer
            For i = 0 To 6
                dayOS(i) = 1
            Next i
            Dim shft As shift
            With ActiveCell
                .Value = emp.getNum
                .Offset(0, 1).Value = emp.getFName & " " & emp.getLName
                For Each shft In emp.getShifts
                    Dim today As Integer
                    today = shft.getDay
                    If shft.getHrs > 0 And today > 0 Then
                        If dayOS(today - 1) = 1 Then
                            .Offset(0, (today * 6) - (4 * dayOS(today - 1))).Value = shft.getHrs
                            .Offset(0, today * 6 - 3 * dayOS(today - 1)).Value = shft.getPhase
                            dayOS(today - 1) = 0
                        ElseIf dayOS(today - 1) = 0 Then
                            .Offset(0, (today * 6) - 1).Value = shft.getHrs
                            .Offset(0, (today * 6)).Value = shft.getPhase
                            dayOS(today - 1) = -1
                        Else
                            MsgBox ("TOO MANY PHASES IN ONE DAY")
                            End
                        End If
                    End If
                Next shft
                    
    '            For x = 0 To 13
    '                If emp.getDayHrs(x) > 0 Then
    '                    .Offset(0, os).Value = emp.getDayHrs(x)
    '                    os = os + 1
    '                    .Offset(0, os).Value = emp.getDayPhase(x)
    '                    os = os + 2
    '                Else
    '                    os = os + 3
    '                End If
    '            Next x
            End With
            tNum = tNum + 2
        Next emp
    bk.Close
    num = num + 1
    ii = ii + 1
    Loop
    GoTo 30

10:
    Workbooks.Open Filename:=file, UpdateLinks:=0
    Set bk = ActiveWorkbook
    Resume Next
    
30:
    wb.Worksheets(1).Activate
    Application.DisplayAlerts = True
End Sub

Function getClass(emp As Employee) As String
    Dim i As Integer
    i = 1
5:
    If emp.getNum = ThisWorkbook.Worksheets("MASTER").Cells(i, 4) Then
        getClass = ThisWorkbook.Worksheets("MASTER").Cells(i, 5).Value
    Else
        i = i + 1
        GoTo 5
    End If
End Function

Sub copySheet()
    Workbooks("Design File v2.xlsm").Worksheets("MASTER").Copy before:=ThisWorkbook.Worksheets(1)
    
End Sub

Sub delSheet(sht As String)
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(sht).Delete
    Application.DisplayAlerts = True
End Sub

Sub genTags()
    Set wb = ThisWorkbook
    wb.Application.DisplayAlerts = False
    Dim xlPath As String
    Dim title As String
    xlPath = ThisWorkbook.path
    title = Format(ThisWorkbook.Worksheets("Tag").Range("E10").Value, "mm.dd.yy") & ".xlsx"
    Dim num As Integer
    Dim timeBook As Workbook
    Set timeBook = Workbooks.Add()
    timeBook.Application.DisplayAlerts = False
    Dim lNum As Integer
    Dim emp As Employee
    Set emp = New Employee
    num = 1
                
    If eRoster Is Nothing Then
        Set eRoster = New Collection
    End If
    
    If eRoster.count = 0 Then
        Call genLead(1)
        wb.Application.DisplayAlerts = False
        On Error Resume Next
    End If
        
    For Each emp In eRoster
        Dim s As shift
        Set s = New shift
        wb.Worksheets("Tag").Copy before:=timeBook.Worksheets(1)
        Dim sht As Worksheet
        Set sht = timeBook.Worksheets(1)
        Do While sheetExists("Tag #" & num)
            num = num + 1
        Loop
            
        sht.name = "Tag #" & num
        sht.Cells(7, 4).Value = emp.getFName & " " & emp.getLName
        sht.Cells(8, 4).Value = emp.getNum
        For Each s In emp.getShifts
            For i = 12 To 36
                If sht.Cells(i, 1).Value = s.getPhase Then
                    sht.Cells(i, s.getDay + 4).Value = sht.Cells(i, s.getDay + 4).Value + s.getHrs
                End If
            Next i
        Next s
        sht.Visible = xlSheetVisible
        timeBook.Worksheets("Tag #" & num).Move after:=timeBook.Worksheets(Sheets.count)
    Next emp
    timeBook.Worksheets("Sheet1").Delete
    wb.Worksheets("ROSTER").Copy before:=timeBook.Worksheets(1)
    Call genRoster
    timeBook.SaveAs Filename:=xlPath & "\Time_Cards_" & title
    timeBook.Worksheets("Tag #1").Activate
    timeBook.Close
    wb.Worksheets("ROSTER").Range("A9:G1000").Clear
'    For i = 6 To Sheets.Count
'        wb.Worksheets(i).Delete
'    Next i
    MsgBox ("Your time sheets are ready! They can be found at: " & xlPath & "\Time_Cards_" & title)
    timeBook.Application.DisplayAlerts = True
    Application.DisplayAlerts = True

End Sub

Public Sub showSheets()
    For i = 2 To Sheets.count
        ThisWorkbook.Worksheets(i).Visible = True
    Next i
End Sub

Public Sub hideSheets()
    For i = 2 To Sheets.count
        Worksheets(i).Visible = xlVeryHidden
    Next i
End Sub

