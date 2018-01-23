Attribute VB_Name = "encryption"
Public Function encryptPassword(pw As String) As String
    Dim pwi() As Long
    Dim test() As String
    Dim epw As String
    epw = vbnullStrig
    ReDim test(Len(pw))
    ReDim pwi(Len(pw))
    Dim x As Integer
    x = 1
    For i = 0 To Len(pw) - 1
        test(i) = Left(pw, 1)
        pwi(i) = Asc(test(i))
        pw = Right(pw, Len(pw) - 1)
        pwi(i) = pwi(i) Xor ThisWorkbook.Worksheets("KEY").Range("A" & x).Value
        If pwi(i) = 0 Then pwi(i) = 1
        epw = epw & Chr(pwi(i))
    Next i
    encryptPassword = epw
End Function

Public Function testPW(pw As String, tpw As String) As Boolean
    Dim x As Integer
    Dim y As Integer
    x = 0
    y = 0
    If Len(pw) <> Len(tpw) Then
        testPW = False
        Exit Function
    End If
    For i = 0 To Len(tpw)
        On Error Resume Next
        y = Asc(Left(tpw, 1))
        x = Asc(Left(pw, 1))
        If x <> y Then
            testPW = False
            Exit Function
        End If
        tpw = Right(tpw, Len(tpw) - 1)
        pw = Right(pw, Len(pw) - 1)
        On Error GoTo 0
    Next i
    testPW = True
End Function

Public Sub test2()
    Dim t As String
    t = InputBox("Input Password")
    MsgBox ("input: " & t)
    t = encryptPassword(t)
    MsgBox ("encrypted: " & t)
    t = encryptPassword(t)
    MsgBox ("decrypted: " & t)
End Sub

