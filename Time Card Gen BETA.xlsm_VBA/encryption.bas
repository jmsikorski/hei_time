Attribute VB_Name = "encryption"
Private Function encryptPassword(pw As String) As String
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
    If encryptPassword(tpw) = pw Then
        testPW = True
    Else
        testPW = False
    End If
End Function

