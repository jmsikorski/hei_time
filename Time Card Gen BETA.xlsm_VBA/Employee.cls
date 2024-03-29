VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private fName As String
Private lName As String
Private num As Integer
Private eClass As String
Private ePerDiem As Boolean
Private shifts As Collection
Private lead As Integer

Public Property Let eLead(i As Integer)
    lead = i
End Property

Public Property Let efName(s As String)
    fName = s
End Property

Public Property Let elName(s As String)
    lName = s
End Property

Public Property Let emnum(s As Integer)
    num = s
End Property

Public Property Let emClass(s As String)
    eClass = s
End Property

Public Sub addShift(s As shift)
    shifts.Add s
End Sub

Public Property Get getLead() As Integer
    getLead = lead
End Property

Public Property Get getFName() As String
    getFName = fName
End Property

Public Property Get getLName() As String
    getLName = lName
End Property

Public Property Get getNum() As Integer
    getNum = num
End Property

Public Property Get getClass() As String
    getClass = eClass
End Property

Public Property Get getShifts() As Collection
    Set getShifts = shifts
End Property

Public Property Let emPerDiem(pd As Boolean)
    ePerDiem = pd
End Property

Public Property Get getPerDiem() As Boolean
    getPerDiem = ePerDiem
End Property

Private Sub Class_Initialize()
    fName = "NONE"
    lName = "NONE"
    num = -1
    eClass = "NONE"
    ePerDiem = False
    lead = -1
    Set shifts = New Collection
End Sub

Public Property Get getFullname() As String
    getFullname = fName & " " & lName
End Property

Public Property Get newEmployee(ByVal i As Integer) As Integer
    On Error GoTo 10
    Dim rg As Range
    Set rg = ThisWorkbook.Worksheets("ROSTER").Range("A" & i + 2)
        With Me
            .efName = rg.Offset(0, 3).Value
            .elName = rg.Offset(0, 2).Value
            .emnum = rg.Offset(0, 4).Value
            .emClass = rg.Offset(0, 1).Value
            If rg.Offset(0, 5).Value = "YES" Then
                .emPerDiem = True
            Else
                .emPerDiem = False
            End If
        End With
    On Error GoTo 0
    newEmployee = 1
    Exit Property
10
    On Error GoTo 0
    newEmployee = -1
End Property

