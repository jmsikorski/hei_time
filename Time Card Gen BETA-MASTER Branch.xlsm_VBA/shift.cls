VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private d As Double
Private phase As Long
Private hrs As Double

Private Sub Class_Initialize()
    d = 0
    phase = 0
    hrs = 0
End Sub

Public Property Let setDay(dy As Double)
    d = dy
End Property

Public Property Get getDay() As Double
    getDay = d
End Property

Public Property Let setPhase(p As Long)
    phase = p
End Property

Public Property Get getPhase() As Long
    getPhase = phase
End Property

Public Property Let setHrs(h As Double)
    hrs = h
End Property

Public Property Get getHrs() As Double
    getHrs = hrs
End Property
