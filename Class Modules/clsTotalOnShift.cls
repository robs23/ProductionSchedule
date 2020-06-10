VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTotalOnShift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pShiftName As String 'combination of date and shift Number e.g. "2017-04-01_3" for 3rd shift of 2017-04-01
Private pMainTotal As Double
Private pSecTotal As Double

Public Sub add(mainTotal As Double, secTotal As Double)
pMainTotal = pMainTotal + mainTotal
pSecTotal = pSecTotal + secTotal
End Sub

Public Property Get name() As String
name = pShiftName
End Property

Public Property Get mainTotal() As Double
mainTotal = pMainTotal
End Property

Public Property Get secTotal() As Double
secTotal = pSecTotal
End Property

Public Property Let name(val As String)
pShiftName = val
End Property

