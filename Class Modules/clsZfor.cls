VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsZfor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pZforNumber As Long
Private pRn3000input As Variant
Private pRn4000input As Variant

Public Sub initialize(zforNumber As Long, rn3000 As Variant, rn4000 As Variant)
pZforNumber = zforNumber
pRn3000input = rn3000
pRn4000input = rn4000
End Sub

Public Function index() As Long
index = pZforNumber
End Function

Public Function rn3000input() As Variant
rn3000input = pRn3000input
End Function

Public Function rn4000input() As Variant
rn4000input = pRn4000input
End Function
