VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsZfin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private ind As Long
Private desc As String
Private unitW As Double
Private pCount As Long
Private bCount As Integer
Private pType As String
Private pCustomer As String

Public Property Let index(iIndex As Long)
    ind = iIndex
End Property

Public Property Let name(nname As String)
    desc = nname
End Property

Public Property Let palletCount(pcPerPallet As Long)
    pCount = pcPerPallet
End Property

Public Property Let unitWeight(uWeight As Double)
    unitW = uWeight
End Property

Public Property Let boxCount(pcPerBox As Integer)
    bCount = pcPerBox
End Property

Public Property Let palletType(pallet As String)
    pType = pallet
End Property

Public Property Let customer(custName As String)
    pCustomer = custName
End Property

Public Property Get index() As Long
    index = ind
End Property

Public Property Get name() As String
    name = desc
End Property

Public Property Get palletCount() As Long
    palletCount = pCount
End Property

Public Property Get boxCount() As Integer
    boxCount = bCount
End Property

Public Property Get unitWeight() As Double
    unitWeight = unitW
End Property

Public Property Get palletType() As String
    palletType = pType
End Property

Public Property Get boxPerPallet() As Integer
    boxPerPallet = pCount / bCount
End Property

Public Property Get customer() As String
    customer = pCustomer
End Property
