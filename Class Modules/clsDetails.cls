VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pName As String 'combination of parent index and machine name e.g. 4009033_Linia 10
Private pMachine As String
Private pCustomer As String
Private pParentIndex As Long
Private pParentName As String
Private pMainUnit As Double
Private pSecUnit As Double

Public Property Get name() As String
name = pName
End Property


Public Sub append(parentIndex As Variant, parentName As Variant, machine As Variant, customer As Variant, MainUnit As Double, secUnit As Double)
If Not IsMissing(parentIndex) Then pParentIndex = parentIndex
If Not IsMissing(parentName) Then pParentName = parentName
If Not IsMissing(machine) Then pMachine = machine
If Not IsMissing(customer) Then pCustomer = customer
If Not IsMissing(parentIndex) And Not IsMissing(machine) Then pName = CStr(parentIndex) & "_" & machine
pMainUnit = pMainUnit + MainUnit
pSecUnit = pSecUnit + secUnit
End Sub


Public Function toString() As String
toString = pParentIndex & " " & pParentName & " - " & Round(pMainUnit, 2) & " - " & pMachine
End Function

Public Property Get MainUnit() As Double
MainUnit = pMainUnit
End Property

Public Property Let MainUnit(val As Double)
pMainUnit = val
End Property
