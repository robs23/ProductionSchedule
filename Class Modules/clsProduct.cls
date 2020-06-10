VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pIndex As Long
Private pName As String
Private pPrimaryAmount As Double
Private pSecondaryAmount As Double
Private pMachines As String
Private pDirections As String
Private pCustomers As String
Private pLplants As String
Private ppalletType As String
Private pChep As Boolean

Public Property Get index() As Long
index = pIndex
End Property

Public Property Let index(ind As Long)
pIndex = ind
End Property

Public Property Get name() As String
name = pName
End Property

Public Property Get primaryAmount() As Double
primaryAmount = pPrimaryAmount
End Property

Public Property Get secondaryAmount() As Double
secondaryAmount = pSecondaryAmount
End Property

Public Property Get machines() As String
machines = pMachines
End Property

Public Property Get directions() As String
directions = pDirections
End Property

Public Property Get customers() As String
customers = pCustomers
End Property

Public Property Get Lplants() As String
Lplants = pLplants
End Property

Public Sub initialize(index As Long, name As String, Optional palletType As Variant, Optional palletChep As Variant)
pIndex = index
pName = name
If Not IsMissing(palletType) Then
    ppalletType = palletType
End If
If Not IsMissing(palletChep) Then
    If palletChep = 1 Then
        pChep = True
    ElseIf palletChep = 0 Then
        pChep = False
    Else
        pChep = palletChep
    End If
End If
End Sub

Public Sub append(Optional primaryAmount As Variant, Optional secondaryAmount As Variant, Optional mach As Variant, Optional direction As Variant, Optional customer As String, Optional Lplant As Variant)

If Not IsMissing(primaryAmount) And Not IsNull(primaryAmount) Then
    pPrimaryAmount = pPrimaryAmount + primaryAmount
End If
If Not IsMissing(secondaryAmount) And Not IsNull(secondaryAmount) Then
    pSecondaryAmount = pSecondaryAmount + secondaryAmount
End If
If Not IsMissing(mach) And Not IsNull(mach) Then
    If InStr(1, pMachines, mach, vbTextCompare) = 0 Then
        If Len(pMachines) = 0 Then
            pMachines = mach
        Else
            pMachines = pMachines & ";" & mach
        End If
    End If
End If
If Not IsMissing(direction) And Not IsNull(direction) Then
    If InStr(1, pDirections, direction, vbTextCompare) = 0 Then
        If Len(pDirections) = 0 Then
            pDirections = direction
        Else
            pDirections = pDirections & ";" & direction
        End If
    End If
End If
If Not IsMissing(customer) And Not IsNull(customer) Then
    If InStr(1, pCustomers, customer, vbTextCompare) = 0 Then
        If Len(pCustomers) = 0 Then
            pCustomers = customer
        Else
            pCustomers = pCustomers & ";" & customer
        End If
    End If
End If
If Not IsMissing(Lplant) And Not IsNull(Lplant) Then
    If InStr(1, pLplants, Lplant, vbTextCompare) = 0 Then
        If Len(pLplants) = 0 Then
            pLplants = Lplant
        Else
            pLplants = pLplants & ";" & Lplant
        End If
    End If
End If
End Sub

Public Function toString() As String
toString = pIndex & " " & Trim(pName)
If Len(pMachines) > 0 Then toString = toString & ", " & Trim(pMachines)
If Len(pDirections) > 0 Then toString = toString & ", " & pDirections
If Len(pLplants) > 0 Then toString = toString & ", " & Trim(pLplants)
End Function
