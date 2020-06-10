VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSplitterItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'this will be single item line in the sheet
Private pIndex As Long 'products index
Private pName As String
Private pProperties As String
Private pUnitRatio As Integer
Private pCustomer As String
Private pMachine As String
Private pMaterialUnit As String
Private pParentIndex As String
Private pParentName As String
Private pMaterialType As String
Private pShiftObjects As New Collection 'collection of items in single row in worksheet

Public Property Get index() As Long
index = pIndex
End Property

Public Property Get name() As String
name = pName
End Property

Public Property Get properties() As String
properties = pProperties
End Property

Public Property Get customer() As String
customer = pCustomer
End Property

Public Property Get machine() As String
machine = pMachine
End Property

Public Property Get unitRatio() As Integer
unitRatio = pUnitRatio
End Property

Public Property Get materialUnit() As String
materialUnit = pMaterialUnit
End Property

Public Sub initialize(index As Long, name As String, Optional properties As Variant, Optional customer As Variant, Optional unitRatio As Variant, Optional materialUnit As Variant, Optional materialType As Variant)
On Error GoTo err_trap

pIndex = index
pName = name
If Not IsMissing(properties) Then pProperties = properties
If Not IsMissing(customer) Then pCustomer = customer
If Not IsMissing(unitRatio) Then
    If IsNumeric(unitRatio) Then
        pUnitRatio = Int(unitRatio)
    Else
        pUnitRatio = 0
    End If
End If
If Not IsMissing(materialUnit) Then
    pMaterialUnit = materialUnit
End If
If Not IsMissing(materialType) Then
    pMaterialType = materialType
End If


Exit_here:
Exit Sub

err_trap:
MsgBox "Error in ""Initiaize"" of clsSplitterItem. Error number: " & Err.Number & ", " & Err.Description
Resume Exit_here

End Sub

Public Sub append(sDate As Date, sNumber As Integer, mUnit As Double, sUnit As Double, operationId As Long, mesString As Variant, machine As String, Optional parentIndex As Variant, Optional parentName As Variant)
Dim nShift As clsShift

On Error GoTo err_trap

saveMachine machine
Set nShift = newShift(CStr(sDate) & "_" & sNumber)
nShift.initialize sDate, sNumber, mUnit, sUnit, operationId, mesString, parentIndex, parentName, machine, pCustomer

If Not IsMissing(parentName) Then
    If InStr(1, pParentName, parentIndex & " " & parentName, vbTextCompare) = 0 Then
        pParentName = pParentName & parentIndex & " " & parentName & ", " & vbNewLine
    End If
End If

'If Not IsMissing(materialType) Then
'    pMaterialType = materialType
'End If

Exit_here:
Exit Sub

err_trap:
MsgBox "Error in ""append"" of clsSplitterItem. Error number: " & Err.Number & ", " & Err.Description
Resume Exit_here

End Sub

Public Function toString()
Dim str As String

str = "Index: " & pIndex & vbNewLine
str = str & "Nazwa: " & pName & vbNewLine
If pProperties <> "" Then str = str & "Właściwości: " & pProperties & vbNewLine
If pMachine <> "" Then str = str & "Maszyna: " & pMachine & vbNewLine
If pCustomer <> "" Then str = str & "Klient: " & pCustomer & vbNewLine
If pMaterialType <> "" Then str = str & "Kategoria: " & pMaterialType & vbNewLine
If pParentName <> "" Then str = str & "Zfin: " & pParentName & vbNewLine


toString = str

End Function

Private Function newShift(ref As String) As clsShift
Dim bool As Boolean
Dim cShift As clsShift

On Error GoTo err_trap

For Each cShift In pShiftObjects
    If cShift.shiftName = ref Then
        Set newShift = cShift
        bool = True
        Exit For
    End If
Next cShift

If Not bool Then
    Set newShift = New clsShift
    'newItem.initialize
    pShiftObjects.add newShift, ref
End If

Exit_here:
Exit Function

err_trap:
MsgBox "Error in ""newShift"" of clsSplitterItem. Error number: " & Err.Number & ", " & Err.Description
Resume Exit_here

End Function

Public Function getShifts(Optional shiftName As Variant) As Variant 'shiftName e.g. '2017-11-06_3' for 3rd shift
If IsMissing(shiftName) Then
    Set getShifts = pShiftObjects
Else
    If inCollection(CStr(shiftName), pShiftObjects) Then
        getShifts = pShiftObjects(shiftName)
    Else
        Set getShifts = Null
    End If
End If
End Function

Public Function getTotal() As Double()
Dim shift As clsShift
Dim tot(1) As Double
Dim M As Double
Dim S As Double

For Each shift In pShiftObjects
    M = M + shift.MainUnit
    S = S + shift.secUnit
Next shift

tot(0) = M
tot(1) = S

getTotal = tot

End Function

Private Sub saveMachine(mach As String)
If InStr(1, pMachine, mach, vbTextCompare) = 0 Then
    If Len(pMachine) > 0 Then
        pMachine = pMachine & "; " & mach
    Else
        pMachine = pMachine & mach
    End If
End If
End Sub


