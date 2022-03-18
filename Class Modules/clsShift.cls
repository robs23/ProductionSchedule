VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsShift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pShiftDate As Date
Private pShiftName As String 'combination of date and shift Number e.g. "2017-04-01_3" for 3rd shift of 2017-04-01
Private pShiftNumber As Integer 'I or II or III
Private pMainUnit As Double 'value in main unit
Private pSecUnit As Double 'value in secondary unit
Private pOperationId As Long 'operation id from mes
Private pMesString As String 'operation string from mes
Private pDetails As New Collection

Public Sub initialize(sDate As Date, sNumber As Integer, mUnit As Double, secUnit As Double, operationId As Long, mesString As Variant, Optional parentIndex As Variant, Optional parentName As Variant, Optional machine As Variant, Optional customer As Variant)
Dim det As clsDetails
Dim name As String

pShiftDate = sDate
pShiftNumber = sNumber
pOperationId = operationId
If Not IsNull(mesString) Then pMesString = mesString
pShiftName = CStr(sDate) & "_" & sNumber
pMainUnit = pMainUnit + mUnit
pSecUnit = pSecUnit + secUnit
name = CStr(parentIndex) & "_" & machine
Set det = newDetail(name)
det.append parentIndex, parentName, machine, customer, mUnit, secUnit

End Sub

Public Property Get shiftDate() As Date
shiftDate = pShiftDate
End Property

Public Property Get shiftNumber() As Integer
shiftNumber = pShiftNumber
End Property

Public Property Get shiftName() As String
shiftName = pShiftName
End Property

Public Property Get MainUnit() As Double
MainUnit = pMainUnit
End Property

Public Property Get secUnit() As Double
secUnit = pSecUnit
End Property

Public Property Get operationId() As Long
operationId = pOperationId
End Property

Public Property Get mesString() As String
mesString = pMesString
End Property


Private Function newDetail(name As String) As clsDetails
Dim bool As Boolean
Dim cDetail As clsDetails

On Error GoTo err_trap

For Each cDetail In pDetails
    If cDetail.name = name Then
        Set newDetail = cDetail
        bool = True
        Exit For
    End If
Next cDetail

If Not bool Then
    Set newDetail = New clsDetails
    'newItem.initialize
    pDetails.add newDetail, CStr(name)
End If

exit_here:
Exit Function

err_trap:
MsgBox "Error in ""newDetail"" of clsShift. Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Function

Public Function printDetails() As String
Dim str As String
Dim det As clsDetails

For Each det In pDetails
    str = str & det.toString & vbNewLine
Next det
printDetails = str
End Function

Public Function getDetails(Optional index As Variant) As Variant
If IsMissing(index) Then
    Set getDetails = pDetails
Else
    If inCollection(CStr(index), pDetails) Then
        Set getDetails = pDetails(index)
    Else
        Set getDetails = Null
    End If
End If
End Function
