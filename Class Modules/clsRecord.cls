VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pIndex As Long
Private pName As String
Private pSheet As String
Private pLocation As Range 'where it is at?


Public Sub initialize(index As Long, name As String, sheet As String, location As Range)

On Error GoTo err_trap

pIndex = index
pName = name
pSheet = sheet
Set pLocation = location

Exit_here:
Exit Sub

err_trap:
MsgBox "Error in ""Initiaize"" of clsRecord. Error number: " & Err.Number & ", " & Err.Description
Resume Exit_here

End Sub

Public Property Get index() As Long
index = pIndex
End Property

Public Property Get name() As String
name = pName
End Property

Public Property Get sheet() As String
sheet = pSheet
End Property

Public Property Get location() As Range
Set location = pLocation
End Property

