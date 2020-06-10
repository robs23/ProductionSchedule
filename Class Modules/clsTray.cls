VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pTrayId As String
Private pTrayName As String

Public Property Let TrayId(val As String)
pTrayId = val
End Property

Public Property Let TrayName(val As String)
pTrayName = val
End Property

Public Property Get TrayId() As String
TrayId = pTrayId
End Property

Public Property Get TrayName() As String
TrayName = pTrayName
End Property


