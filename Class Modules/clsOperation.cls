VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pOperationId As Long 'mes's operation id
Private pMesString As String 'mes's operation string e.g. PAK/20170301/....

Public Property Let operationId(val As Long)
pOperationId = val
End Property

Public Property Get operationId() As Long
operationId = pOperationId
End Property

Public Property Let mesString(val As String)
pMesString = val
End Property

Public Property Get mesString() As String
mesString = pMesString
End Property

