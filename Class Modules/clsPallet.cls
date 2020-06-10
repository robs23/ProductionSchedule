VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPallet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private palletId As Integer
Private palletName As String
Private palletLength As Integer
Private palletWidth As Integer
Private palletChep As Boolean

Public Property Let id(pId As Integer)
    palletId = pId
End Property

Public Property Let name(pName As String)
    palletName = pName
End Property

Public Property Let length(pLength As Integer)
    palletLength = pLength
End Property

Public Property Let width(pWidth As Integer)
    palletWidth = pWidth
End Property

Public Property Let chep(bool As Boolean)
    palletChep = bool
End Property

Public Property Get id() As Integer
    id = palletId
End Property

Public Property Get name() As String
    name = palletName
End Property

Public Property Get length() As Integer
    length = palletLength
End Property

Public Property Get width() As Integer
    width = palletWidth
End Property

Public Property Get chep() As Boolean
    chep = palletChep
End Property

