Attribute VB_Name = "ObjectStore"
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal length As Long)

Public Sub StoreObjRef(obj As Object, propertyName As String)
' Store an object reference
 Dim longObj As Long
 longObj = ObjPtr(obj)
 updateProperty propertyName, longObj
End Sub
 
Function RetrieveObjRef(propertyName As String) As Object
' Retrieve the object reference
 Dim longObj As Long, obj As Object
 longObj = ThisWorkbook.CustomDocumentProperties(propertyName)
 CopyMemory obj, longObj, 4
 Set RetrieveObjRef = obj
End Function

Public Sub bbRib()
If rib Is Nothing Then
    Set rib = RetrieveObjRef("ribbon_ref")
End If
End Sub


