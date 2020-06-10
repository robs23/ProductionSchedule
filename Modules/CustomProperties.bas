Attribute VB_Name = "CustomProperties"
Public Sub createCustomProperty(theName As String, theValue As Variant)
Dim theType As Variant

Select Case VarType(theValue)
    Case 0 To 1
    theType = Null
    Case 2 To 3
    theType = msoPropertyTypeNumber
    Case 4 Or 5 Or 14
    theType = msoPropertyTypeFloat
    Case 7
    theType = msoPropertyTypeDate
    Case 8
    theType = msoPropertyTypeString
    Case 11
    theType = msoPropertyTypeBoolean
    Case Else
    theType = Null
End Select

If theType = Null Then
    MsgBox "Type of variable ""theValue"" passed to ""createCustomProperty"" could not be determined or is unsuported. No custom property has been created", vbOKOnly + vbExclamation
Else
    ThisWorkbook.CustomDocumentProperties.add name:=theName, LinkToContent:=False, Type:=theType, value:=theValue
'    MsgBox "Property " & theName & " has been created successfully and set to " & theValue, vbOKOnly + vbInformation


End If

End Sub

Public Function propertyExists(name As String) As Boolean
Dim prop As DocumentProperty
propertyExists = False
For Each prop In ThisWorkbook.CustomDocumentProperties
    If prop.name = name Then
        propertyExists = True
        Exit For
    End If
Next prop
End Function

Public Sub updateProperty(propName As String, propValue As Variant)

With ThisWorkbook.CustomDocumentProperties
    If propertyExists(propName) Then
        .item(propName).value = propValue
    Else
        createCustomProperty propName, propValue
    End If
End With

End Sub

Public Sub debugCustomProperties()
Dim prop As DocumentProperty
For Each prop In ThisWorkbook.CustomDocumentProperties
    Debug.Print prop.name
Next prop
End Sub

