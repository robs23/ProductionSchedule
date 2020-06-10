Attribute VB_Name = "SQL"
Public Sub updateConnection()

If Not adoConn Is Nothing Then
    If adoConn.State = 0 Then
        adoConn.Open ConnectionString
        adoConn.CommandTimeout = 90
    End If
Else
    Set adoConn = New ADODB.Connection
    adoConn.Open ConnectionString
    adoConn.CommandTimeout = 90
End If
End Sub

Public Sub closeConnection()

If Not adoConn Is Nothing Then
    If adoConn.State = 1 Then
        adoConn.Close
    End If
    Set adoConn = Nothing
End If
End Sub

Public Sub sendToBase()
Dim iStr As String
Dim sStr As String
Dim i As Integer

updateConnection
With ThisWorkbook.Sheets("tutaj")
    For i = 1 To 48
        If .Cells(i, 1) = "" Then
            Exit For
        Else
            sStr = sStr & "(" & .Cells(i, 1) & ",'" & .Cells(i, 2) & "','zcom','" & Now & "',1),"
        End If
    Next i
End With
sStr = Left(sStr, Len(sStr) - 1)
iStr = "INSERT INTO tbZfin (zfinIndex, zfinName, zfinType, creationDate, createdBy) Values " & sStr

adoConn.Execute iStr

closeConnection
End Sub

Public Sub connectScada()
'Dim cmd As ADODB.Command
'Set cmd = New ADODB.Command

If scadaConn Is Nothing Then
    Set scadaConn = New ADODB.Connection
    scadaConn.Provider = "SQLOLEDB"
    scadaConn.ConnectionString = ScadaConnectionString
    scadaConn.Open
    scadaConn.CommandTimeout = 90
Else
    If scadaConn.State = adStateClosed Then
        Set scadaConn = New ADODB.Connection
        scadaConn.Provider = "SQLOLEDB"
        scadaConn.ConnectionString = ScadaConnectionString
        scadaConn.Open
        scadaConn.CommandTimeout = 90
    End If
End If


End Sub

Public Sub disconnectScada()

If Not scadaConn Is Nothing Then
    If scadaConn.State = 1 Then
        scadaConn.Close
    End If
    Set scadaConn = Nothing
End If
End Sub

