Attribute VB_Name = "Runtime"
Public Sub updateSearch()
If verifyRecordset Then
    finder.Show
End If

End Sub

Public Function verifyRecordset() As Boolean
Dim bool As Boolean

bool = False

If currentSchedule Is Nothing Then
    'no records is found, force user to update schedule
    MsgBox "Pamięć podręczna została wyczyszczona w wyniku błędu lub zamknięcia pliku. Proszę akutalizować harmonogram by skorzystać z tej funkcji", vbOKOnly + vbCritical, "Dane utracone"
Else
    If currentSchedule.getRecords.Count = 0 Then
        'no records is found, force user to update schedule
        MsgBox "Pamięć podręczna została wyczyszczona w wyniku błędu lub zamknięcia pliku. Proszę akutalizować harmonogram by skorzystać z tej funkcji", vbOKOnly + vbCritical, "Dane utracone"
    Else
        'open finder
        bool = True
    End If
End If

verifyRecordset = bool

End Function

Public Function getCollectionMember(ByVal objStr As String, col As Collection) As Variant
'returns an object if it exists in collection, otherwise returns Nothing
Dim obj As Variant

If inCollection(objStr, col) Then
    Set getCollectionMember = col(objStr)
Else
    Set getCollectionMember = Nothing
End If

End Function

Public Sub clearTrace()
Dim sht As Worksheet
Dim splitters As New Collection
Dim S As clsSchduleSplitter
Dim rng As Range
Dim c As Range
Dim start As Integer
Dim dailyUnit As String 'the unit used as total for daily summary
Dim shiftUnit As String 'the unit used as total for shift summary
Dim shiftRow As Integer 'either 4 or 5, depends on if daily summary is displayed above
Dim curDate As Date
Dim dFrom As Date
Dim dTo As Date
Dim theDate As Date
Dim theShift As Integer

On Error GoTo err_trap

If currentSchedule Is Nothing Then
    MsgBox "Ta opcja wymaga, abyś najpierw zaktualizował harmonogram z zaznaczoną opcją ""zakres wg tygodnia"" (zakładka ""Zakres dat"").", vbInformation + vbOKOnly, "Funkcja niedostępna"
Else
    Set splitters = currentSchedule.getSplitters()
    dFrom = currentSchedule.startDate
    dTo = currentSchedule.endDate
    scheduleMode = currentSchedule.mode
    For Each S In splitters
        Set rng = S.workingRange
        rng.Cells.Interior.Color = vbWhite
        shiftRow = 4
        If TotalDaily <> False Then
            shiftRow = 5
            If scheduleMode = 1 Then
                If TotalDaily = "Wsad" Then
                    dailyUnit = "s"
                Else
                    dailyUnit = "m"
                End If
            ElseIf scheduleMode = 3 Or scheduleMode = 4 Then
                If TotalDaily = sUnit Then
                    dailyUnit = "s"
                ElseIf TotalDaily = pUnit Then
                    dailyUnit = "m"
                End If
            End If
        End If
        If TotalShift <> False Then
            If scheduleMode = 1 Then
                If TotalShift = "Wsad" Then
                    shiftUnit = "s"
                Else
                    shiftUnit = "m"
                End If
            ElseIf scheduleMode = 3 Or scheduleMode = 4 Then
                If TotalShift = sUnit Then
                    shiftUnit = "s"
                ElseIf TotalShift = pUnit Then
                    shiftUnit = "m"
                End If
            End If
        End If
        curDate = dFrom
        Do Until curDate > dTo
            Select Case DatePart("h", curDate)
                Case 6: theShift = 1
                Case 14: theShift = 2
                Case 22: theShift = 3
            End Select
            theDate = DateSerial(year(curDate), Month(curDate), Day(curDate))
            theCol = S.getShiftColumn(theDate, theShift)
            If TotalShift <> False Then ThisWorkbook.Sheets(S.name).Cells(shiftRow, theCol) = S.total(theDate, shiftUnit, theShift)
            theCol = S.getShiftColumn(theDate, 1)
            If TotalDaily <> False Then ThisWorkbook.Sheets(S.name).Cells(4, theCol) = S.total(theDate, shiftUnit)
            curDate = DateAdd("h", 8, curDate) 'increment
        Loop
        If currentSchedule.hasCharts Then
            If currentSchedule.getGraphs(S.name).has2ndAxis() Then currentSchedule.getGraphs(S.name).remove2ndAxis
        End If
    Next S
End If

exit_here:
Exit Sub

err_trap:
MsgBox "Error in ""ClearTrace"" of Runtime. Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Sub

Public Function getLatestVersion(Optional ver As Variant) As Date
Dim rs As ADODB.Recordset
Dim sql As String

On Error GoTo err_trap

updateConnection

If Not IsMissing(ver) Then
    sql = "SELECT ov.createdOn as newest FROM tbOperationDataVersions ov WHERE ov.operDataVerId=" & ver
Else
    sql = "SELECT Max(ov.createdOn) as newest FROM tbOperationDataVersions ov"
End If

Set rs = CreateObject("adodb.recordset")
rs.Open sql, adoConn
If Not rs.EOF Then
    rs.MoveFirst
    getLatestVersion = rs.Fields("newest")
End If
rs.Close

exit_here:
Set rs = Nothing
closeConnection
Exit Function

err_trap:
MsgBox "Error in ""getLatestVersion"" of Runtime. Error number: " & Err.Number & ", " & Err.Description
Resume exit_here


End Function

Public Function missingAdded() As String
Dim curProducts As New Collection
Dim compProducts As New Collection
Dim nProd As clsProduct
Dim dStr As String
Dim iStr As String
Dim uStr As String

Set curProducts = currentSchedule.getProducts
Set compProducts = comparativeSchedule.getProducts

'get all deleted products

For Each nProd In compProducts
    If Not inCollection(CStr(nProd.index), curProducts) Then
        If Len(dStr) = 0 Then
            dStr = "Usunięto następujące produkty:" & vbNewLine & "- " & nProd.toString & "; " & nProd.primaryAmount & " " & pUnit & "; " & nProd.secondaryAmount & " " & sUnit
        Else
            dStr = dStr & vbNewLine & "- " & nProd.toString & "; " & nProd.primaryAmount & " " & pUnit & "; " & nProd.secondaryAmount & " " & sUnit
        End If
    End If
Next nProd

'get all inserted products

For Each nProd In curProducts
    If Not inCollection(CStr(nProd.index), compProducts) Then
        If Len(iStr) = 0 Then
            iStr = vbNewLine & vbNewLine & "Dodano następujące produkty:" & vbNewLine & "- " & nProd.toString & "; " & nProd.primaryAmount & " " & pUnit & "; " & nProd.secondaryAmount & " " & sUnit
        Else
            iStr = iStr & vbNewLine & "- " & nProd.toString & "; " & nProd.primaryAmount & " " & pUnit & "; " & nProd.secondaryAmount & " " & sUnit
        End If
    End If
Next nProd

missingAdded = dStr & iStr

End Function

Public Sub detailForSelection()
If verifyRecordset Then
    currentSchedule.getDetailsForSelectedArea
End If
End Sub
