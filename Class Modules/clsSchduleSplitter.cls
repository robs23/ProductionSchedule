VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSchduleSplitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'this will represent each worksheet
Private pName As String 'this will become worksheet's name
Private pItems As New Collection 'collection of rows in worksheet
Private shiftTotals As New Collection 'total of production on given shift from all products
Private lastRow As Long
Private lastCol As Long
Private firstRow As Long
Private firstCol As Long
Private totalRng As Range
Private bodyRng As Range
Private headerRng As Range
Private dataRng As Range

Public Property Let name(theName As String)
pName = theName
End Property

Public Property Get name() As String
name = pName
End Property

Public Property Get firstDataCol() As Long 'number of first column after item description
firstDataCol = firstCol
End Property

Public Property Get lastDataRow() As Long 'number of last row in dataRng
lastDataRow = lastRow
End Property

Private Function newItem(index As Long) As clsSplitterItem
Dim bool As Boolean
Dim cItem As clsSplitterItem

On Error GoTo err_trap

For Each cItem In pItems
    If cItem.index = index Then
        Set newItem = cItem
        bool = True
        Exit For
    End If
Next cItem

If Not bool Then
    Set newItem = New clsSplitterItem
    'newItem.initialize
    pItems.add newItem, CStr(index)
End If

Exit_here:
Exit Function

err_trap:
MsgBox "Error in ""newItem"" of clsSchduleSplitter. Error number: " & Err.Number & ", " & Err.Description
Resume Exit_here

End Function

Public Sub append(index As Long, name As String, sDate As Date, sNumber As Integer, mUnit As Double, sUnit As Double, machName As String, Optional cust As Variant, Optional properties As Variant, Optional unitRatio As Variant, Optional operationId As Variant, Optional mesString As Variant, Optional materialUnit As Variant, Optional parentIndex As Variant, Optional parentName As Variant, Optional materialType As Variant)
Dim cItem As clsSplitterItem
Dim cTotal As clsTotalOnShift

Set cTotal = newTotalOnShift(CStr(sDate) & "_" & sNumber)
cTotal.add mUnit, sUnit
Set cItem = newItem(index)
With cItem
    .initialize index, name, properties, cust, unitRatio, materialUnit, materialType
    If currentSchedule.mode < 5 Then
        .append sDate, sNumber, mUnit, sUnit, CLng(operationId), mesString, machName, index, name
    Else
        .append sDate, sNumber, mUnit, sUnit, CLng(operationId), mesString, machName, parentIndex, parentName
    End If
End With
End Sub

Public Sub deployResults()
Dim sht As Worksheet
Dim c As Range
Dim cItem As clsSplitterItem
Dim cShift As clsShift
Dim i As Integer
Dim x As Integer
Dim shifts As Collection
Dim col As Variant
Dim cm As Comment

Set sht = ThisWorkbook.Sheets(pName)

For Each c In sht.Range("A1:A10")
    If c.MergeCells = False Then
        i = c.Row
        Exit For
    End If
Next c

For Each cItem In pItems
    sht.Range(sht.Cells(i, 1), sht.Cells(i + 1, 1)).Merge
    sht.Cells(i, 1) = cItem.index
    If showComments Then
        Set cm = sht.Cells(i, 1).AddComment
        cm.Text cItem.toString
        With cm.Shape.TextFrame
            .Characters(InStr(1, cm.Text, "Index:", vbTextCompare), 6).Font.Bold = True
            .Characters(InStr(1, cm.Text, "Nazwa:", vbTextCompare), 6).Font.Bold = True
            .Characters(InStr(1, cm.Text, "Maszyna:", vbTextCompare), 8).Font.Bold = True
            .Characters(InStr(1, cm.Text, "Klient:", vbTextCompare), 7).Font.Bold = True
            .Characters(InStr(1, cm.Text, "Kategoria:", vbTextCompare), 10).Font.Bold = True
            .Characters(InStr(1, cm.Text, "Zfin:", vbTextCompare), 5).Font.Bold = True
            .AutoSize = True
        End With
    End If
    currentSchedule.newRecord cItem.index, cItem.name, pName, sht.Cells(i, 1)
    x = 2
    If showName Then
        sht.Range(sht.Cells(i, x), sht.Cells(i + 1, x)).Merge
        sht.Cells(i, x) = cItem.name
        x = x + 1
    End If
    If showCustomer Then
        sht.Range(sht.Cells(i, x), sht.Cells(i + 1, x)).Merge
        sht.Cells(i, x) = cItem.customer
        x = x + 1
    End If
    If showMachine Then
        sht.Range(sht.Cells(i, x), sht.Cells(i + 1, x)).Merge
        sht.Cells(i, x) = cItem.machine
        x = x + 1
    End If
    If showProperties Then
        sht.Range(sht.Cells(i, x), sht.Cells(i + 1, x)).Merge
        sht.Cells(i, x) = cItem.properties
        x = x + 1
    End If
    If showUnitRatio Then
        sht.Range(sht.Cells(i, x), sht.Cells(i + 1, x)).Merge
        sht.Cells(i, x) = cItem.unitRatio
        x = x + 1
    End If
    If cItem.materialUnit <> "" Then
        sht.Cells(i, x) = cItem.materialUnit
    Else
        sht.Cells(i, x) = pUnit
    End If
    sht.Cells(i + 1, x) = sUnit
    x = x + 1
    If showTotalIndex Then
        sht.Cells(i, x) = cItem.getTotal()(0)
        sht.Cells(i + 1, x) = cItem.getTotal()(1)
        x = x + 1
    End If
    Set shifts = cItem.getShifts
    For Each cShift In shifts
        col = getShiftColumn(cShift.shiftDate, cShift.shiftNumber)
        If Not IsNull(col) Then
            sht.Cells(i, col) = cShift.MainUnit
            If cShift.secUnit > 0 Then sht.Cells(i + 1, col) = cShift.secUnit
        End If
    Next cShift
    i = i + 2
Next cItem
printTotals
formatMe
End Sub

Public Function getShiftColumn(theDate As Date, theShift As Integer) As Variant
Dim sht As Worksheet
Dim rng As Range
Dim c As Range
Dim x As Integer
Dim y As Integer

On Error GoTo err_trap

Set sht = ThisWorkbook.Sheets(pName)
Set rng = sht.Range(headerAddress)
Set c = rng.Find(theDate, SearchOrder:=xlByRows, SearchDirection:=xlNext, LookIn:=xlFormulas)

If Not c Is Nothing Then
    x = c.Column + theShift - 1
    getShiftColumn = x
Else
    getShiftColumn = Null
End If

Exit_here:
Exit Function

err_trap:
MsgBox "Error in ""GetShiftColumn"" of clsSchduleSplitter. Error number: " & Err.Number & ", " & Err.Description
Resume Exit_here

End Function

Private Function newTotalOnShift(ref As String) As clsTotalOnShift
Dim bool As Boolean
Dim cShift As clsTotalOnShift

On Error GoTo err_trap

For Each cShift In shiftTotals
    If cShift.name = ref Then
        Set newTotalOnShift = cShift
        bool = True
        Exit For
    End If
Next cShift

If Not bool Then
    Set newTotalOnShift = New clsTotalOnShift
    newTotalOnShift.name = ref
    shiftTotals.add newTotalOnShift, ref
End If

Exit_here:
Exit Function

err_trap:
MsgBox "Error in ""newTotalOnShift"" of clsSchduleSplitter. Error number: " & Err.Number & ", " & Err.Description
Resume Exit_here

End Function

Public Function total(theDate As Date, unit As String, Optional shift As Variant, Optional comparativeSchedule As Variant) As Double
Dim tot As Double
Dim val As Variant
Dim cItem As clsSplitterItem
Dim cShift As clsShift

On Error GoTo err_trap

theDate = DateSerial(year(theDate), Month(theDate), Day(theDate))

If IsMissing(comparativeSchedule) Then
    If IsMissing(shift) Then
        'we need total of the day
        If inCollection(CStr(theDate) & "_" & 1, shiftTotals) Then
            If unit = "m" Then
                val = shiftTotals(CStr(theDate) & "_" & 1).mainTotal
            Else
                val = shiftTotals(CStr(theDate) & "_" & 1).secTotal
            End If
            tot = tot + val
        End If
        If inCollection(CStr(theDate) & "_" & 2, shiftTotals) Then
            If unit = "m" Then
                val = shiftTotals(CStr(theDate) & "_" & 2).mainTotal
            Else
                val = shiftTotals(CStr(theDate) & "_" & 2).secTotal
            End If
            tot = tot + val
        End If
        If inCollection(CStr(theDate) & "_" & 3, shiftTotals) Then
            If unit = "m" Then
                val = shiftTotals(CStr(theDate) & "_" & 3).mainTotal
            Else
                val = shiftTotals(CStr(theDate) & "_" & 3).secTotal
            End If
            tot = tot + val
        End If
    Else
        If inCollection(CStr(theDate) & "_" & shift, shiftTotals) Then
            If unit = "m" Then
                val = shiftTotals(CStr(theDate) & "_" & shift).mainTotal
            Else
                val = shiftTotals(CStr(theDate) & "_" & shift).secTotal
            End If
            tot = tot + val
        End If
    End If
Else
    If IsMissing(shift) Then
        'daily
        For Each cItem In pItems
            For Each cShift In cItem.getShifts()
                If cShift.shiftDate = theDate Then
                    'this item has appropriate shift
                    If comparativeSchedule.operationExists(cShift.mesString) Then
                        'this operation is present in comparative schedule so it can be counted in comparison
                        If unit = "m" Then
                            val = cShift.MainUnit
                        Else
                            val = cShift.secUnit
                        End If
                        tot = tot + val
                    End If
                End If
            Next cShift
        Next cItem
    Else
        'desired shift is given
        For Each cItem In pItems
            For Each cShift In cItem.getShifts()
                If cShift.shiftDate = theDate And cShift.shiftNumber = shift Then
                    'this item has appropriate shift
                    If comparativeSchedule.operationExists(cShift.mesString) Then
                        'this operation is present in comparative schedule so it can be counted in comparison
                        If unit = "m" Then
                            val = cShift.MainUnit
                        Else
                            val = cShift.secUnit
                        End If
                        tot = tot + val
                    End If
                    Exit For
                End If
            Next cShift
        Next cItem
    End If
End If

total = tot

Exit_here:
Exit Function

err_trap:
MsgBox "Error in ""Total"" of clsSchduleSplitter. Error number: " & Err.Number & ", " & Err.Description
Resume Exit_here

End Function

Private Sub printTotals()

Dim c As Range
Dim d As Range
Dim sht As Worksheet
Dim theDate As Date
Dim tot As Double
Dim shift As Integer
Dim theRow As Integer

Set sht = ThisWorkbook.Sheets(pName)

firstRow = 4

If TotalDaily <> False Or TotalShift <> False Then
    If TotalDaily <> False Then firstRow = firstRow + 1
    If TotalShift <> False Then firstRow = firstRow + 1
    For Each c In sht.Range("A3:H3")
        If c.MergeCells = False Then
            Set c = sht.Cells(2, c.Column)
            Set d = c
            Exit For
        End If
    Next c
    Do Until CStr(c) = ""
        theDate = c
        If TotalDaily <> False Then
            If scheduleMode = 1 Then
                If TotalDaily = "Wsad" Then
                    tot = Me.total(theDate, "s")
                Else
                    tot = Me.total(theDate, "m")
                End If
                sht.Cells(4, c.Column) = tot
            ElseIf scheduleMode = 3 Or scheduleMode = 4 Then
                If TotalDaily = sUnit Then
                    tot = Me.total(theDate, "s")
                ElseIf TotalDaily = pUnit Then
                    tot = Me.total(theDate, "m")
                End If
                sht.Cells(4, c.Column) = tot
            End If
        End If
        Set c = c.Offset(0, 1)
    Loop
    Set c = d
    If TotalDaily <> False Then theRow = 5 Else theRow = 4
    Do Until CStr(c) = ""
        theDate = c
        If TotalShift <> False Then
            If scheduleMode = 1 Then
                If TotalShift = "Wsad" Then
                    sht.Cells(theRow, c.Column) = Me.total(theDate, "s", 1)
                    sht.Cells(theRow, c.Column + 1) = Me.total(theDate, "s", 2)
                    sht.Cells(theRow, c.Column + 2) = Me.total(theDate, "s", 3)
                Else
                    sht.Cells(theRow, c.Column) = Me.total(theDate, "m", 1)
                    sht.Cells(theRow, c.Column + 1) = Me.total(theDate, "m", 2)
                    sht.Cells(theRow, c.Column + 2) = Me.total(theDate, "m", 3)
                End If
            ElseIf scheduleMode = 3 Or scheduleMode = 4 Then
                If TotalShift = sUnit Then
                    sht.Cells(theRow, c.Column) = Me.total(theDate, "s", 1)
                    sht.Cells(theRow, c.Column + 1) = Me.total(theDate, "s", 2)
                    sht.Cells(theRow, c.Column + 2) = Me.total(theDate, "s", 3)
                ElseIf TotalShift = pUnit Then
                    sht.Cells(theRow, c.Column) = Me.total(theDate, "m", 1)
                    sht.Cells(theRow, c.Column + 1) = Me.total(theDate, "m", 2)
                    sht.Cells(theRow, c.Column + 2) = Me.total(theDate, "m", 3)
                End If
            End If
        End If
        Set c = c.Offset(0, 1)
    Loop
End If

End Sub

Public Sub formatMe()
Dim c As Range
Dim sht As Worksheet
Dim x As Integer
Dim y As Integer
Dim letter As String

Set sht = ThisWorkbook.Sheets(pName)
Set c = sht.Range("A:A").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)

If Not c Is Nothing Then
    lastRow = c.Row + 1
End If

Set c = sht.Range("A1:BOR1").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)

If Not c Is Nothing Then
    lastCol = c.Column + 2
End If

Set bodyRng = sht.Range(c, sht.Cells(lastRow, lastCol))

Set c = sht.Range("A3")
Do Until c.MergeCells = False
    Set c = c.Offset(0, 1)
Loop

firstCol = c.Column

sht.Select
With ActiveWindow
    .Zoom = 90
    .SplitColumn = firstCol - 1
    .SplitRow = firstRow - 1
    .FreezePanes = True
End With


Set dataRng = sht.Range(c, sht.Cells(lastRow, lastCol))
Set headerRng = sht.Range(sht.Cells(1, 1), sht.Cells(firstRow - 1, lastCol))

If lastRow > 0 And lastCol > 0 Then
    Set totalRng = sht.Range(sht.Cells(1, 1), sht.Cells(lastRow, lastCol))
    totalRng.Interior.Color = vbWhite
    totalRng.Cells.HorizontalAlignment = xlCenter
    totalRng.Cells.RowHeight = 15
    totalRng.Cells.ColumnWidth = 6
    sht.Range("A:A").ColumnWidth = 10
    If showName Then
        sht.Range("B:B").ColumnWidth = 25
        sht.Range("B:B").HorizontalAlignment = xlLeft
    End If
    If showCustomer Then
        letter = columnPosition("Klient", False)
        sht.Range(letter & ":" & letter).ColumnWidth = 15
        sht.Range(letter & ":" & letter).HorizontalAlignment = xlLeft
    End If
    If showMachine Then
        letter = columnPosition("Maszyna", False)
        sht.Range(letter & ":" & letter).ColumnWidth = 10
        sht.Range(letter & ":" & letter).HorizontalAlignment = xlCenter
    End If
    With totalRng.Cells.Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThin
    End With
    For y = firstRow To lastRow Step 2
        sht.Range(sht.Cells(y, firstCol), sht.Cells(y, lastCol)).Borders(xlEdgeBottom).LineStyle = xlLineStyleNone
    Next y
End If

End Sub

Public Property Get workingRange() As Range
Set workingRange = dataRng
End Property

Public Property Get getHeader() As Range
Set getHeader = headerRng
End Property

Public Property Get getTotalRange() As Range
Set getTotalRange = totalRng
End Property

Public Property Get getLastCol() As Integer
getLastCol = lastCol
End Property

Public Property Get getLastRow() As Integer
getLastRow = lastRow
End Property


Private Function columnPosition(columnName As String, returnColumnId As Boolean) As Variant
Dim c As Range

For Each c In headerRng
    If c.value = columnName Then
        If returnColumnId = True Then
            'returned will be column number
            columnPosition = c.Column
        Else
            'returned will be column letter
            columnPosition = cell2letter(c.Column)
        End If
    End If
Next c


End Function

Public Function getItems(Optional index As Variant) As Variant
If IsMissing(index) Then
    Set getItems = pItems
Else
    If inCollection(CStr(index), pItems) Then
        Set getItems = pItems(index)
    Else
        Set getItems = Null
    End If
End If
End Function
