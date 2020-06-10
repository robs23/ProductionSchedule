VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'this will represent all data in current range
Private dFrom As Date
Private dTo As Date
Private pMode As Integer '1-prażenie, 2-mielenie, 3-pakowanie, 4-palety, 5-opakowania
Private pHasCharts As Boolean
Private pChartType As String 'dzienny / dzienny narastająco or zmianowy / zmianowy narastająco
Private pSplitBy As String
Private splitters As New Collection 'collection of worksheets in current workbook
Private records As New Collection 'collection of all records in all worksheets
Private operations As New Collection 'collection of all mes id numbers
Private graphs As New Collection 'collection of graph
Private products As New Collection 'collection of all products within this schedule
Private pVersionString As String
Private pVersionDate As Date

Public Property Let setDates(start As Date, finish As Date)
dFrom = start
dTo = finish
End Property

Public Property Get versionDate() As Date
versionDate = pVersionDate
End Property

Public Property Get versionString() As String
versionString = pVersionString
End Property

Public Property Get startDate() As Date
startDate = dFrom
End Property

Public Property Get endDate() As Date
endDate = dTo
End Property

Public Property Get splitBy() As String
splitBy = pSplitBy
End Property

Public Property Get mode() As Integer
mode = pMode
End Property

Public Property Let hasCharts(val As Boolean)
pHasCharts = val
End Property

Public Property Get hasCharts() As Boolean
hasCharts = pHasCharts
End Property

Public Property Let chartType(val As String)
pChartType = val
End Property

Public Property Get chartType() As String
chartType = pChartType
End Property

Public Sub initialize(start As Date, finish As Date, mode As Integer, splitBy As String, rs As ADODB.Recordset, versionDate As Date)

On Error GoTo err_trap

Dim nSplitter As clsSchduleSplitter
Dim i As Integer
Dim splitCol As String
Dim mFldName As String
Dim sFldName As String
Dim secUnit As Variant
Dim rBatch As Variant
Dim pStr As String
Dim unitRatio As Double
Dim ratioFld As String
Dim nProduct As clsProduct

dFrom = start
dTo = finish
pMode = mode
scheduleMode = mode
pSplitBy = splitBy
pVersionDate = versionDate


pVersionString = "Tydz. " & DatePart("ww", versionDate, vbMonday, vbFirstFourDays) & " " & WeekdayName(weekday(versionDate, vbMonday)) & " godz. " & DatePart("h", versionDate) & ":" & DatePart("n", versionDate)

If pUnit = "KG" Then
    mFldName = "KG"
ElseIf pUnit = "BOX" Then
    mFldName = "BOX"
ElseIf pUnit = "PC" Then
    mFldName = "PC"
Else
    mFldName = "matAmount"
End If

If sUnit = "KG" Then
    sFldName = "KG"
ElseIf sUnit = "BOX" Then
    sFldName = "BOX"
ElseIf sUnit = "PAL" Then
    sFldName = "PAL"
ElseIf sUnit = "Wsad" Then
    sFldName = ""
End If

rs.MoveFirst
Do Until rs.EOF
    If Not IsNull(rs.Fields("mesString")) Then newOperation rs.Fields("mesId"), rs.Fields("mesString") 'store mes operation number - needed in comparisons between different plan versions
    Set nSplitter = newSplitter(Left(Trim(rs.Fields(splitBy)), 30))
    If showProperties Then
        If rs.Fields("bean") = 1 Or rs.Fields("bean") = -1 Then
            pStr = "Z"
        ElseIf IsNull(rs.Fields("bean")) Then
            pStr = "*"
        ElseIf rs.Fields("bean") = 0 Then
            pStr = "M"
        End If
        If rs.Fields("decaf") = 1 Or rs.Fields("decaf") = -1 Then
            pStr = pStr & "-B"
        ElseIf IsNull(rs.Fields("decaf")) Then
            pStr = pStr & "-*"
        ElseIf rs.Fields("decaf") = 0 Then
            pStr = pStr & "-K"
        End If
    End If
    If InStr(1, mFldName & "_" & sFldName, "PC", vbTextCompare) > 0 And InStr(1, mFldName & "_" & sFldName, "PAL", vbTextCompare) > 0 Then
        ratioFld = "PC_PAL"
    ElseIf InStr(1, mFldName & "_" & sFldName, "PC", vbTextCompare) > 0 And InStr(1, mFldName & "_" & sFldName, "BOX", vbTextCompare) > 0 Then
        ratioFld = "PC_BOX"
    ElseIf InStr(1, mFldName & "_" & sFldName, "KG", vbTextCompare) > 0 And InStr(1, mFldName & "_" & sFldName, "PAL", vbTextCompare) > 0 Then
        ratioFld = "KG_PAL"
    ElseIf InStr(1, mFldName & "_" & sFldName, "KG", vbTextCompare) > 0 And InStr(1, mFldName & "_" & sFldName, "BOX", vbTextCompare) > 0 Then
        ratioFld = "KG_BOX"
    ElseIf InStr(1, mFldName & "_" & sFldName, "BOX", vbTextCompare) > 0 And InStr(1, mFldName & "_" & sFldName, "PAL", vbTextCompare) > 0 Then
        ratioFld = "BOX_PAL"
    ElseIf InStr(1, mFldName & "_" & sFldName, "PC", vbTextCompare) > 0 And InStr(1, mFldName & "_" & sFldName, "KG", vbTextCompare) > 0 Then
        ratioFld = "PC_KG"
    Else
        ratioFld = ""
    End If
    If sFldName = "" Then
        If mode = 1 Then
            If inCollection(rs.Fields("zfinIndex"), roastingBatches) Then
                If nSplitter.name = "RN3000" Then
                    rBatch = roastingBatches(Trim(CStr(rs.Fields("zfinIndex")))).rn3000input
                Else
                    rBatch = roastingBatches(Trim(CStr(rs.Fields("zfinIndex")))).rn4000input
                End If
                If IsNumeric(rBatch) And rBatch <> 0 Then
                    secUnit = Round(rs.Fields(mFldName) / rBatch, 1)
                Else
                    secUnit = 0
                End If
            Else
                secUnit = 0
            End If
            nSplitter.append rs.Fields("zfinIndex"), Trim(rs.Fields("zfinName")), DateSerial(year(rs.Fields("plMoment")), Month(rs.Fields("plMoment")), Day(rs.Fields("plMoment"))), rs.Fields("plShift"), rs.Fields(mFldName), CDbl(secUnit), Trim(rs.Fields("Maszyna")), , pStr, rBatch, rs.Fields("mesId"), rs.Fields("mesString")
        Else
            nSplitter.append rs.Fields("zfinIndex"), Trim(rs.Fields("zfinName")), DateSerial(year(rs.Fields("plMoment")), Month(rs.Fields("plMoment")), Day(rs.Fields("plMoment"))), rs.Fields("plShift"), rs.Fields(mFldName), CDbl(secUnit), Trim(rs.Fields("Maszyna")), , pStr, rBatch, rs.Fields("mesId"), rs.Fields("mesString"), rs.Fields("unit"), rs.Fields("iZfin"), rs.Fields("nZfin"), rs.Fields("Kategoria")
        End If
    Else
        If ratioFld <> "" Then
            nSplitter.append rs.Fields("zfinIndex"), Trim(rs.Fields("zfinName")), DateSerial(year(rs.Fields("plMoment")), Month(rs.Fields("plMoment")), Day(rs.Fields("plMoment"))), rs.Fields("plShift"), rs.Fields(mFldName), rs.Fields(sFldName), Trim(rs.Fields("Maszyna")), rs.Fields("Klient"), pStr, rs.Fields(ratioFld), rs.Fields("mesId"), rs.Fields("mesString")
        Else
            nSplitter.append rs.Fields("zfinIndex"), Trim(rs.Fields("zfinName")), DateSerial(year(rs.Fields("plMoment")), Month(rs.Fields("plMoment")), Day(rs.Fields("plMoment"))), rs.Fields("plShift"), rs.Fields(mFldName), rs.Fields(sFldName), Trim(rs.Fields("Maszyna")), rs.Fields("Klient"), , , rs.Fields("mesId"), rs.Fields("mesString")
        End If
    End If
    Select Case mode
        Case 1 'prażenie
            Set nProduct = newProduct(rs.Fields("zfinIndex"), rs.Fields("zfinName"))
            nProduct.append rs.Fields(mFldName), secUnit, rs.Fields("Maszyna")
        Case 3 'pakowanie
            Set nProduct = newProduct(rs.Fields("zfinIndex"), rs.Fields("zfinName"), rs.Fields("palType"), rs.Fields("Chep"))
            nProduct.append rs.Fields(mFldName), rs.Fields(sFldName), rs.Fields("Maszyna"), rs.Fields("Kierunek"), rs.Fields("Klient"), rs.Fields("loc")
        Case 4 'palety
            Set nProduct = newProduct(rs.Fields("zfinIndex"), rs.Fields("zfinName"), rs.Fields("palType"), rs.Fields("Chep"))
            nProduct.append rs.Fields(mFldName), rs.Fields(sFldName)
        Case 5 'opakowania
            Set nProduct = newProduct(rs.Fields("zfinIndex"), rs.Fields("zfinName"))
            nProduct.append rs.Fields(mFldName), 1, rs.Fields("Maszyna")
    End Select
    rs.MoveNext
Loop

Exit_here:
Exit Sub

err_trap:
MsgBox "Error in ""Initialize"" of clsSchedule. Error Number: " & Err.Number & ", " & Err.Description
Resume Exit_here

End Sub

Private Function newSplitter(name As String) As clsSchduleSplitter
Dim cSplitter As clsSchduleSplitter
Dim bool As Boolean

On Error GoTo err_trap

For Each cSplitter In splitters
    If cSplitter.name = name Then
        Set newSplitter = cSplitter
        bool = True
        Exit For
    End If
Next cSplitter

If Not bool Then
    Set newSplitter = New clsSchduleSplitter
    newSplitter.name = name
    splitters.add newSplitter, name
End If

Exit_here:
Exit Function

err_trap:
MsgBox "Error in ""newSplitter"" of clsSchedule. Error number: " & Err.Number & ", " & Err.Description
Resume Exit_here

End Function

Public Function getSplitters(Optional shtName As Variant) As Variant
If IsMissing(shtName) Then
    Set getSplitters = splitters
Else
    Set getSplitters = splitters(shtName)
End If
End Function

Public Function getRecords() As Collection
Set getRecords = records
End Function

Public Function getOperations() As Collection
Set getOperations = operations
End Function

Public Function getGraphs(Optional id As Variant) As Variant
If IsMissing(id) Then
    Set getGraphs = graphs
Else
    Set getGraphs = graphs(CStr(id))
End If
End Function

Public Function getProducts(Optional id As Variant) As Variant
If IsMissing(id) Then
    Set getProducts = products
Else
    Set getProducts = products(CStr(id))
End If
End Function

Public Sub newRecord(index As Long, name As String, sheet As String, location As Range)
Dim cRecord As clsRecord
Dim bool As Boolean

On Error GoTo err_trap

'For Each cRecord In records
'    If cRecord.index = index And cRecord.sheet = sheet Then
'        Set newRecord = cRecord
'        bool = True
'        Exit For
'    End If
'Next cRecord

If Not bool Then
    Set cRecord = New clsRecord
    cRecord.initialize index, name, sheet, location
    records.add cRecord, CStr(index & "_" & sheet)
End If

Exit_here:
Exit Sub

err_trap:
MsgBox "Error in ""newRecord"" of clsSchedule. Error number: " & Err.Number & ", " & Err.Description
Resume Exit_here

End Sub

Public Function filterRecordset(keyWord As String) As Collection
Dim cRecord As clsRecord
Dim fRecords As New Collection 'collection of filtered records


For Each cRecord In records
    If InStr(1, CStr(cRecord.index), keyWord, vbTextCompare) > 0 Then
        fRecords.add cRecord, CStr(cRecord.index & "_" & cRecord.sheet)
    ElseIf InStr(1, cRecord.name, keyWord, vbTextCompare) > 0 Then
        fRecords.add cRecord, CStr(cRecord.index & "_" & cRecord.sheet)
    End If
Next cRecord
    
Set filterRecordset = fRecords

End Function

Public Function total(splitName As String, theDate As Date, unit As String, Optional shift As Variant) As Double
Dim split As clsSchduleSplitter

For Each split In splitters
    If split.name = splitName Then
        total = split.total(theDate, unit, shift)
        Exit For
    End If
Next split

End Function

Public Sub createGraphs()
Dim splitter As clsSchduleSplitter
Dim graph As clsGraph
Dim title As String
Dim nTitle As String
Dim rng As Range
Dim unit As String
Dim regval As Variant
Dim curDate As Date
Dim sDate As Date 'short date of curDate
Dim shift As Integer
Dim sShift As String
Dim theTotal As Double
Dim xVal As Variant
Dim yVal As Variant
Dim mode As Integer '0-shiftly,1-daily
Dim cumul As Integer '0-without cumulation, 1-WITH cumulation
Dim step As Integer 'number of hours to increment by. 8 in SHIFTLY mode, 24 in DAILY mode

On Error GoTo err_trap

regval = registryKeyExists(regPath & "Charts")
pHasCharts = True
pChartType = regval

If InStr(1, regval, "narastająco", vbTextCompare) > 0 Then
    title = "Zbiorcza ilość " & sUnit
    cumul = 1
Else
    title = "Ilość " & sUnit
    cumul = 0
End If
If InStr(1, regval, "Zmianowy", vbTextCompare) > 0 Then
    title = title & " na zmianę dla <splitter>"
    mode = 0
    step = 8
Else
    title = title & " na dzień dla <splitter>"
    mode = 1
    step = 24
End If

regval = registryKeyExists(regPath & "dateRangeType")
If regval = "Weekly" Then
    title = title & " w tygodniu "
    regval = registryKeyExists(regPath & "LoadedWeek")
    title = title & regval & "|"
    regval = registryKeyExists(regPath & "LoadedYear")
    title = title & regval
ElseIf regval = "Custom" Then
    title = title & " w okresie "
    regval = registryKeyExists(regPath & "CustomRangeStartDate")
    title = title & regval & "|"
    regval = registryKeyExists(regPath & "CustomRangeEndDate")
    title = title & regval
End If

For Each splitter In splitters
    theTotal = 0
    Set graph = New clsGraph
    nTitle = Replace(title, "<splitter>", splitter.name)
    Set rng = ThisWorkbook.Sheets(splitter.name).Range(ThisWorkbook.Sheets(splitter.name).Cells(splitter.lastDataRow + 1, splitter.firstDataCol), ThisWorkbook.Sheets(splitter.name).Cells(splitter.lastDataRow + 25, splitter.firstDataCol + 16))
    If cumul = 1 Then
        graph.initialize nTitle, splitter.name & "_" & sUnit, xlLine, "Czas", "Ilość " & sUnit & ".", splitter.name, rng
    ElseIf cumul = 0 Then
        graph.initialize nTitle, splitter.name & "_" & sUnit, xlColumnClustered, "Czas", "Ilość " & sUnit & ".", splitter.name, rng
    End If
    Select Case DatePart("h", dFrom)
    Case 6, 14, 22
        curDate = dFrom
    Case Else
        curDate = DateAdd("h", 6, DateSerial(year(dFrom), Month(dFrom), Day(dFrom)))
    End Select
    
    Do Until curDate > dTo
        Select Case CInt(DatePart("h", curDate))
            Case 6
            shift = 1
            sShift = "I"
        Case 14
            shift = 2
            sShift = "II"
        Case 22
            shift = 3
            sShift = "III"
        End Select
        sDate = DateSerial(year(curDate), Month(curDate), Day(curDate))
        If mode = 0 Then
            'show day AND shift in shiftly mode
            xVal = WeekdayName(weekday(curDate, vbMonday)) & " " & sShift
            If cumul = 0 Then
                theTotal = splitter.total(sDate, sUnit, shift)
            Else
                theTotal = theTotal + splitter.total(sDate, sUnit, shift)
            End If
        Else
            xVal = WeekdayName(weekday(curDate, vbMonday))
            If cumul = 0 Then
                theTotal = splitter.total(sDate, sUnit)
            Else
                theTotal = theTotal + splitter.total(sDate, sUnit)
            End If
        End If
        graph.append xVal, theTotal
        curDate = DateAdd("h", step, curDate)
    Loop
    graph.createChart
    graphs.add graph, graph.sheetName
Next splitter

Exit_here:
Exit Sub

err_trap:
MsgBox "Error in ""createGraphs"" of clsSchedule. Error number: " & Err.Number & ", " & Err.Description
Resume Exit_here

End Sub

Public Sub newOperation(id As Long, mesStr As String)

Dim cOperation As clsOperation
Dim bool As Boolean

On Error GoTo err_trap

For Each cOperation In operations
    If cOperation.operationId = id Then
        bool = True
        Exit For
    End If
Next cOperation

If Not bool Then
    Set cOperation = New clsOperation
    cOperation.operationId = id
    cOperation.mesString = mesStr
    operations.add cOperation, CStr(id)
End If

Exit_here:
Exit Sub

err_trap:
MsgBox "Error in ""newOperation"" of clsSchedule. Error number: " & Err.Number & ", " & Err.Description
Resume Exit_here

End Sub

Public Function highestOperationId() As Long
Dim cOp As clsOperation

highestOperationId = 0

For Each cOp In operations
    If cOp.operationId > highestOperationId Then
        highestOperationId = cOp.operationId
    End If
Next cOp

End Function

Public Function operationExists(arg As Variant) As Boolean
Dim cOp As clsOperation

If VarType(arg) = vbString Then
    'operation's mes string has been passed in
    operationExists = False
    
    For Each cOp In operations
        If cOp.mesString = arg Then
            operationExists = True
            Exit For
        End If
    Next cOp
Else
    'operation's id has been passed in
    operationExists = inCollection(CStr(arg), operations)
End If

End Function

Private Function newProduct(index As Long, name As String, Optional palletType As Variant, Optional palletChep As Variant) As clsProduct
Dim nProduct As clsProduct
Dim bool As Boolean

bool = False

If inCollection(CStr(index), products) Then
    bool = True
    Set nProduct = products(CStr(index))
End If

If bool = False Then
    Set nProduct = New clsProduct
    nProduct.initialize index, name, palletType, palletChep
    products.add nProduct, CStr(index)
End If

Set newProduct = nProduct

End Function

Public Sub getDetailsForSelectedArea()
Dim sht As Worksheet
Dim prods As New Collection
Dim prod As clsProduct
Dim str As String
Dim splitter As clsSchduleSplitter
Dim item As clsSplitterItem
Dim shift As clsShift
Dim shiftStr As String
Dim dets As New Collection
Dim det As clsDetails

If isNothing(Selection) Then
    MsgBox "Najpierw zaznacz komórkę / grupę komórek", vbInformation + vbOKOnly, "Brak zaznaczenia"
Else
    Set sht = ActiveCell.Worksheet
    Set prods = getProductsByLocation(Selection)
    If prods.Count > 0 Then
        If prods.Count = 1 Then
            Set splitter = splitters(sht.name)
            Set item = splitter.getItems(CStr(prods(1).index))
            shiftStr = getSelectedShifts
            For Each shift In item.getShifts
                If InStr(1, shiftStr, shift.shiftName, vbTextCompare) Then
                    For Each det In shift.getDetails
                        If inCollection(det.name, dets) Then
                            dets(det.name).MainUnit = dets(det.name).MainUnit + det.MainUnit
                        Else
                            dets.add det, det.name
                        End If
                    Next det
                End If
            Next shift
            For Each det In dets
                str = str & det.toString & vbNewLine
            Next det
        Else
            str = "Zbyt wiele indeksów w zaznaczeniu! Zaznacz obszar w granicach 1 indeksu"
        End If
    Else
        str = "Brak produktów w zaznaczeniu"
    End If
    If str <> "" Then MsgBox str
End If
End Sub

Public Function getProductsByLocation(rng As Range) As Collection
Dim prods As New Collection
Dim rec As clsRecord
Dim prod As clsProduct
Dim c As Range

For Each c In rng
    For Each rec In records
        If rec.sheet = rng.Worksheet.name Then
            If rec.location.Row = c.Row Or rec.location.Row = c.Row - 1 Then
                If Not inCollection(CStr(rec.index), prods) Then
                    Set prod = New clsProduct
                    prod.index = rec.index
                    prods.add prod, CStr(prod.index)
                End If
                Exit For
            End If
        End If
    Next rec
Next c

Set getProductsByLocation = prods

End Function

Public Function getSelectedShifts() As String
Dim str As String
Dim c As Range
Dim sht As Worksheet
Dim shift As Integer
Dim d As String

If Not isNothing(Selection) Then
    Set sht = Selection.Worksheet
    For Each c In Selection
        If sht.Cells(3, c.Column) = "I" Or sht.Cells(3, c.Column) = "II" Or sht.Cells(3, c.Column) = "III" Then
            Select Case sht.Cells(3, c.Column)
                Case Is = "I"
                shift = 1
                Case Is = "II"
                shift = 2
                Case Is = "III"
                shift = 3
            End Select
            d = sht.Cells(2, c.Column - shift + 1)
            If InStr(1, str, CStr(d & "_" & shift), vbTextCompare) = 0 Then
                str = str & d & "_" & shift & ","
            End If
        End If
    Next c
End If

If Len(str) > 0 Then str = Left(str, Len(str) - 1)
getSelectedShifts = str

End Function
