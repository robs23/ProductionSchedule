VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} tracer 
   Caption         =   "Rysuj ślad"
   ClientHeight    =   1320
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   3750
   OleObjectBlob   =   "tracer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "tracer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private dFrom As Date
Private dTo As Date
Private isCreated As Boolean
Private xValues() As Variant 'array of xValues for chart
Private yValues() As Variant 'array of yValues for chart

Private Sub loadVersions()
Dim sql As String
Dim v() As String
Dim dday As Integer
Dim hhour As Integer
Dim i As Integer
Dim rs As ADODB.Recordset
Dim rngStart As String
Dim regval As Variant
Dim w As Integer
Dim y As Integer
Dim bool As Boolean

On Error GoTo err_trap

bool = False

regval = registryKeyExists(regPath & "RangeStart")
If regval <> False And regval <> "" Then
    rngStart = regval
    regval = registryKeyExists(regPath & "LoadedWeek")
    If regval <> False And regval <> 0 Then
        w = regval
        regval = registryKeyExists(regPath & "LoadedYear")
        If regval <> False And regval <> 0 Then
            y = regval
            bool = True
        End If
    End If
End If

If bool Then
    updateConnection
    v = split(rngStart, " ", , vbTextCompare)
    dday = WeekDayName2Int(v(0))
    hhour = CInt(Left(v(1), 2))
    dFrom = DateAdd("h", hhour, Week2Date(CLng(w), CLng(y), dday, vbFirstFourDays))
    dTo = DateAdd("h", 167, dFrom)
    'Else
    '    dFrom = DateAdd("h", CDbl(Left(Me.cmbTFrom, 2)), Me.txtFrom)
    '    dTo = DateAdd("h", CDbl(Left(Me.cmbTTo, 2)), Me.txtTo)
    
    sql = "SET DATEFIRST 1; " _
        & "SELECT oh.operDataVer, ov.createdOn, 'W' + CAST(DATEPART(ISO_WEEK,ov.createdOn) as varchar) + '_' +  DATENAME(dw,ov.createdOn) + '_' + CAST(DATEPART(HOUR,ov.createdOn) as varchar) + ':' + CAST(DATEPART(MINUTE,ov.createdOn) as varchar) as dayOfWeek, " _
        & "CAST(COUNT(oh.operationId) as float) / CAST((SELECT COUNT(ohAll.operationId) FROM tbOperationData ohAll WHERE ohAll.plMoment BETWEEN '" & dFrom & "' AND '" & dTo & "') AS float) as perc " _
        & "FROM tbOperationDataHistory oh JOIN tbOperationDataVersions ov ON ov.operDataVerId=oh.operDataVer " _
        & "WHERE oh.plMoment BETWEEN '" & dFrom & "' AND '" & dTo & "' " _
        & "GROUP BY oh.operDataVer, ov.createdOn, 'W' + CAST(DATEPART(ISO_WEEK,ov.createdOn) as varchar) + '_' +  DATENAME(dw,ov.createdOn) + '_' + CAST(DATEPART(HOUR,ov.createdOn) as varchar) + ':' + CAST(DATEPART(MINUTE,ov.createdOn) as varchar) " _
        & "HAVING CAST(COUNT(oh.operationId) as float) / CAST((SELECT COUNT(ohAll.operationId) FROM tbOperationData ohAll WHERE ohAll.plMoment BETWEEN '" & dFrom & "' AND '" & dTo & "') AS float) > 0.8 " _
        & "ORDER BY ov.createdOn DESC"
    
    Set rs = CreateObject("adodb.recordset")
    rs.Open sql, adoConn
    If Not rs.EOF Then
        For i = Me.cmbVersions.ListCount - 1 To 0 Step -1
            Me.cmbVersions.RemoveItem i
        Next i
        
        i = 0
        
        Do Until rs.EOF
            Me.cmbVersions.AddItem rs.Fields("operDataVer")
            Me.cmbVersions.Column(1, i) = rs.Fields("dayOfWeek")
            Me.cmbVersions.Column(2, i) = rs.Fields("createdOn")
            i = i + 1
            rs.MoveNext
        Loop
        Me.cmbVersions.ColumnCount = 3
        Me.cmbVersions.ColumnWidths = "0 pt;85 pt; 85 pt"
    End If
    
    rs.Close
Else
    MsgBox "Coś poszło nie tak. Spróbuj najpierw aktualizować harmonogram i następnie skorzystaj z rysowania śladu", vbInformation + vbOKOnly, "Coś nie pykło"
End If

exit_here:
Set rs = Nothing
closeConnection
Exit Sub

err_trap:
MsgBox "Error in ""loadVersions"" of tracer. Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Sub

Private Sub restoreVersion()
 For i = Me.cmbVersions.ListCount - 1 To 0 Step -1
    Me.cmbVersions.RemoveItem i
Next i

End Sub


Private Sub btnOK_Click()
If Me.cmbVersions <> "" Then
    getPlan
Else
    MsgBox "Z rozwijanej listy wybierz wersję planu, której ślad chcesz narysować", vbOKOnly + vbInformation, "Brak wyboru"
End If
End Sub

Private Sub UserForm_Initialize()
loadVersions
End Sub

Private Sub getPlan()
Dim regval As Variant
Dim zfinList As Variant
Dim bool As Boolean
Dim sql As String
Dim rs As ADODB.Recordset
Dim index As String
Dim split As String
Dim records As New Collection
Dim theSheet As String
Dim theRow As Integer
Dim theCol As Variant
Dim theMode As Integer '1-prażenie,2-mielenie,3-pakowanie
Dim c As Range
Dim theDate As Date
Dim theShift As Integer
Dim sShift As String
Dim currentSplitters As New Collection
Dim splitter As clsSchduleSplitter
Dim nextDaily As Double
Dim compShift As Double
Dim compDaily As Double
Dim totCurrentShift As Double
Dim totComparativeShift As Double
Dim totCurrentDaily As Double
Dim totComparativeDaily As Double
Dim dailyUnit As String 'the unit used as total for daily summary
Dim shiftUnit As String 'the unit used as total for shift summary
Dim shiftRow As Integer 'either 4 or 5, depends on if daily summary is displayed above
Dim sign As String
Dim xVal As Variant
Dim daily As Boolean
Dim cumul As Boolean
Dim updatedOn As Date
Dim changes As String

On Error GoTo err_trap

Application.ScreenUpdating = False
Application.StatusBar = "Przetwarzam dane.. Proszę czekać.."

updatedOn = CDate(Me.cmbVersions.Column(2, Me.cmbVersions.ListIndex))

bool = False

regval = registryKeyExists(regPath & "Type")
If regval <> False And regval <> "" Then
    If regval = "Prażenie" Then
        theMode = 1
        sql = "SELECT o.mesId,o.mesString, z.zfinIndex,z.zfinName, od.plMoment,od.plShift, m.machineName as Maszyna, od.plAmount as KG, o.type, zp.[beans?] as bean, zp.[decafe?] as decaf, 'Wszystkie' as brak " _
            & "FROM tbOperations o LEFT JOIN tbOperationDataHistory od ON od.operationId=o.operationId LEFT JOIN tbZfin z ON z.zfinId=o.zfinId LEFT JOIN tbMachine m ON m.machineId=od.plMach LEFT JOIN tbZfinProperties zp ON zp.zfinId=z.zfinId " _
            & "WHERE od.plMoment >= '" & dFrom & "' AND od.plMoment < '" & dTo & "' AND o.type = 'r' AND od.operDataVer = " & Me.cmbVersions
    ElseIf regval = "Pakowanie" Then
        theMode = 3
        sql = "SELECT o.mesId,o.mesString, z.zfinIndex,z.zfinName, od.plMoment,od.plShift, m.machineName as Maszyna, cs.custString as Klient, od.plAmount as PC, ROUND(od.plAmount*u.unitWeight,1) as KG, ROUND(od.plAmount/u.pcPerBox,1) AS BOX, ROUND(od.plAmount/u.pcPerPallet,1) AS PAL ,o.type, zp.[beans?] as bean, zp.[decafe?] as decaf,CASE WHEN cs.custString IS NOT NULL THEN LEFT(cs.custString,2) ELSE 'UNKNOWN' END as Kierunek,CASE WHEN p.palletLength+p.palletWidth = 2000 THEN 'EUR' ELSE CASE WHEN p.palletLength+p.palletWidth = 2200 THEN 'IND' ELSE 'UNKNOWN' END END AS palType, p.palletChep AS Chep, " _
            & "u.pcPerPallet as PC_PAL, u.pcPerBox as PC_BOX, ROUND(1/u.unitWeight,1) as PC_KG, ROUND(u.pcPerPallet*u.unitWeight,1) as KG_PAL, ROUND(u.pcPerBox*u.unitWeight,1) as KG_BOX, u.pcPerPallet/u.pcPerBox as BOX_PAL, 'Wszystkie' as brak, cs.location as loc " _
            & "FROM tbOperations o LEFT JOIN tbOperationDataHistory od ON od.operationId=o.operationId LEFT JOIN tbZfin z ON z.zfinId=o.zfinId LEFT JOIN tbMachine m ON m.machineId=od.plMach LEFT JOIN tbZfinProperties zp ON zp.zfinId=z.zfinId LEFT JOIN tbUom u ON u.zfinId=z.zfinId LEFT JOIN tbCustomerString cs ON cs.custStringId = z.custString LEFT JOIN tbPallets p ON p.palletId=u.palletType " _
            & "WHERE od.plMoment >= '" & dFrom & "' AND od.plMoment < '" & dTo & "' AND o.type = 'p' And od.operDataVer = " & Me.cmbVersions
    ElseIf regval = "Palety" Then
        theMode = 4
        sql = "SELECT o.mesId,o.mesString, p.palletId as zfinIndex,p.palletName as zfinName,od.plMoment,od.plShift,m.machineName as Maszyna, cs.custString as Klient, 'Palety' as Palety, od.plAmount as PC,ROUND(od.plAmount*u.unitWeight,1) as KG, ROUND(od.plAmount/u.pcPerBox,1) AS BOX, ROUND(od.plAmount/u.pcPerPallet,1) AS PAL ,o.type, zp.[beans?] as bean,zp.[decafe?] as decaf, " _
            & "CASE WHEN cs.custString IS NOT NULL THEN LEFT(cs.custString,2) ELSE 'UNKNOWN' END as Kierunek,CASE WHEN p.palletLength+p.palletWidth = 2000 THEN 'EUR' ELSE CASE WHEN p.palletLength+p.palletWidth = 2200 THEN 'IND' ELSE 'UNKNOWN' END END AS palType, p.palletChep AS Chep, u.pcPerPallet as PC_PAL, u.pcPerBox as PC_BOX, ROUND(1/u.unitWeight,1) as PC_KG, ROUND(u.pcPerPallet*u.unitWeight,1) as KG_PAL, ROUND(u.pcPerBox*u.unitWeight,1) as KG_BOX, u.pcPerPallet/u.pcPerBox as BOX_PAL, 'Wszystkie' as brak, cs.location as loc " _
            & "FROM tbOperations o LEFT JOIN tbOperationDataHistory od ON od.operationId=o.operationId LEFT JOIN tbZfin z ON z.zfinId=o.zfinId LEFT JOIN tbMachine m ON m.machineId=od.plMach LEFT JOIN tbZfinProperties zp ON zp.zfinId=z.zfinId LEFT JOIN tbUom u ON u.zfinId=z.zfinId LEFT JOIN tbCustomerString cs ON cs.custStringId = z.custString LEFT JOIN tbPallets p ON p.palletId=u.palletType " _
            & "WHERE od.plMoment >= '" & dFrom & "' AND od.plMoment < '" & dTo & "' AND o.type = 'p' And od.operDataVer = " & Me.cmbVersions
    ElseIf regval = "Opakowania" Then
        theMode = 5
        sql = "DECLARE @dateFrom datetime, " _
            & "@dateTo datetime " _
            & "SET @dateFrom = '" & dFrom & "' " _
            & "SET @dateTo = '" & dTo & "' " _
            & "SELECT o.mesId, theBom.bomRecId, o.zfinId, o.mesString, theBom.materialId, mat.zfinIndex, mat.zfinName, mat.zfinType, od.plMoment,od.plShift, m.machineName as Maszyna, (od.plAmount/theBom.pcPerPallet) * theBom.amount as matAmount, theBom.unit, CASE WHEN matType.materialTypeName IS NULL THEN 'Nieznany' ELSE matType.materialTypeName END as Kategoria,zp.[beans?] as bean,zp.[decafe?] as decaf, 'Wszystkie' as brak,zfin.zfinIndex as iZfin, zfin.zfinName as nZfin " _
            & "FROM tbOperations o LEFT JOIN tbOperationDataHistory od ON od.operationId=o.operationId LEFT JOIN tbMachine m ON m.machineId=od.plMach LEFT JOIN tbZfin zfin ON zfin.zfinId=o.zfinId LEFT JOIN " _
            & "(SELECT bomy.*, freshBom.dateAdded, (u.unitWeight*u.pcPerPallet) as KG_PAL, u.pcPerPallet FROM tbBom bomy RIGHT JOIN " _
            & "(SELECT oBom.zfinId,  MAX(oBom.bomRecId) as bomRecId, MAX(oBom.dateAdded) as dateAdded FROM " _
            & "(SELECT iBom.bomRecId, zfinId, br.dateAdded FROM tbBomReconciliation br JOIN ( " _
            & "SELECT bomRecId, zfinId " _
            & "FROM tbBom bom " _
            & "GROUP BY bomRecId, zfinId) iBom ON iBom.bomRecId=br.bomRecId) oBom " _
            & "WHERE oBom.dateAdded <=@dateTo " _
            & "GROUP BY oBom.zfinId) freshBom ON freshBom.zfinId=bomy.zfinId AND freshBom.bomRecId=bomy.bomRecId " _
            & "LEFT JOIN tbUom u ON u.zfinId=bomy.zfinId) theBom ON theBom.zfinId=o.zfinId LEFT JOIN tbZfin mat ON mat.zfinId=theBom.materialId LEFT JOIN tbMaterialType matType ON mat.materialType=matType.materialTypeId " _
            & "LEFT JOIN tbZfinProperties zp ON zp.zfinId=o.zfinId " _
            & "WHERE od.plMoment >= @dateFrom AND od.plMoment < @dateTo AND o.type = 'p' AND mat.zfinType = 'zpkg' And od.operDataVer = " & Me.cmbVersions
    End If
    regval = registryKeyExists(regPath & "ZfinList")
    If regval <> False And regval <> "" Then
        If Len(regval) > 0 Then
            sql = sql & " AND z.zfinIndex IN (" & regval & ")"
        End If
    End If
    
    regval = registryKeyExists(regPath & "SplitBy")
    If regval <> False And regval <> "" Then
        split = regval
        regval = registryKeyExists(regPath & "Bean")
        If regval <> False And regval <> "" Then
            If regval = "Tylko ziarno" Then
                sql = sql & " AND zp.[beans?]=1 "
            ElseIf regval = "Tylko mielona" Then
                sql = sql & " AND zp.[beans?]=0 "
            End If
        End If
        regval = registryKeyExists(regPath & "Bean")
        If regval <> False And regval <> "" Then
            If regval = "Tylko bezkofeinowa" Then
                sql = sql & " AND zp.[decafe?]=1 "
            ElseIf regval = "Tylko kofeinowa" Then
                sql = sql & " AND zp.[decafe?]=0 "
            End If
        End If
        sql = sql & " ORDER BY Maszyna, zfinIndex"
        bool = True
    End If
End If

If bool Then
    updateConnection
    Set rs = CreateObject("adodb.recordset")
    rs.Open sql, adoConn
    If Not rs.EOF Then
        clearTrace
        Set records = currentSchedule.getRecords
        If Not comparativeSchedule Is Nothing Then Set comparativeSchedule = Nothing
        Set comparativeSchedule = New clsSchedule
        With comparativeSchedule
            .initialize dFrom, dTo, theMode, CStr(split), rs, updatedOn
        End With
        rs.MoveFirst
        Do Until rs.EOF
            theSheet = Trim(rs.Fields(split))
            index = rs.Fields("zfinIndex") & "_" & theSheet
            If inCollection(index, records) Then
                theRow = records(index).location.Row
                theDate = DateSerial(year(rs.Fields("plMoment")), Month(rs.Fields("plMoment")), Day(rs.Fields("plMoment")))
                theCol = currentSchedule.getSplitters(theSheet).getShiftColumn(theDate, rs.Fields("plShift"))
                ThisWorkbook.Sheets(theSheet).Cells(theRow, theCol).Interior.ColorIndex = 44 'color = orange
                ThisWorkbook.Sheets(theSheet).Cells(theRow + 1, theCol).Interior.ColorIndex = 44 'color = orange
            End If
            rs.MoveNext
        Loop
        If TotalDaily <> False Or TotalShift <> False Then
            'check if we need to print totals
            shiftRow = 4
            If TotalDaily <> False Then
                shiftRow = 5
                If scheduleMode = 1 Then
                    If TotalDaily = "Wsad" Then
                        dailyUnit = "s"
                    Else
                        dailyUnit = "m"
                    End If
                ElseIf scheduleMode = 3 Then
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
                ElseIf scheduleMode = 3 Then
                    If TotalShift = sUnit Then
                        shiftUnit = "s"
                    ElseIf TotalShift = pUnit Then
                        shiftUnit = "m"
                    End If
                End If
            End If
            If currentSchedule.hasCharts Then
                'there are charts
                If InStr(1, currentSchedule.chartType, "narastająco", vbTextCompare) > 0 Then
                    cumul = True
                Else
                    cumul = False
                End If
                If InStr(1, currentSchedule.chartType, "Zmianowy", vbTextCompare) > 0 Then
                    daily = False
                Else
                    daily = True
                End If
            End If
            Set currentSplitters = currentSchedule.getSplitters
            For Each splitter In currentSplitters
                Erase xValues
                Erase yValues
                curDate = dFrom
                totCurrentShift = 0
                totComparativeShift = 0
                totCurrentDaily = 0
                totComparativeDaily = 0
                Do Until curDate > dTo
                    Select Case DatePart("h", curDate)
                        Case 6
                        theShift = 1
                        sShift = "I"
                        Case 14
                        theShift = 2
                        sShift = "II"
                        Case 22
                        theShift = 3
                        sShift = "III"
                    End Select
                    theDate = DateSerial(year(curDate), Month(curDate), Day(curDate))
                    If inCollection(splitter.name, comparativeSchedule.getSplitters) Then
                        compShift = comparativeSchedule.getSplitters(splitter.name).total(theDate, shiftUnit, theShift)
                    Else
                        compShift = 0
                    End If
                    totCurrentShift = totCurrentShift + splitter.total(theDate, shiftUnit, theShift, comparativeSchedule)
                    totComparativeShift = totComparativeShift + compShift
                    If theShift = 1 Then totCurrentDaily = totCurrentDaily + splitter.total(DateAdd("d", -1, theDate), shiftUnit, , comparativeSchedule)
                    If theShift = 1 Then
                        If inCollection(splitter.name, comparativeSchedule.getSplitters) Then
                            compDaily = comparativeSchedule.getSplitters(splitter.name).total(DateAdd("d", -1, theDate), shiftUnit)
                        Else
                            compDaily = 0
                        End If
                        totComparativeDaily = totComparativeDaily + compDaily
                    End If
                    theCol = splitter.getShiftColumn(theDate, theShift)
                    If daily And cumul And theShift = 1 Then appendChart WeekdayName(weekday(theDate, vbMonday)), totComparitiveDaily
                    If daily And Not cumul And theShift = 1 Then appendChart WeekdayName(weekday(theDate, vbMonday)), compDaily
                    If Not daily And cumul Then appendChart WeekdayName(weekday(theDate, vbMonday)) & " " & sShift, totComparativeShift
                    If Not daily And Not cumul Then appendChart WeekdayName(weekday(theDate, vbMonday)) & " " & sShift, compShift
                    If TotalDaily <> False And theShift = 1 Then
                        If IsNumeric(totCurrentDaily - totComparativeDaily) Then
                            If Round(totCurrentDaily - totComparativeDaily, 1) >= 0 Then
                                ThisWorkbook.Sheets(splitter.name).Cells(4, theCol - 3).Interior.ColorIndex = 43
                                If Round(totCurrentDaily - totComparativeDaily, 1) > 0 Then
                                    sign = "+"
                                Else
                                    sign = ""
                                End If
                            Else
                                ThisWorkbook.Sheets(splitter.name).Cells(4, theCol - 3).Interior.ColorIndex = 46
                                sign = ""
                            End If
                        End If
                        ThisWorkbook.Sheets(splitter.name).Cells(4, theCol - 3) = sign & CStr(Round(totCurrentDaily - totComparativeDaily, 1))
                        ThisWorkbook.Sheets(splitter.name).Cells(4, theCol - 3).NumberFormat = "@"
                    End If
                    If TotalShift <> False Then
                        If IsNumeric(totCurrentShift - totComparativeShift) Then
                            If Round(totCurrentShift - totComparativeShift, 1) >= 0 Then
                                ThisWorkbook.Sheets(splitter.name).Cells(shiftRow, theCol).Interior.ColorIndex = 43
                                If Round(totCurrentShift - totComparativeShift, 1) > 0 Then
                                    sign = "+"
                                Else
                                    sign = ""
                                End If
                            Else
                                ThisWorkbook.Sheets(splitter.name).Cells(shiftRow, theCol).Interior.ColorIndex = 46
                                sign = ""
                            End If
                        End If
                        ThisWorkbook.Sheets(splitter.name).Cells(shiftRow, theCol) = sign & CStr(Round(totCurrentShift - totComparativeShift, 1))
                        ThisWorkbook.Sheets(splitter.name).Cells(shiftRow, theCol).NumberFormat = "@"
                    End If
                    curDate = DateAdd("h", 8, curDate) 'increment
                Loop
                '---------------------------------------------------------------------------
                'we need to go one more time to have last day marked
                '---------------------------------------------------------------------------
                
                If TotalDaily <> False Then
                    curDate = DateAdd("h", -8, curDate) 'increment
                    theDate = DateSerial(year(curDate), Month(curDate), Day(curDate))
                    theCol = splitter.getShiftColumn(theDate, 1)
                    totCurrentDaily = totCurrentDaily + splitter.total(theDate, shiftUnit, , comparativeSchedule)
                    If inCollection(splitter.name, comparativeSchedule.getSplitters) Then
                        totComparativeDaily = totComparativeDaily + comparativeSchedule.getSplitters(splitter.name).total(theDate, shiftUnit)
                    End If
                    'test totComparativeShift, totCurrentShift, theDate, theShift, splitter.name
                    If IsNumeric(totCurrentDaily - totComparativeDaily) Then
                        If Round(totCurrentDaily - totComparativeDaily, 1) >= 0 Then
                            ThisWorkbook.Sheets(splitter.name).Cells(4, theCol).Interior.ColorIndex = 43
                            If Round(totCurrentDaily - totComparativeDaily, 1) > 0 Then
                                sign = "+"
                            Else
                                sign = ""
                            End If
                        Else
                            ThisWorkbook.Sheets(splitter.name).Cells(4, theCol).Interior.ColorIndex = 46
                            sign = ""
                        End If
                    End If
                    ThisWorkbook.Sheets(splitter.name).Cells(4, theCol) = sign & CStr(Round(totCurrentDaily - totComparativeDaily, 1))
                    ThisWorkbook.Sheets(splitter.name).Cells(4, theCol).NumberFormat = "@"
                End If
                If currentSchedule.hasCharts Then currentSchedule.getGraphs(splitter.name).add2ndAxis xValues, yValues, comparativeSchedule.versionString, currentSchedule.versionString
            Next splitter
            changes = missingAdded
            If Len(changes) > 0 Then
                MsgBox "Znaleziono zmiany. " & vbNewLine & changes
            End If
        End If
        Me.Hide
    Else
        MsgBox "Brak danych w wybranej wersji planu.", vbOKOnly + vbInformation, "Brak danych"
    End If
    Unload Me
Else
    MsgBox "Coś poszło nie tak. Spróbuj najpierw aktualizować harmonogram i następnie skorzystaj z rysowania śladu", vbInformation + vbOKOnly, "Coś nie pykło"
End If

exit_here:
Set rs = Nothing
Application.ScreenUpdating = True
Application.StatusBar = ""
closeConnection
Exit Sub

err_trap:
MsgBox "Error in ""getPlan"" of tracer. Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Sub

Private Sub Test(past As Double, present As Double, d As Date, shift As Integer, split As String)
Dim sht As Worksheet
Dim i As Integer

If Not isCreated Then
    Set sht = ThisWorkbook.Worksheets.add
    sht.name = "test"
    isCreated = True
Else
    Set sht = ThisWorkbook.Sheets("test")
End If

For i = 1 To 1000
    If sht.Range("A" & i) = "" Then
        sht.Range("A" & i) = CStr(d) & "_" & shift
        sht.Range("B" & i) = past
        sht.Range("C" & i) = present
        sht.Range("D" & i) = present - past
        sht.Range("E" & i) = split
        Exit For
    End If
Next i

End Sub

Private Sub appendChart(xVal As Variant, yVal As Variant)
Dim i As Integer

If isArrayEmpty(xValues) Then
    ReDim xValues(0) As Variant
    ReDim yValues(0) As Variant
    xValues(0) = xVal
    yValues(0) = yVal
Else
    i = UBound(xValues)
    ReDim Preserve xValues(i + 1) As Variant
    ReDim Preserve yValues(i + 1) As Variant
    xValues(i + 1) = xVal
    yValues(i + 1) = yVal
End If
End Sub
