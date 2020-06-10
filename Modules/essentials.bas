Attribute VB_Name = "essentials"
Option Explicit

' All API's from http://allapi.mentalis.org/apilist/apilist.php
Private Const POINTERSIZE As Long = 4
Private Const ZEROPOINTER As Long = 0

Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef destination As Any, _
                                                  ByRef source As Any, _
                                                  ByVal length As Long)
                                      

Public Function GetPointerToObject(ByRef objThisObject As Object) As Long
    Dim lngThisPointer As Long

    RtlMoveMemory lngThisPointer, objThisObject, POINTERSIZE
    GetPointerToObject = lngThisPointer

End Function


Public Function GetObjectFromPointer(ByVal lngThisPointer As Long) As Object
    Dim objThisObject As Object

    RtlMoveMemory objThisObject, lngThisPointer, POINTERSIZE
    Set GetObjectFromPointer = objThisObject
    RtlMoveMemory objThisObject, ZEROPOINTER, POINTERSIZE

End Function

Public Function isArrayEmpty(parArray As Variant, Optional dimension As Variant) As Boolean
'Returns true if:
'  - parArray is not an array
'  - parArray is a dynamic array that has not been initialised (ReDim)
'  - parArray is a dynamic array has been erased (Erase)

  If IsArray(parArray) = False Then isArrayEmpty = True
  On Error Resume Next
    If IsMissing(dimension) Then
        If UBound(parArray) < LBound(parArray) Then isArrayEmpty = True: Exit Function Else: isArrayEmpty = False
    Else
        If UBound(parArray, dimension) < LBound(parArray, dimension) Then isArrayEmpty = True: Exit Function Else: isArrayEmpty = False
    End If
End Function

Public Sub createLineChart(data As Variant, wsName As String, chartName As String) 'data = 2-dimensional array with X labels and values, wsName = name of sheet where chart is to be placed in
Dim rng As Range
Dim lineChart As ChartObject
Dim i As Integer
Dim N As Integer
Dim index As String
Dim x As Integer
Dim xVal() As String
Dim val() As Integer

For i = 3 To 1000 Step 2
    index = ThisWorkbook.Sheets(wsName).Cells(i, 6)
    If index = "" Then
        Exit For
    End If
Next i
                           
Set rng = ThisWorkbook.Sheets(wsName).Range("H" & i + 3 & ":y" & i + 27)

'Set lineChart = Sheets("Charts").ChartObjects.Add(Left:=rng.Left, width:=rng.width, Top:=rng.Top, Height:=rng.Height)
'
'    lineChart.Chart.ChartWizard Gallery:=xlLine, HasLegend:=True, Title:="Carrier's performance by " & sPeriod
'    With lineChart.Chart
'        For xx = 2 To 100
'            company2 = ThisWorkbook.Sheets("Charts").Cells(1, xx)
'            If company2 = "" Then
'                Exit For
'            Else
'                .SeriesCollection.NewSeries
'                With .SeriesCollection(xx - 1)
'                    .Values = Worksheets("Charts").Range(Cells(2, xx), Cells(pSpan + 2, xx))
'                    .XValues = Worksheets("Charts").Range("A2:A" & pSpan + 2)
'                    .name = company2
'                    .Smooth = True
'                    .MarkerStyle = xlMarkerStyleNone
''                    .ApplyDataLabels
'''                    .DataLabels.Select
'                End With
'            End If
'        Next xx
'    End With
'
'Set lineChart = Nothing
'Set rng = Nothing

'Set rng = ThisWorkbook.Sheets("Charts").Range("M30:y55")

ReDim xVal(UBound(data, 2)) As String
ReDim val(UBound(data, 2)) As Integer

For N = LBound(data, 2) To UBound(data, 2)
    xVal(N) = data(0, N)
    val(N) = data(1, N)
Next N

Set lineChart = Sheets(wsName).ChartObjects.add(Left:=rng.Left, width:=rng.width, Top:=rng.Top, Height:=rng.Height)

    lineChart.Chart.ChartWizard Gallery:=xlColumnClustered, HasLegend:=True, title:=chartName

'    For x = LBound(data, 2) To UBound(data, 2)
        With lineChart.Chart
            .SeriesCollection.NewSeries
            With .SeriesCollection(1)
                .values = val
                .xValues = xVal
                .name = wsName
    '                    .Smooth = True
    '                    .MarkerStyle = xlMarkerStyleNone
    '                    .ApplyDataLabels
    ''                    .DataLabels.Select
            End With
        End With
'    Next x
Set lineChart = Nothing
Set rng = Nothing

End Sub

Public Function isNothing(obj As Variant) As Boolean
Dim bool As Boolean

bool = False

On Error Resume Next

If IsObject(obj) Then
    If obj Is Nothing Then
        bool = True
    End If
Else
    If IsEmpty(obj) Then
        bool = True
    Else
        If IsNull(obj) Then
            bool = True
        End If
    End If
End If

isNothing = bool

End Function

Public Function inCollection(ind As String, col As Collection) As Boolean
Dim v As Variant
Dim isError As Boolean

isError = False

On Error GoTo err_trap

Set v = col(ind)

Exit_here:
If isError Then
    inCollection = False
Else
    inCollection = True
End If
Exit Function

err_trap:
isError = True
Resume Exit_here


End Function

Function IsUserFormLoaded(ByVal UFName As String) As Boolean
    Dim UForm As Object
     
    IsUserFormLoaded = False
    For Each UForm In VBA.UserForms
        If UForm.name = UFName Then
            IsUserFormLoaded = True
            Exit For
        End If
    Next
End Function 'IsUserFormLoaded

Public Function cell2letter(c As Integer) As String
Dim arr() As String
With ThisWorkbook.Sheets(1)
    arr = split(.Cells(1, c).Address(True, False), "$")
    cell2letter = arr(0)
End With
End Function

Public Function sheetExists(sheetToFind As String) As Boolean
Dim sheet As Worksheet
    sheetExists = False
    For Each sheet In Worksheets
        If sheetToFind = sheet.name Then
            sheetExists = True
            Exit Function
        End If
    Next sheet
End Function
