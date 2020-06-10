Attribute VB_Name = "graphs"
Public Sub deleteAllGraphs(shtName As String)
Dim sht As Worksheet
Dim cht As ChartObject
Set sht = ThisWorkbook.Sheets(shtName)

For Each cht In sht.ChartObjects
    cht.Delete
Next cht

End Sub


Public Sub createLineChart(data As Variant, chartName As String, putInRng As Range) 'data = 2-dimensional array with X labels and values, wsName = name of sheet where chart is to be placed in
Dim lineChart As ChartObject
Dim i As Integer
Dim N As Integer
Dim index As String
Dim x As Integer
Dim xVal() As String
Dim val() As Integer
Dim sht As Worksheet

Set sht = putInRng.Worksheet

ReDim xVal(UBound(data, 2)) As String
ReDim val(UBound(data, 2)) As Integer

For N = LBound(data, 2) To UBound(data, 2)
    xVal(N) = data(0, N)
    val(N) = data(1, N)
Next N

Set lineChart = sht.ChartObjects.add(Left:=putInRng.Left, width:=putInRng.width, Top:=putInRng.Top, Height:=putInRng.Height)

    lineChart.Chart.ChartWizard Gallery:=xlColumnClustered, HasLegend:=True, title:=chartName

'    For x = LBound(data, 2) To UBound(data, 2)
        With lineChart.Chart
            .SeriesCollection.NewSeries
            With .SeriesCollection(1)
                .values = val
                .xValues = xVal
                .name = chartName
    '                    .Smooth = True
    '                    .MarkerStyle = xlMarkerStyleNone
    '                    .ApplyDataLabels
    ''                    .DataLabels.Select
            End With
        End With
'    Next x
Set lineChart = Nothing
End Sub

