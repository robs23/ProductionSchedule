VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pGraphTitle As String
Private pGraphName As String
Private pValues() As Variant 'multidimensional array of values that should be displayed on X and on Y axis
Private pxValues() As Variant 'array of x values
Private pyValues() As Variant 'array of y values
Private pxTitle As String 'title of X axis
Private pyTitle As String 'title of Y axis
Private pSheetName As String
Private pLocation As Range 'location of where the graph is placed
Private chrt As ChartObject ' the graph itself
Private pGraphType As XlChartType
Private pFirstAxis As Series 'first axis
Private pSecondAxis As Series 'second axis
Private pHas2ndAxis As Boolean 'does it have second axis?

Public Property Get sheetName() As String
sheetName = pSheetName
End Property

Public Property Get xTitle() As String
xTitle = pxTitle
End Property

Public Property Get yTitle() As String
yTitle = pyTitle
End Property

Public Property Get graphTitle() As String
graphTitle = pGraphTitle
End Property

Public Property Get location() As Range
Set location = pLocation
End Property

Public Property Get has2ndAxis() As Boolean
has2ndAxis = pHas2ndAxis
End Property

Public Sub initialize(graphTitle As String, graphName As String, graphType As XlChartType, xTitle As String, yTitle As String, sheetName As String, location As Range)
pGraphTitle = graphTitle
pGraphName = graphName
pGraphType = graphType
pxTitle = xTitle
pyTitle = yTitle
pSheetName = sheetName
Set pLocation = location
End Sub

Public Sub append(xValue As Variant, yValue As Variant)

On Error GoTo err_trap

If isArrayEmpty(pValues, 2) Then
    ReDim pValues(1, 0) As Variant
    pValues(0, 0) = xValue
    pValues(1, 0) = yValue
    ReDim pxValues(0) As Variant
    pxValues(0) = xValue
    ReDim pyValues(0) As Variant
    pyValues(0) = yValue
Else
    ReDim Preserve pValues(1, UBound(pValues, 2) + 1) As Variant
    pValues(0, UBound(pValues, 2)) = xValue
    pValues(1, UBound(pValues, 2)) = yValue
    ReDim Preserve pxValues(UBound(pxValues) + 1) As Variant
    pxValues(UBound(pxValues)) = xValue
    ReDim Preserve pyValues(UBound(pyValues) + 1) As Variant
    pyValues(UBound(pyValues)) = yValue
End If

exit_here:
Exit Sub

err_trap:
MsgBox "Error in ""append"" of clsGraph. Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Sub

Public Sub createChart()
Dim i As Integer
Dim N As Integer
Dim index As String
Dim x As Integer
Dim xVal() As String
Dim val() As Integer
Dim sht As Worksheet

Set sht = pLocation.Worksheet

Set chrt = sht.ChartObjects.add(Left:=pLocation.Left, width:=pLocation.width, Top:=pLocation.Top, Height:=pLocation.Height)

chrt.Chart.ChartWizard Gallery:=pGraphType, HasLegend:=True, title:=pGraphTitle
With chrt.Chart
    Set pFirstAxis = .SeriesCollection.NewSeries
'    With .SeriesCollection(1)
    With pFirstAxis
        .values = pyValues
        .xValues = pxValues
        .name = pGraphName
        .Format.line.ForeColor.RGB = RGB(30, 144, 255)
        If pGraphType = xlLine Then
            .MarkerForegroundColor = RGB(30, 144, 255)
            .MarkerBackgroundColor = RGB(30, 144, 255)
        End If
'                        .Smooth = True
'                        .MarkerStyle = xlMarkerStyleNone
'                        .ApplyDataLabels
'                    .DataLabels.Select
    End With
End With
chrt.Chart.HasLegend = False
End Sub

Public Sub add2ndAxis(xValues As Variant, yValues As Variant, newAxisName As String, renameExistentAxisTo As String)

Set pSecondAxis = chrt.Chart.SeriesCollection.NewSeries

With pSecondAxis
    .values = yValues
    .xValues = xValues
    .name = newAxisName
    .Format.line.ForeColor.RGB = RGB(220, 20, 60)
    If pGraphType = xlLine Then
        .MarkerForegroundColor = RGB(220, 20, 60)
        .MarkerBackgroundColor = RGB(220, 20, 60)
    End If
End With

pFirstAxis.name = renameExistentAxisTo
pHas2ndAxis = True
chrt.Chart.HasLegend = True
chrt.Chart.Legend.Position = xlLegendPositionBottom
End Sub

Public Sub remove2ndAxis()
pSecondAxis.Delete
Set pSecondAxis = Nothing
pHas2ndAxis = False
chrt.Chart.HasLegend = False
End Sub
