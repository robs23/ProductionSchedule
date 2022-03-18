VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} printer 
   Caption         =   "Ustawienia wydruku"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6120
   OleObjectBlob   =   "printer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "printer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub populator()
Dim i As Integer
Dim sht As Worksheet
Dim pnt As clsPrinter

For i = Me.cmbContent.ListCount To 1 Step -1
    Me.cmbContent.RemoveItem i
Next i

For i = Me.cmbRange.ListCount To 1 Step -1
    Me.cmbRange.RemoveItem i
Next i

For i = Me.cmbPrinter.ListCount To 1 Step -1
    Me.cmbPrinter.RemoveItem i
Next i

Me.cmbContent.AddItem "Dane"
Me.cmbContent.AddItem "Wykresy"
Me.cmbContent.AddItem "Dane i wykresy"

Me.cmbRange.AddItem "Wszystkie arkusze"

For Each sht In ThisWorkbook.Sheets
    Me.cmbRange.AddItem sht.name
Next sht

getPrinters
For Each pnt In printersCollection
    Me.cmbPrinter.AddItem pnt.PrinterName
Next pnt

End Sub

Private Sub btnPrint_Click()
Dim theMode As Integer

If validate Then
    updateRegistry regPath & "PrintContent", Me.cmbContent
    updateRegistry regPath & "PrintRange", Me.cmbRange
    updateRegistry regPath & "PrintChosenPrinter", Me.cmbPrinter
    If Me.cmbContent = "Dane" Then
        theMode = 1
    ElseIf Me.cmbContent = "Wykresy" Then
        theMode = 2
    ElseIf Me.cmbContent = "Dane i wykresy" Then
        theMode = 3
    End If
    If Me.cmbRange <> "Wszystkie arkusze" Then
        printData theMode, Me.cmbRange
    Else
        printData theMode
    End If
End If
End Sub

Private Function validate() As Boolean

If Me.cmbContent = "" Or Me.cmbRange = "" Or Me.cmbPrinter = "" Then
    MsgBox "Aby kontynuować wszystkie ustawienia muszą być uzupełnione!", vbOKOnly + vbExclamation, "Wybierz ustawienia wydruku"
Else
    If InStr(1, Me.cmbContent, "wykresy", vbTextCompare) <> 0 Then
        If currentSchedule.hasCharts = True Then
            validate = True
        Else
            MsgBox "Brak wykresów w aktualnym harmonogramie! W polu ""Zawartość wydruku"" wybierz ""Dane"" lub aktualizuj harmonogram włączając wykresy (zakładka ""Wygląd"") przed wydrukiem ", vbOKOnly + vbExclamation, "Brak wykresów"
            validate = False
        End If
    Else
        validate = True
    End If
End If
End Function



Private Sub UserForm_Initialize()
populator
bringValues
End Sub


Private Sub bringValues()
Dim regval As Variant

regval = registryKeyExists(regPath & "PrintContent")
If regval <> False And isListed("cmbContent", regval) Then Me.cmbContent = regval

regval = registryKeyExists(regPath & "PrintRange")
If regval <> False And isListed("cmbRange", regval) Then Me.cmbRange = regval

regval = registryKeyExists(regPath & "PrintChosenPrinter")
If regval <> False And isListed("cmbPrinter", regval) Then
    Me.cmbPrinter = regval
End If


End Sub

Private Function isListed(cmbName As String, val As Variant) As Boolean
Dim cmb As ComboBox
Dim i As Integer
Dim bool As Boolean

bool = False

val = CStr(val)

Set cmb = Me.Controls(cmbName)

For i = 0 To cmb.ListCount - 1
    If cmb.List(i) = val Then
        bool = True
        Exit For
    End If
Next i

isListed = bool

End Function

Private Sub printData(theMode As Integer, Optional shtName As Variant)
'theMode = 1-dane, 2-wykresy, 3-dane i wyrkesy
Dim sht As Worksheet
Dim newSht As Worksheet
Dim tot As Integer

On Error GoTo err_trap

Application.StatusBar = "Trwa drukowanie wybranych arkuszy.. Proszę czekać.."
Application.ScreenUpdating = False
Application.Cursor = xlWait

setPrinter

If Not IsMissing(shtName) Then
    If theMode = 2 Then
        With ThisWorkbook.Sheets(CStr(shtName))
            .ChartObjects(1).Select
            .ChartObjects(1).Activate
        End With
        ActiveChart.printOut
    Else
        'add new sheet
        Set newSht = ThisWorkbook.Sheets.add
        newSht.name = "HiddenXXX"
        Set sht = ThisWorkbook.Sheets(CStr(shtName))
        printSheet sht, theMode
    End If
Else
    If theMode = 2 Then
        tot = ThisWorkbook.Sheets.Count
        i = 0
        For Each sht In ThisWorkbook.Sheets
            i = i + 1
            Application.StatusBar = "Trwa drukowanie wybranych arkuszy.. Ukończono " & Round(i / tot * 100, 0) & "% Proszę czekać.."
            With sht
                .ChartObjects(1).Select
                .ChartObjects(1).Activate
            End With
            ActiveChart.printOut
        Next sht
    Else
        tot = ThisWorkbook.Sheets.Count
        i = 0
        For Each sht In ThisWorkbook.Sheets
            If sht.name <> "HiddenXXX" Then
                'add new sheet
                Set newSht = ThisWorkbook.Sheets.add
                newSht.name = "HiddenXXX"
                i = i + 1
                Application.StatusBar = "Trwa drukowanie wybranych arkuszy.. Ukończono " & Round(i / tot * 100, 0) & "% Proszę czekać.."
                printSheet sht, theMode
                Application.DisplayAlerts = False
                ThisWorkbook.Sheets("HiddenXXX").Delete
                Application.DisplayAlerts = True
            End If
        Next sht
    End If
End If

exit_here:
If sheetExists("HiddenXXX") Then
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("HiddenXXX").Delete
    Application.DisplayAlerts = True
End If
Application.ScreenUpdating = True
Application.StatusBar = ""
Application.Cursor = xlDefault
Exit Sub

err_trap:
MsgBox "Error in ""printData"" of printer. Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Sub


Private Sub setPrinter()
Dim regval As Variant
Dim pnt As clsPrinter

Application.PrintCommunication = False
'setup printer

regval = registryKeyExists(regPath & "PrintChosenPrinter")
If regval <> False Then
    regval = CStr(regval)
    If Not getCollectionMember(regval, printersCollection) Is Nothing Then
        Set pnt = getCollectionMember(regval, printersCollection)
        Application.ActivePrinter = pnt.PriterString
    End If
End If

Application.PrintCommunication = True
End Sub

Private Sub printSheet(sht As Worksheet, theMode As Integer)
Dim newSht As Worksheet
Dim rng As Range
Dim titleRng As Range
Dim shtName As String
Dim destRng As Range
Dim x As Integer 'number of rows to be added to range if charts are to be printed

shtName = sht.name

Set newSht = ThisWorkbook.Sheets("HiddenXXX")
Set rng = currentSchedule.getSplitters(shtName).getTotalRange
rng.Copy
newSht.Range("A2").PasteSpecial xlPasteColumnWidths
newSht.Range("A2").PasteSpecial xlPasteValues, , False, False
newSht.Range("A2").PasteSpecial xlPasteFormats, , False, False
newSht.Select
ActiveWindow.Zoom = 90

If theMode = 3 Then
    sht.ChartObjects(1).Copy
    'newSht.Range("A5").PasteSpecial xlPasteAll
    newSht.Range(newSht.Cells(currentSchedule.getSplitters(shtName).getLastRow + 2, currentSchedule.getSplitters(shtName).firstDataCol), newSht.Cells(currentSchedule.getSplitters(shtName).getLastRow + 2, currentSchedule.getSplitters(shtName).firstDataCol)).PasteSpecial xlPasteAll
    x = 25
End If

Set destRng = newSht.Range(newSht.Cells(1, 1), newSht.Cells(currentSchedule.getSplitters(shtName).getLastRow + 1 + x, currentSchedule.getSplitters(shtName).getLastCol))

Set titleRng = newSht.Range(newSht.Cells(1, 1), newSht.Cells(1, currentSchedule.getSplitters(shtName).getLastCol))
With titleRng
    .Merge
    .value = shtName
    .Font.Bold = True
    .Font.Size = 20
    .RowHeight = 30
End With

Application.PrintCommunication = False
With newSht.PageSetup
    .Orientation = xlLandscape
    .FitToPagesWide = True
    .FitToPagesTall = False
    .BottomMargin = Application.CentimetersToPoints(0.5)
    .LeftMargin = Application.CentimetersToPoints(0.5)
    .RightMargin = Application.CentimetersToPoints(0.5)
    .TopMargin = Application.CentimetersToPoints(0.5)
End With
Application.PrintCommunication = True


destRng.printOut

End Sub

