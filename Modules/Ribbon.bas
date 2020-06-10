Attribute VB_Name = "Ribbon"
Public currentWeekID As Variant
Public currentYearID As Variant
Public weekCtl As IRibbonControl
Public yearCtl As IRibbonControl
Public rib As IRibbonUI

Sub GetSelectedWeekID(control As IRibbonControl, ByRef itemID As Variant)
    Set weekCtl = control
    If isNothing(currentWeekID) Then
        If Not isNothing(ThisWorkbook.CustomDocumentProperties("week")) Then
            currentWeekID = "ddWeek" & ThisWorkbook.CustomDocumentProperties("week")
        Else
            currentWeekID = "ddWeek" & IsoWeekNumber(Date)
            updateProperty "week", IsoWeekNumber(Date)
        End If
    End If
    itemID = currentWeekID
End Sub

Sub GetSelectedYearID(control As IRibbonControl, ByRef itemID As Variant)
    Dim i As Integer
    
    If isNothing(currentYearID) Then
        If Not isNothing(ThisWorkbook.CustomDocumentProperties("year")) Then
            i = ThisWorkbook.CustomDocumentProperties("year") - 2014
        Else
             i = year(Date) - 2014
        End If
        currentYearID = "ddYear" & i
        updateProperty "year", i + 2014
    End If
    itemID = currentYearID
End Sub


Public Sub OnRibbonLoad(objRibbon As IRibbonUI)
    Set rib = objRibbon
    StoreObjRef rib, "ribbon_ref"
End Sub


Public Sub bringWeek(control As IRibbonControl)
updater.Show
End Sub

Public Sub traceChanges(control As IRibbonControl)
Dim regval As Variant
Dim bool As Boolean

bool = False

If verifyRecordset Then
    regval = registryKeyExists(regPath & "dateRangeType")
    If regval <> False And regval <> "" Then
        If regval = "Weekly" Then
            'it's set for weekly, we can continue
            bool = True
        Else
            'something went wrong, most probably it's custom-ranged
            MsgBox "Aby rysować ślad musisz najpierw aktualizować harmonogram z zaznaczoną opcją ""zakres wg tygodnia"" (zakładka ""Zakres dat"").", vbInformation + vbOKOnly, "Niedostępne dla niestandardowego zakresu"
        End If
    Else
        'something went wrong, most probably it's custom-ranged
        MsgBox "Aby rysować ślad musisz najpierw aktualizować harmonogram z zaznaczoną opcją ""zakres wg tygodnia"" (zakładka ""Zakres dat"").", vbInformation + vbOKOnly, "Niedostępne dla niestandardowego zakresu"
    End If
End If

If bool Then
    tracer.Show
End If

End Sub

Public Sub cleanTrace(control As IRibbonControl)
clearTrace

End Sub

Public Sub jump2index(control As IRibbonControl)
updateSearch
End Sub

Public Sub printSetup(control As IRibbonControl)
If verifyRecordset Then
    printer.Show
End If
End Sub




