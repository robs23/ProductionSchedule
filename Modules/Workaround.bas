Attribute VB_Name = "Workaround"
Public Sub UpdateProductionPlan()
updater.Show
End Sub

Public Sub ShowTracer()
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
