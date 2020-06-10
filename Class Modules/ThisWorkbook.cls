VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Activate()
Dim regval As Variant
Dim updatedOn As Date
Dim title As String

Application.OnKey "^{TAB}", "updateSearch"

'regPath = "HKEY_CURRENT_USER\Software\Prod_sched\"
regval = registryKeyExists(regPath & "TitleString")
If regval <> False Then
    title = regval
    regval = registryKeyExists(regPath & "UpdatedOn")
    If regval <> False Then
        updatedOn = CDate(regval)
    End If
    Application.Caption = title & Abs(DateDiff("h", Now, updatedOn)) & " godz. temu"
End If
End Sub

Private Sub Workbook_Deactivate()

On Error Resume Next

With Application
    .OnKey "^{TAB}"
    .Caption = ""
    .CommandBars("Cell").Controls("Remove connection").Delete
    .CommandBars("Cell").Controls("Cancel order").Delete
End With

On Error GoTo 0

End Sub


Private Sub Workbook_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)

Dim cBut As CommandBarButton

On Error Resume Next

With Application

    .CommandBars("Cell").Controls("Szczegóły zaznaczenia").Delete

    Set cBut = .CommandBars("Cell").Controls.add(Temporary:=True)

End With



With cBut

   .Caption = "Szczegóły zaznaczenia"

   .Style = msoButtonCaption

   .OnAction = "detailForSelection"

End With


On Error GoTo 0

End Sub

Public Sub przyklad()
Dim dFrom As Date
Dim dTo As Date
Dim hFrom As Integer
Dim hTo As Integer
Dim v() As String

dFrom = ThisWorkbook.Sheets("przyklad").Range("B3")
dTo = ThisWorkbook.Sheets("przyklad").Range("F3")


hFrom = dateToMin(ThisWorkbook.Sheets("przyklad").Range("C3").Text)
hTo = dateToMin(ThisWorkbook.Sheets("przyklad").Range("G3").Text)

dFrom = DateAdd("n", hFrom, dFrom)
dTo = DateAdd("n", hTo, dTo)
MsgBox DateDiff("n", dFrom, dTo)

End Sub

Public Function dateToMin(hour As String) As Long
Dim v() As String

v = split(hour, ":", , vbTextCompare)

dateToMin = (v(0) * 60) + v(1)

End Function

