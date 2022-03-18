VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} updater 
   Caption         =   "Opcje aktualizacji"
   ClientHeight    =   6072
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6870
   OleObjectBlob   =   "updater.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "updater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oldWeek As Integer
Private oldYear As Integer
Private updatedOn As Date

Private Sub btnHide_Click()
Me.Hide
End Sub

Private Sub constructTitle()
Dim titleStr As String

titleStr = "Plan "
If Me.cmbType = "Prażenie" Then
    titleStr = titleStr & "prażenia "
ElseIf Me.cmbType = "Pakowanie" Then
    titleStr = titleStr & "pakowania "
ElseIf Me.cmbType = "Palety" Then
    titleStr = titleStr & "wykorzystania palet "
End If

If Me.optWeek = True Then
    titleStr = titleStr & "za tydzień " & Me.cmbWeek & "|" & Me.cmbYear & " aktualizowany "
Else
    titleStr = titleStr & "za okres " & Me.txtFrom.value & "-" & Me.txtTo.value & " aktualizowany "
End If

updateRegistry regPath & "TitleString", titleStr
Application.Caption = titleStr & Abs(DateDiff("h", Now, updatedOn)) & " godz. temu"
End Sub

Private Sub btnListFromClipboard_Click()
Dim DataObj As MsForms.DataObject
Dim clip As String
Dim clips() As String
Dim ind As String
Dim i As Integer

Set DataObj = New MsForms.DataObject '<~~ Amended as per jp's suggestion

On Error GoTo err_trap

'~~> Get data from the clipboard.
DataObj.GetFromClipboard
'~~> Get clipboard contents
clip = DataObj.GetText(1)
clips = split(clip, vbNewLine, , vbTextCompare)

For i = 0 To UBound(clips) - 1
    If Len(clips(i)) > 0 Then
        ind = ind & clips(i) & ","
    End If
Next i

If Len(ind) > 0 Then ind = Left(ind, Len(ind) - 1)

Me.txtZfinList = ind

exit_here:
Exit Sub
   
err_trap:
If Err <> 0 Then MsgBox "Data on clipboard is not text or is empty"
Resume exit_here

End Sub

Private Sub btnRemoveList_Click()
Me.txtZfinList.value = ""
End Sub

Private Sub btnUpdate_Click()
Dim rs As ADODB.Recordset
Dim rsSplit As ADODB.Recordset
Dim nSht As Worksheet
Dim sql As String
Dim dFrom As Date
Dim dTo As Date
Dim rng As Range
Dim cSplitter As clsSchduleSplitter
Dim allSplitters As Collection
Dim tempSheet As Worksheet
Dim beanStr As String
Dim decafStr As String
Dim v() As String
Dim dday As Integer
Dim hhour As Integer
Dim splitBy As String
Dim tblName As String
Dim verStr As String
Dim theMode As Integer '1-prazenie, 2-mielenie, 3-pakowanie, 4-palety,5-opakowania, 6-lista

On Error GoTo err_trap

If verify Then
    Application.ScreenUpdating = False
    Application.StatusBar = "Pracuję.. proszę czekać.."
    Application.Cursor = xlWait
    updateRegistry regPath & "Type", Me.cmbType
    updateRegistry regPath & "SplitBy", Me.cmbSplit
    updateRegistry regPath & "Unit", Me.cmbUnit
    updateRegistry regPath & "Charts", Me.cmbChart
    updateRegistry regPath & "SecondaryUnit", Me.cmbSecUnit
    updateRegistry regPath & "ProductProperties", Me.cmbProductProperties
    updateRegistry regPath & "ProductName", Me.cmbProductName
    updateRegistry regPath & "TotalShift", Me.cmbTotalShift
    updateRegistry regPath & "TotalDaily", Me.cmbTotalDaily
    updateRegistry regPath & "TotalIndex", Me.cmbTotalIndex
    updateRegistry regPath & "Customer", Me.cmbCustomer
    updateRegistry regPath & "Machine", Me.cmbMachine
    updateRegistry regPath & "RangeStart", Me.cmbFirst
    updateRegistry regPath & "UnitRatio", Me.cmbRatio
    updateRegistry regPath & "Bean", Me.cmbBean
    updateRegistry regPath & "Decafe", Me.cmbDecaf
    updateRegistry regPath & "ShowComments", Me.cmbComments
    updateRegistry regPath & "ZfinList", Me.txtZfinList.value
    If Me.optWeek Then
        updateRegistry regPath & "dateRangeType", "Weekly"
    Else
        updateRegistry regPath & "dateRangeType", "Custom"
    End If
    If Me.cmbWeek <> "" Then updateRegistry regPath & "LoadedWeek", CInt(Me.cmbWeek)
    If Me.cmbYear <> "" Then updateRegistry regPath & "LoadedYear", CInt(Me.cmbYear)
    If Me.optCustom Then
        updateRegistry regPath & "CustomRangeStartDate", CStr(Me.txtFrom.value)
        updateRegistry regPath & "CustomRangeEndDate", CStr(Me.txtTo.value)
        updateRegistry regPath & "CustomRangeStartTime", Me.cmbTFrom
        updateRegistry regPath & "CustomRangeEndTime", Me.cmbTTo
    End If
    If Me.cmbType = "Prażenie" Then getRoastingBatches
    pUnit = Me.cmbUnit
    sUnit = Me.cmbSecUnit
    If Me.optCustom = True Then
        dFrom = DateAdd("h", CDbl(Left(Me.cmbTFrom, 2)), Me.txtFrom)
        dTo = DateAdd("h", CDbl(Left(Me.cmbTTo, 2)), Me.txtTo)
    Else
        v = split(Me.cmbFirst, " ", , vbTextCompare)
        dday = WeekDayName2Int(v(0))
        hhour = CInt(Left(v(1), 2))
        dFrom = DateAdd("h", hhour, Week2Date(CLng(Me.cmbWeek), CLng(Me.cmbYear), dday, vbFirstFourDays))
        dTo = DateAdd("h", 167, dFrom)
    End If
    If Me.cmbVersion = 0 Then
        tblName = "tbOperationData"
        verStr = ""
        updatedOn = getLatestVersion()
        updateRegistry regPath & "UpdatedOn", CStr(updatedOn)
    ElseIf Me.cmbVersion > 0 Then
        tblName = "tbOperationDataHistory"
        verStr = "AND od.operDataVer = " & Me.cmbVersion
        updatedOn = getLatestVersion(Me.cmbVersion)
        updateRegistry regPath & "UpdatedOn", CStr(updatedOn)
    End If
    If Me.cmbType = "Prażenie" Then
        sql = "SELECT o.mesId, o.mesString, z.zfinIndex,z.zfinName, od.plMoment,od.plShift, m.machineName as Maszyna, od.plAmount as KG, o.type, zp.[beans?] as bean, zp.[decafe?] as decaf, 'Wszystkie' as brak " _
            & "FROM tbOperations o LEFT JOIN " & tblName & " od ON od.operationId=o.operationId LEFT JOIN tbZfin z ON z.zfinId=o.zfinId LEFT JOIN tbMachine m ON m.machineId=od.plMach LEFT JOIN tbZfinProperties zp ON zp.zfinId=z.zfinId " _
            & "WHERE od.plMoment >= '" & dFrom & "' AND od.plMoment < '" & dTo & "' AND o.type = 'r' " & verStr
    ElseIf Me.cmbType = "Pakowanie" Then
        sql = "SELECT o.mesId,o.mesString, z.zfinIndex,z.zfinName, od.plMoment,od.plShift, m.machineName as Maszyna, cs.custString as Klient, od.plAmount as PC, ROUND(od.plAmount*u.unitWeight,1) as KG, ROUND(od.plAmount/u.pcPerBox,1) AS BOX, ROUND(od.plAmount/u.pcPerPallet,1) AS PAL ,o.type, zp.[beans?] as bean, zp.[decafe?] as decaf,CASE WHEN cs.custString IS NOT NULL THEN LEFT(cs.custString,2) ELSE 'UNKNOWN' END as Kierunek,CASE WHEN p.palletLength+p.palletWidth = 2000 THEN 'EUR' ELSE CASE WHEN p.palletLength+p.palletWidth = 2200 THEN 'IND' ELSE 'UNKNOWN' END END AS palType, p.palletChep AS Chep, " _
            & "u.pcPerPallet as PC_PAL, u.pcPerBox as PC_BOX, ROUND(1/u.unitWeight,1) as PC_KG, ROUND(u.pcPerPallet*u.unitWeight,1) as KG_PAL, ROUND(u.pcPerBox*u.unitWeight,1) as KG_BOX, u.pcPerPallet/u.pcPerBox as BOX_PAL, 'Wszystkie' as brak, cs.location as loc " _
            & "FROM tbOperations o LEFT JOIN " & tblName & " od ON od.operationId=o.operationId LEFT JOIN tbZfin z ON z.zfinId=o.zfinId LEFT JOIN tbMachine m ON m.machineId=od.plMach LEFT JOIN tbZfinProperties zp ON zp.zfinId=z.zfinId LEFT JOIN tbUom u ON u.zfinId=z.zfinId LEFT JOIN tbCustomerString cs ON cs.custStringId = z.custString LEFT JOIN tbPallets p ON p.palletId=u.palletType " _
            & "WHERE od.plMoment >= '" & dFrom & "' AND od.plMoment < '" & dTo & "' AND o.type = 'p' " & verStr
    ElseIf Me.cmbType = "Palety" Then
        sql = "SELECT o.mesId,o.mesString, p.palletId as zfinIndex,p.palletName as zfinName,od.plMoment,od.plShift,m.machineName as Maszyna, cs.custString as Klient, 'Palety' as Palety, od.plAmount as PC,ROUND(od.plAmount*u.unitWeight,1) as KG, ROUND(od.plAmount/u.pcPerBox,1) AS BOX, ROUND(od.plAmount/u.pcPerPallet,1) AS PAL ,o.type, zp.[beans?] as bean,zp.[decafe?] as decaf, " _
            & "CASE WHEN cs.custString IS NOT NULL THEN LEFT(cs.custString,2) ELSE 'UNKNOWN' END as Kierunek,CASE WHEN p.palletLength+p.palletWidth = 2000 THEN 'EUR' ELSE CASE WHEN p.palletLength+p.palletWidth = 2200 THEN 'IND' ELSE 'UNKNOWN' END END AS palType, p.palletChep AS Chep, u.pcPerPallet as PC_PAL, u.pcPerBox as PC_BOX, ROUND(1/u.unitWeight,1) as PC_KG, ROUND(u.pcPerPallet*u.unitWeight,1) as KG_PAL, ROUND(u.pcPerBox*u.unitWeight,1) as KG_BOX, u.pcPerPallet/u.pcPerBox as BOX_PAL, 'Wszystkie' as brak, cs.location as loc " _
            & "FROM tbOperations o LEFT JOIN " & tblName & " od ON od.operationId=o.operationId LEFT JOIN tbZfin z ON z.zfinId=o.zfinId LEFT JOIN tbMachine m ON m.machineId=od.plMach LEFT JOIN tbZfinProperties zp ON zp.zfinId=z.zfinId LEFT JOIN tbUom u ON u.zfinId=z.zfinId LEFT JOIN tbCustomerString cs ON cs.custStringId = z.custString LEFT JOIN tbPallets p ON p.palletId=u.palletType " _
            & "WHERE od.plMoment >= '" & dFrom & "' AND od.plMoment < '" & dTo & "' AND o.type = 'p' " & verStr
    ElseIf Me.cmbType = "Opakowania" Then
        sql = "SELECT o.mesId, theBom.bomRecId, o.zfinId, o.mesString, theBom.materialId, mat.zfinIndex, mat.zfinName, mat.zfinType, od.plMoment,od.plShift, m.machineName as Maszyna, (od.plAmount/theBom.pcPerPallet) * theBom.amount as matAmount, theBom.unit, CASE WHEN matType.materialTypeName IS NULL THEN 'Nieznany' ELSE matType.materialTypeName END as Kategoria,zp.[beans?] as bean,zp.[decafe?] as decaf, 'Wszystkie' as brak,z.zfinIndex as iZfin, z.zfinName as nZfin " _
            & "FROM tbOperations o LEFT JOIN " & tblName & " od ON od.operationId=o.operationId LEFT JOIN tbMachine m ON m.machineId=od.plMach LEFT JOIN tbZfin z ON z.zfinId=o.zfinId LEFT JOIN " _
            & "(SELECT bomy.*, freshBom.dateAdded, (u.unitWeight*u.pcPerPallet) as KG_PAL, u.pcPerPallet FROM tbBom bomy RIGHT JOIN " _
            & "(SELECT oBom.zfinId,  MAX(oBom.bomRecId) as bomRecId, MAX(oBom.dateAdded) as dateAdded FROM " _
            & "(SELECT iBom.bomRecId, zfinId, br.dateAdded FROM tbBomReconciliation br JOIN ( " _
            & "SELECT bomRecId, zfinId " _
            & "FROM tbBom bom " _
            & "GROUP BY bomRecId, zfinId) iBom ON iBom.bomRecId=br.bomRecId) oBom " _
            & "WHERE oBom.dateAdded <= '" & dTo & "' " _
            & "GROUP BY oBom.zfinId) freshBom ON freshBom.zfinId=bomy.zfinId AND freshBom.bomRecId=bomy.bomRecId " _
            & "LEFT JOIN tbUom u ON u.zfinId=bomy.zfinId) theBom ON theBom.zfinId=o.zfinId LEFT JOIN tbZfin mat ON mat.zfinId=theBom.materialId LEFT JOIN tbMaterialType matType ON mat.materialType=matType.materialTypeId " _
            & "LEFT JOIN tbZfinProperties zp ON zp.zfinId=o.zfinId " _
            & "WHERE od.plMoment >= '" & dFrom & "' AND od.plMoment < '" & dTo & "' AND o.type = 'p' AND mat.zfinType = 'zpkg' " & verStr
    End If
    If Len(Me.txtZfinList.value) > 0 Then
        sql = sql & " AND z.zfinIndex IN (" & Me.txtZfinList.value & ")"
    End If
    If Me.cmbBean = "Tylko ziarno" Then
        sql = sql & " AND zp.[beans?]=1 "
    ElseIf Me.cmbBean = "Tylko mielona" Then
        sql = sql & " AND zp.[beans?]=0 "
    End If
    If Me.cmbDecaf = "Tylko bezkofeinowa" Then
        sql = sql & " AND zp.[decafe?]=1 "
    ElseIf Me.cmbDecaf = "Tylko kofeinowa" Then
        sql = sql & " AND zp.[decafe?]=0 "
    End If
    If Me.cmbType = "Prażenie" Then
        If Me.cmbSplit = "Maszyna" Then
            splitBy = "Maszyna"
        Else
            splitBy = "Brak"
        End If
        theMode = 1
    ElseIf Me.cmbType = "Palety" Then
        splitBy = "Brak"
        theMode = 4
    ElseIf Me.cmbType = "Opakowania" Then
        theMode = 5
        If Me.cmbSplit = "Kategoria" Then
            splitBy = "Kategoria"
        ElseIf Me.cmbSplit = "Maszyna" Then
            splitBy = "Maszyna"
        Else
            splitBy = "Brak"
        End If
    Else
        theMode = 3
        If Me.cmbSplit = "Maszyna" Then
            splitBy = "Maszyna"
        ElseIf Me.cmbSplit = "Kierunek" Then
            splitBy = "Kierunek"
        ElseIf Me.cmbSplit = "Klient" Then
            splitBy = "Klient"
        ElseIf Me.cmbSplit = "Brak" Then
            splitBy = "Brak"
        End If
    End If
    sql = sql & "ORDER BY " & splitBy & " DESC, plMoment ASC"
    updateConnection
    Set rs = CreateObject("adodb.recordset")
    rs.Open sql, adoConn
    If rs.EOF Then
        MsgBox "Brak wyników dla wybranego okresu!", vbExclamation + vbOKOnly, "Brak wyników"
    Else
        constructTitle
        If IsUserFormLoaded("finder") Then
            Unload finder
        End If
        Set currentSchedule = New clsSchedule
        With currentSchedule
            .initialize dFrom, dTo, theMode, splitBy, rs, updatedOn
            Set allSplitters = .getSplitters
        End With
'        sql = "SELECT DISTINCT m.machineName " _
'            & "FROM tbOperations o LEFT JOIN tbOperationData od ON od.operationId=o.operationId LEFT JOIN tbMachine m ON m.machineId=od.plMach " _
'            & "WHERE od.plMoment BETWEEN '" & dFrom & "' AND '" & dTo & "' AND o.type = 'r' " _
'            & "ORDER BY m.machineName DESC;"
'        Set rsSplit = CreateObject("adodb.recordset")
'        rsSplit.Open sql, adoConn
'        rsSplit.MoveFirst
        Set tempSheet = ThisWorkbook.Sheets.add
        clear tempSheet.name 'remove all sheets, we'll be adding new ones
        For Each cSplitter In allSplitters
            Set nSht = ThisWorkbook.Worksheets.add
            nSht.name = cSplitter.name
            createHeader dFrom, dTo, cSplitter.name
            cSplitter.deployResults
        Next cSplitter
        If Me.cmbChart <> "" And Me.cmbChart <> "Brak" Then
            currentSchedule.createGraphs
        End If
        Application.DisplayAlerts = False
        tempSheet.Delete
        Application.DisplayAlerts = True
        Me.Hide
    End If
    rs.Close
End If

exit_here:
Set rs = Nothing
Set rsSplit = Nothing
closeConnection
Application.ScreenUpdating = True
Application.StatusBar = ""
Application.Cursor = xlDefault
Exit Sub

err_trap:
MsgBox "Error in ""btnUupdate_Click"" of ""Updater"". Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Sub

Private Sub clear(shtName As String)
Dim sht As Worksheet

Application.DisplayAlerts = False

For Each sht In ThisWorkbook.Sheets
    If sht.name <> shtName Then
        sht.Delete
    End If
Next sht

Application.DisplayAlerts = True
End Sub

Private Function verify() As Boolean
Dim bool As Boolean

On Error GoTo err_trap

bool = True

If Me.cmbChart = "" Then Me.cmbChart = "Brak"
If Me.cmbVersion = "" Then Me.cmbVersion = "Najnowsza"
If Me.cmbFirst = "" Then Me.cmbFirst = "Niedziela 22:00"
 
If Me.optCustom = True Then
    Me.MultiPage1.value = 0
    If Me.txtFrom < #1/1/2010# Or Me.txtTo < #1/1/2010# Then
        MsgBox "Podaj datę ""Od"" i ""Do"" zestawienia", vbExclamation + vbOKOly, "Brak daty"
        bool = False
    Else
        If Me.txtFrom > Me.txtTo Then
            MsgBox "Data ""Do"" powinna być późniejsza niż data ""Od""", vbExclamation + vbOKOly, "Nieprawidłowy zakres"
            bool = False
        End If
    End If
Else
    If Me.cmbWeek = "" Or Me.cmbYear = "" Then
        MsgBox "Wybierz tydzień i rok zestawienia", vbExclamation + vbOKOly, "Nie wybrano tygodnia/roku"
        bool = False
    End If
End If

If bool Then
    If Me.cmbType = "" Then
        MsgBox "Wybierz ""Typ"" z listy rozwijanej (zakładka ""Opcje"")", vbExclamation + vbOKOly, "Wybierz typ"
        bool = False
    Else
        If Me.cmbUnit = "" Then
            MsgBox "Wybierz ""Jednostka podst."" z listy rozwijanej (zakładka ""Opcje"")", vbExclamation + vbOKOly, "Wybierz jednostke"
            bool = False
        Else
            If Me.cmbSplit = "" Then
                MsgBox "Wybierz ""Kryterium podziału"" z listy rozwijanej (zakładka ""Opcje"")", vbExclamation + vbOKOly, "Wybierz kryterium podziału"
                bool = False
            End If
        End If
    End If
End If

exit_here:
verify = bool
Exit Function

err_trap:
MsgBox "Error in ""Verify"" of ""Updater"". Error number: " & Err.Number & ", " & Err.Description
bool = False
Resume exit_here

End Function

Private Sub cmbType_AfterUpdate()
Dim i As Integer

If Me.cmbType = "Prażenie" Then
    For i = Me.cmbSecUnit.ListCount - 1 To 0 Step -1
        Me.cmbSecUnit.RemoveItem i
    Next i
    
    For i = Me.cmbUnit.ListCount - 1 To 0 Step -1
        Me.cmbUnit.RemoveItem i
    Next i
    
    For i = Me.cmbSplit.ListCount - 1 To 0 Step -1
        Me.cmbSplit.RemoveItem i
    Next i
    
    Me.cmbUnit.AddItem "PC"
    Me.cmbUnit.AddItem "KG"
    Me.cmbUnit.AddItem "PAL"
    Me.cmbSecUnit.AddItem "Wsad"
    Me.cmbSplit.AddItem "Maszyna"
    Me.cmbSplit.AddItem "Brak"
    Me.cmbSplit = "Maszyna"
    Me.cmbUnit = "KG"
    Me.cmbSecUnit = "Wsad"
    Me.cmbSplit.Enabled = True
    Me.cmbUnit.Enabled = False
    Me.cmbSecUnit.Enabled = False
ElseIf Me.cmbType = "Pakowanie" Then
    For i = Me.cmbSecUnit.ListCount - 1 To 0 Step -1
        Me.cmbSecUnit.RemoveItem i
    Next i
    
    For i = Me.cmbUnit.ListCount - 1 To 0 Step -1
        Me.cmbUnit.RemoveItem i
    Next i
    
    For i = Me.cmbSplit.ListCount - 1 To 0 Step -1
        Me.cmbSplit.RemoveItem i
    Next i
    
    Me.cmbUnit.AddItem "PC"
    Me.cmbUnit.AddItem "KG"
    Me.cmbUnit.AddItem "BOX"
    Me.cmbUnit.AddItem "PAL"
    Me.cmbUnit = "PC"
    Me.cmbSecUnit.AddItem "KG"
    Me.cmbSecUnit.AddItem "BOX"
    Me.cmbSecUnit.AddItem "PAL"
    Me.cmbSecUnit = "PAL"
    Me.cmbSplit.AddItem "Maszyna"
    Me.cmbSplit.AddItem "Klient"
    Me.cmbSplit.AddItem "Kierunek"
    Me.cmbSplit.AddItem "Brak"
    Me.cmbSplit.Enabled = True
    Me.cmbUnit.Enabled = True
    Me.cmbSecUnit.Enabled = True
ElseIf Me.cmbType = "Palety" Then
    For i = Me.cmbSecUnit.ListCount - 1 To 0 Step -1
        Me.cmbSecUnit.RemoveItem i
    Next i
    
    For i = Me.cmbUnit.ListCount - 1 To 0 Step -1
        Me.cmbUnit.RemoveItem i
    Next i
    
    For i = Me.cmbSplit.ListCount - 1 To 0 Step -1
        Me.cmbSplit.RemoveItem i
    Next i
    
    Me.cmbUnit.AddItem "PC"
    Me.cmbUnit.AddItem "KG"
    Me.cmbUnit.AddItem "PAL"
    Me.cmbUnit = "PC"
    Me.cmbSecUnit.AddItem "PAL"
    Me.cmbSecUnit = "PAL"
    Me.cmbSplit.AddItem "Brak"
    Me.cmbSplit = "Brak"
    Me.cmbSplit.Enabled = True
    Me.cmbUnit.Enabled = True
    Me.cmbSecUnit.Enabled = True
ElseIf Me.cmbType = "Opakowania" Then
    For i = Me.cmbSecUnit.ListCount - 1 To 0 Step -1
        Me.cmbSecUnit.RemoveItem i
    Next i
    
    For i = Me.cmbUnit.ListCount - 1 To 0 Step -1
        Me.cmbUnit.RemoveItem i
    Next i
    
    For i = Me.cmbSplit.ListCount - 1 To 0 Step -1
        Me.cmbSplit.RemoveItem i
    Next i
    Me.cmbSecUnit.AddItem "Brak"
    Me.cmbSecUnit = "Brak"
    Me.cmbSplit.AddItem "Kategoria"
    Me.cmbSplit.AddItem "Brak"
    Me.cmbSplit.AddItem "Maszyna"
    Me.cmbSplit = "Kategoria"
    Me.cmbUnit.AddItem "Wg BOM"
    Me.cmbUnit = "Wg BOM"
    Me.cmbUnit.Enabled = False
    Me.cmbSecUnit.Enabled = False
    Me.cmbSplit.Enabled = True
End If
End Sub


Private Sub MultiPage1_Click(ByVal index As Long)
Select Case MultiPage1.SelectedItem.name
    Case "pgOptions": loadVersions
End Select
End Sub

Private Sub optCustom_Click()
Me.cmbWeek.Enabled = False
Me.cmbYear.Enabled = False
Me.txtFrom.Enabled = True
Me.txtTo.Enabled = True
Me.cmbTFrom.Enabled = True
Me.cmbTTo.Enabled = True
Me.weekSwitcher.Enabled = False
Me.yearSwitcher.Enabled = False

restoreVersion
End Sub

Private Sub optWeek_Click()
Me.cmbWeek.Enabled = True
Me.cmbYear.Enabled = True
Me.txtFrom.Enabled = False
Me.cmbTFrom.Enabled = False
Me.cmbTTo.Enabled = False
Me.txtTo.Enabled = False
Me.weekSwitcher.Enabled = True
Me.yearSwitcher.Enabled = True
End Sub

Private Sub weekSwitcher_Change()

If oldWeek > Me.weekSwitcher.value Then
    If Me.cmbWeek.ListIndex > 0 Then
        Me.cmbWeek.ListIndex = Me.cmbWeek.ListIndex - 1
        restoreVersion
    End If
ElseIf oldWeek < Me.weekSwitcher.value Then
    If Me.cmbWeek.ListIndex < 52 Then
        Me.cmbWeek.ListIndex = Me.cmbWeek.ListIndex + 1
        restoreVersion
    End If
End If

oldWeek = Me.weekSwitcher.value

End Sub


Private Sub yearSwitcher_Change()
If oldYear > Me.yearSwitcher.value Then
    If Me.cmbYear.ListIndex > 0 Then
        Me.cmbYear.ListIndex = Me.cmbYear.ListIndex - 1
        restoreVersion
    End If
ElseIf oldYear < Me.yearSwitcher.value Then
    If Me.cmbYear.ListIndex < 9 Then
        Me.cmbYear.ListIndex = Me.cmbYear.ListIndex + 1
        restoreVersion
    End If
End If

oldYear = Me.yearSwitcher.value
End Sub

Private Sub UserForm_Initialize()
Me.MultiPage1.value = 0
populator
bringSettings
End Sub

Private Sub bringSettings()
Dim week As Integer
Dim year As Integer
Dim i As Integer
Dim regval As Variant

regval = registryKeyExists(regPath & "LoadedWeek")
If regval <> False Then
    week = regval
Else
    week = DatePart("ww", Date, vbSunday, vbFirstFourDays)
End If

regval = registryKeyExists(regPath & "LoadedYear")
If regval <> False Then
    year = regval
Else
    year = DatePart("yyyy", Date, vbSunday, vbFirstFourDays)
End If

Me.cmbWeek = week
oldWeek = Me.cmbWeek.ListIndex + 1
Me.weekSwitcher.value = Me.cmbWeek.ListIndex + 1
Me.cmbYear = year
oldYear = Me.cmbYear.ListIndex + 1
Me.yearSwitcher.value = Me.cmbYear.ListIndex + 1

regval = registryKeyExists(regPath & "CustomRangeStartDate")
If regval <> False And regval <> "" Then
    Me.txtFrom = CDate(regval)
Else
    Me.txtFrom.value = Week2Date(CLng(week), CLng(year), vbSunday, vbFirstFourDays)
End If

regval = registryKeyExists(regPath & "CustomRangeEndDate")
If regval <> False And regval <> "" Then
    Me.txtTo = CDate(regval)
Else
    Me.txtTo.value = DateAdd("d", 6, Me.txtFrom.value)
End If

regval = registryKeyExists(regPath & "CustomRangeStartTime")
If regval <> False Then
    Me.cmbTFrom = regval
Else
    Me.cmbTFrom = "06:00"
End If

regval = registryKeyExists(regPath & "CustomRangeEndTime")
If regval <> False Then
    Me.cmbTTo = regval
Else
    Me.cmbTTo = "06:00"
End If

regval = registryKeyExists(regPath & "Type")
If regval <> False Then Me.cmbType = regval
cmbType_AfterUpdate
regval = registryKeyExists(regPath & "SplitBy")
If regval <> False Then Me.cmbSplit = regval
regval = registryKeyExists(regPath & "Charts")
If regval <> False Then Me.cmbChart = regval
regval = registryKeyExists(regPath & "Unit")
If regval <> False Then Me.cmbUnit = regval
regval = registryKeyExists(regPath & "SecondaryUnit")
If regval <> False Then Me.cmbSecUnit = regval
regval = registryKeyExists(regPath & "ProductProperties")
If regval <> False Then Me.cmbProductProperties = regval
regval = registryKeyExists(regPath & "ProductName")
If regval <> False Then Me.cmbProductName = regval
regval = registryKeyExists(regPath & "TotalShift")
If regval <> False Then Me.cmbTotalShift = regval
regval = registryKeyExists(regPath & "TotalDaily")
If regval <> False Then Me.cmbTotalDaily = regval
regval = registryKeyExists(regPath & "Customer")
If regval <> False Then Me.cmbCustomer = regval
regval = registryKeyExists(regPath & "Machine")
If regval <> False Then Me.cmbMachine = regval
regval = registryKeyExists(regPath & "RangeStart")
If regval <> False Then Me.cmbFirst = regval
regval = registryKeyExists(regPath & "UnitRatio")
If regval <> False Then Me.cmbRatio = regval
regval = registryKeyExists(regPath & "ShowComments")
If regval <> False Then Me.cmbComments = regval
regval = registryKeyExists(regPath & "TotalIndex")
If regval <> False Then Me.cmbTotalIndex = regval
regval = registryKeyExists(regPath & "ZfinList")
If regval <> False Then Me.txtZfinList = regval
regval = registryKeyExists(regPath & "dateRangeType")
If regval <> False Then
    If regval = "Weekly" Then
        Me.optWeek = True
    Else
        Me.optCustom = True
    End If
Else
    Me.optWeek = True
End If

If Me.cmbType = "Prażenie" Then
    Me.cmbSplit.Enabled = False
    Me.cmbUnit.Enabled = False
    Me.cmbSecUnit.Enabled = False
End If
'For i = 0 To Me.cmbWeek.ListCount - 1
'    If Me.cmbWeek.List(i) = week Then
'        Exit For
'    End If
'Next i
'Me.cmbWeek = i
'
'For i = 0 To Me.cmbYear.ListCount - 1
'    If Me.cmbYear.List(i) = year Then
'        Exit For
'    End If
'Next i
'Me.cmbYear = i
End Sub

Private Sub populator()
Dim i As Integer

For i = Me.cmbWeek.ListCount To 1 Step -1
    Me.cmbWeek.RemoveItem i
Next i

For i = 1 To 53
    Me.cmbWeek.AddItem i
Next i

For i = Me.cmbYear.ListCount To 1 Step -1
    Me.cmbYear.RemoveItem i
Next i

For i = 1 To 10
    Me.cmbYear.AddItem i + 2015
Next i

For i = Me.cmbType.ListCount To 1 Step -1
    Me.cmbType.RemoveItem i
Next i

Me.cmbType.AddItem "Pakowanie"
Me.cmbType.AddItem "Prażenie"
Me.cmbType.AddItem "Palety"
Me.cmbType.AddItem "Opakowania"
Me.cmbType.AddItem "Lista"

For i = Me.cmbSplit.ListCount To 1 Step -1
    Me.cmbSplit.RemoveItem i
Next i

Me.cmbSplit.AddItem "Klient"
Me.cmbSplit.AddItem "Kierunek"
Me.cmbSplit.AddItem "Maszyna"
Me.cmbSplit.AddItem "Brak"

For i = Me.cmbUnit.ListCount To 1 Step -1
    Me.cmbUnit.RemoveItem i
Next i

Me.cmbUnit.AddItem "PC"
Me.cmbUnit.AddItem "KG"
Me.cmbUnit.AddItem "BOX"

For i = Me.cmbChart.ListCount To 1 Step -1
    Me.cmbChart.RemoveItem i
Next i

Me.cmbChart.AddItem "Brak"
Me.cmbChart.AddItem "Dzienny"
Me.cmbChart.AddItem "Dzienny narastająco"
Me.cmbChart.AddItem "Zmianowy"
Me.cmbChart.AddItem "Zmianowy narastająco"

restoreVersion

For i = Me.cmbSecUnit.ListCount To 1 Step -1
    Me.cmbSecUnit.RemoveItem i
Next i

Me.cmbSecUnit.AddItem "KG"
Me.cmbSecUnit.AddItem "BOX"
Me.cmbSecUnit.AddItem "PAL"

For i = Me.cmbProductName.ListCount To 1 Step -1
    Me.cmbProductName.RemoveItem i
Next i

Me.cmbProductName.AddItem "Pokazuj"
Me.cmbProductName.AddItem "Nie pokazuj"

For i = Me.cmbProductProperties.ListCount To 1 Step -1
    Me.cmbProductProperties.RemoveItem i
Next i

Me.cmbProductProperties.AddItem "Pokazuj"
Me.cmbProductProperties.AddItem "Nie pokazuj"

For i = Me.cmbTotalShift.ListCount To 1 Step -1
    Me.cmbTotalShift.RemoveItem i
Next i

Me.cmbTotalShift.AddItem "Nie pokazuj"
Me.cmbTotalShift.AddItem "W jednostce podst."
Me.cmbTotalShift.AddItem "W jednostce pomoc."

For i = Me.cmbTotalDaily.ListCount To 1 Step -1
    Me.cmbTotalDaily.RemoveItem i
Next i

Me.cmbTotalDaily.AddItem "Nie pokazuj"
Me.cmbTotalDaily.AddItem "W jednostce podst."
Me.cmbTotalDaily.AddItem "W jednostce pomoc."

For i = Me.cmbCustomer.ListCount To 1 Step -1
    Me.cmbCustomer.RemoveItem i
Next i

Me.cmbCustomer.AddItem "Pokazuj"
Me.cmbCustomer.AddItem "Nie pokazuj"

For i = Me.cmbMachine.ListCount To 1 Step -1
    Me.cmbMachine.RemoveItem i
Next i

Me.cmbMachine.AddItem "Pokazuj"
Me.cmbMachine.AddItem "Nie pokazuj"

For i = Me.cmbBean.ListCount To 1 Step -1
    Me.cmbBean.RemoveItem i
Next i

Me.cmbBean.AddItem "Wszystkie"
Me.cmbBean.AddItem "Tylko ziarno"
Me.cmbBean.AddItem "Tylko mielona"

For i = Me.cmbDecaf.ListCount To 1 Step -1
    Me.cmbDecaf.RemoveItem i
Next i

Me.cmbDecaf.AddItem "Wszystkie"
Me.cmbDecaf.AddItem "Tylko kofeinowa"
Me.cmbDecaf.AddItem "Tylko bezkofeinowa"

For i = Me.cmbFirst.ListCount To 1 Step -1
    Me.cmbFirst.RemoveItem i
Next i

For i = 1 To 7
    Me.cmbFirst.AddItem StrConv(WeekdayName(i, , vbSunday), vbProperCase) & " 06:00"
    Me.cmbFirst.AddItem StrConv(WeekdayName(i, , vbSunday), vbProperCase) & " 14:00"
    Me.cmbFirst.AddItem StrConv(WeekdayName(i, , vbSunday), vbProperCase) & " 22:00"
Next i

For i = Me.cmbRatio.ListCount To 1 Step -1
    Me.cmbRatio.RemoveItem i
Next i


Me.cmbRatio.AddItem "Pokazuj"
Me.cmbRatio.AddItem "Nie pokazuj"

For i = Me.cmbComments.ListCount To 1 Step -1
    Me.cmbComments.RemoveItem i
Next i


Me.cmbComments.AddItem "Pokazuj"
Me.cmbComments.AddItem "Nie pokazuj"

For i = Me.cmbTFrom.ListCount To 1 Step -1
    Me.cmbTFrom.RemoveItem i
Next i

Me.cmbTFrom.AddItem "06:00"
Me.cmbTFrom.AddItem "14:00"
Me.cmbTFrom.AddItem "22:00"

For i = Me.cmbTTo.ListCount To 1 Step -1
    Me.cmbTTo.RemoveItem i
Next i

Me.cmbTTo.AddItem "06:00"
Me.cmbTTo.AddItem "14:00"
Me.cmbTTo.AddItem "22:00"

For i = Me.cmbTotalIndex.ListCount To 1 Step -1
    Me.cmbTotalIndex.RemoveItem i
Next i

Me.cmbTotalIndex.AddItem "Pokazuj"
Me.cmbTotalIndex.AddItem "Nie pokazuj"

End Sub


Private Sub createHeader(dFrom As Date, dTo As Date, Optional shtName As Variant)
Dim sht As Worksheet
Dim rKey As Variant
Dim x As Integer
Dim y As Integer
Dim totalCount As Integer
Dim rowsInHeader As Integer

On Error GoTo err_trap

If Not IsMissing(shtName) Then
    Set sht = ThisWorkbook.Sheets(shtName)
Else
    Set sht = ThisWorkbook.Sheets(1)
End If

currentDate = DateSerial(year(dFrom), Month(dFrom), Day(dFrom))

pUnit = registryKeyExists(regPath & "Unit")
sUnit = registryKeyExists(regPath & "SecondaryUnit")

rKey = registryKeyExists(regPath & "TotalShift")
If rKey <> False And rKey <> "Nie pokazuj" And rKey <> "" Then
    If rKey = "W jednostce podst." Then
        TotalShift = pUnit
    Else
        TotalShift = sUnit
    End If
Else
    TotalShift = False
End If
rKey = registryKeyExists(regPath & "TotalDaily")
If rKey <> False And rKey <> "Nie pokazuj" And rKey <> "" Then
    If rKey = "W jednostce podst." Then
        TotalDaily = pUnit
    Else
        TotalDaily = sUnit
    End If
Else
    TotalDaily = False
End If
rKey = registryKeyExists(regPath & "ProductName")
If rKey <> False And rKey <> "Nie pokazuj" And rKey <> "" Then
    showName = True
Else
    showName = False
End If
rKey = registryKeyExists(regPath & "ProductProperties")
If rKey <> False And rKey <> "Nie pokazuj" And rKey <> "" Then
    showProperties = True
Else
    showProperties = False
End If
rKey = registryKeyExists(regPath & "ShowComments")
If rKey <> False And rKey <> "Nie pokazuj" And rKey <> "" Then
    showComments = True
Else
    showComments = False
End If

rKey = registryKeyExists(regPath & "Customer")
If rKey <> False And rKey <> "Nie pokazuj" And rKey <> "" Then
    showCustomer = True
Else
    showCustomer = False
End If

rKey = registryKeyExists(regPath & "Machine")
If rKey <> False And rKey <> "Nie pokazuj" And rKey <> "" Then
    showMachine = True
Else
    showMachine = False
End If

rKey = registryKeyExists(regPath & "UnitRatio")
If rKey <> False And rKey <> "Nie pokazuj" And rKey <> "" Then
    showUnitRatio = True
Else
    showUnitRatio = False
End If
rKey = registryKeyExists(regPath & "TotalIndex")
If rKey <> False And rKey <> "Nie pokazuj" And rKey <> "" Then
    showTotalIndex = True
Else
    showTotalIndex = False
End If

x = 1
y = 1

rowsInHeader = 3
totalCount = 3

If TotalDaily <> False Then rowsInHeader = rowsInHeader + 1
If TotalShift <> False Then rowsInHeader = rowsInHeader + 1

totalCount = rowsInHeader - totalCount

sht.Range(sht.Cells(y, x), sht.Cells(rowsInHeader, x)).Merge
sht.Cells(y, x) = "Index"
sht.Cells(y, x).Interior.Color = vbWhite
sht.Cells(y, x).Font.Bold = True
sht.Cells(y, x).Font.Size = 10
sht.Cells(y, x).HorizontalAlignment = xlCenter
If showName Then
    x = x + 1
    sht.Range(sht.Cells(y, x), sht.Cells(rowsInHeader, x)).Merge
    sht.Cells(y, x) = "Nazwa"
    sht.Cells(y, x).Interior.Color = vbWhite
    sht.Cells(y, x).Font.Bold = True
    sht.Cells(y, x).Font.Size = 10
    sht.Cells(y, x).HorizontalAlignment = xlCenter
End If
If showCustomer Then
    x = x + 1
    sht.Range(sht.Cells(y, x), sht.Cells(rowsInHeader, x)).Merge
    sht.Cells(y, x) = "Klient"
    sht.Cells(y, x).Interior.Color = vbWhite
    sht.Cells(y, x).Font.Bold = True
    sht.Cells(y, x).Font.Size = 10
    sht.Cells(y, x).HorizontalAlignment = xlCenter
End If
If showMachine Then
    x = x + 1
    sht.Range(sht.Cells(y, x), sht.Cells(rowsInHeader, x)).Merge
    sht.Cells(y, x) = "Maszyna"
    sht.Cells(y, x).Interior.Color = vbWhite
    sht.Cells(y, x).Font.Bold = True
    sht.Cells(y, x).Font.Size = 10
    sht.Cells(y, x).HorizontalAlignment = xlCenter
End If
If showProperties Then
    x = x + 1
    sht.Range(sht.Cells(y, x), sht.Cells(rowsInHeader, x)).Merge
    sht.Cells(y, x) = "Atrybuty"
    sht.Cells(y, x).Interior.Color = vbWhite
    sht.Cells(y, x).Font.Bold = True
sht.Cells(y, x).Font.Size = 10
sht.Cells(y, x).HorizontalAlignment = xlCenter
End If
If showUnitRatio Then
    x = x + 1
    sht.Range(sht.Cells(y, x), sht.Cells(rowsInHeader, x)).Merge
    sht.Cells(y, x) = pUnit & "/" & sUnit
    sht.Cells(y, x).Interior.Color = vbWhite
    sht.Cells(y, x).Font.Bold = True
sht.Cells(y, x).Font.Size = 10
sht.Cells(y, x).HorizontalAlignment = xlCenter
End If
x = x + 1
sht.Range(sht.Cells(y, x), sht.Cells(rowsInHeader, x)).Merge
sht.Cells(y, x) = "j.m."
sht.Cells(y, x).Interior.Color = vbWhite
sht.Cells(y, x).Font.Bold = True
sht.Cells(y, x).Font.Size = 10
sht.Cells(y, x).HorizontalAlignment = xlCenter
If showTotalIndex Then
    x = x + 1
    sht.Range(sht.Cells(y, x), sht.Cells(rowsInHeader, x)).Merge
    sht.Cells(y, x) = "Suma"
    sht.Cells(y, x).Interior.Color = vbWhite
    sht.Cells(y, x).Font.Bold = True
    sht.Cells(y, x).Font.Size = 10
    sht.Cells(y, x).HorizontalAlignment = xlCenter
End If
x = x + 1
Do Until currentDate > dTo
    sht.Range(sht.Cells(y, x), sht.Cells(y, x + 2)).Merge
    sht.Cells(y, x) = StrConv(WeekdayName(weekday(currentDate, vbSunday), False, vbSunday), vbProperCase)
    sht.Cells(y, x).Font.Bold = True
    sht.Cells(y, x).Font.Size = 10
    sht.Cells(y, x).HorizontalAlignment = xlCenter
    y = y + 1
    sht.Range(sht.Cells(y, x), sht.Cells(y, x + 2)).Merge
    sht.Cells(y, x) = currentDate
    sht.Cells(y, x).Font.Bold = False
    sht.Cells(y, x).Font.Size = 8
    sht.Cells(y, x).HorizontalAlignment = xlCenter
    y = y + 1
    sht.Cells(y, x) = "I"
    sht.Cells(y, x).Font.Bold = True
    sht.Cells(y, x).Font.Size = 10
    sht.Cells(y, x).HorizontalAlignment = xlCenter
    sht.Cells(y, x + 1) = "II"
    sht.Cells(y, x + 1).Font.Bold = True
    sht.Cells(y, x + 1).Font.Size = 10
    sht.Cells(y, x + 1).HorizontalAlignment = xlCenter
    sht.Cells(y, x + 2) = "III"
    sht.Cells(y, x + 2).Font.Bold = True
    sht.Cells(y, x + 2).Font.Size = 10
    sht.Cells(y, x + 2).HorizontalAlignment = xlCenter
    If TotalDaily <> False Then
        y = y + 1
        sht.Range(sht.Cells(y, x), sht.Cells(y, x + 2)).Merge
        sht.Cells(y, x).Font.Bold = False
        sht.Cells(y, x).Font.Size = 10
        sht.Cells(y, x).HorizontalAlignment = xlCenter
    End If
    If TotalShift <> False Then
        y = y + 1
        sht.Range(sht.Cells(y, x), sht.Cells(y, x + 2)).Font.Bold = False
        sht.Range(sht.Cells(y, x), sht.Cells(y, x + 2)).Font.Size = 10
        sht.Range(sht.Cells(y, x), sht.Cells(y, x + 2)).HorizontalAlignment = xlCenter
    End If
    currentDate = DateAdd("d", 1, currentDate)
    y = 1
    x = x + 3
Loop
headerAddress = sht.Range(sht.Cells(1, 1), sht.Cells(3, x)).Address

exit_here:
Exit Sub

err_trap:
MsgBox "Error in ""CreateHeader"" of updater. Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Sub

Private Sub getRoastingBatches()
Dim sql As String
Dim rs As ADODB.Recordset
Dim zfor As clsZfor
Dim bool As Boolean
Dim rn3 As Variant
Dim rn4 As Variant

On Error GoTo err_trap

bool = False

If Not roastingBatches Is Nothing Then
    If roastingBatches.Count = 0 Then
        bool = True
    End If
Else
    Set roastingBatches = New Collection
    bool = True
End If

If bool Then
    connectScada
    sql = "select DISTINCT zlec.MaterialNumber, " _
        & "(SELECT AVG(sub.green) FROM " _
        & "(SELECT TOP(10) z.SUMA_ZIELONEJ as green FROM ZLECENIA_PALONA z Join ZLECENIAWARTOSCI w ON z.IDZLECENIE = w.IDZLECENIE " _
                                & "JOIN ZLECENIA zl on w.IDZLECENIE = zl.IDZLECENIE " _
                                & "WHERE z.NUMERPIECA=3000 AND z.SUMA_ZIELONEJ>100 AND zlec.MaterialNumber=zl.MaterialNumber ORDER BY z.DTZAPIS DESC) sub) As RN3000, " _
        & "(SELECT AVG(sub.green) FROM " _
        & "(SELECT TOP(10) z.SUMA_ZIELONEJ as green FROM ZLECENIA_PALONA z Join ZLECENIAWARTOSCI w ON z.IDZLECENIE = w.IDZLECENIE " _
                                & "JOIN ZLECENIA zl on w.IDZLECENIE = zl.IDZLECENIE " _
                                & "WHERE z.NUMERPIECA=4000 AND z.SUMA_ZIELONEJ>100 AND zlec.MaterialNumber=zl.MaterialNumber ORDER BY z.DTZAPIS DESC) sub) As RN4000 " _
    & "FROM ZLECENIA zlec;"
    Set rs = CreateObject("adodb.recordset")
    rs.Open sql, scadaConn
    If Not rs.EOF Then
        rs.MoveFirst
        Do Until rs.EOF
            Set zfor = New clsZfor
            rn3 = rs.Fields("RN3000")
            rn4 = rs.Fields("RN4000")
            If IsNumeric(rn3) Then
                rn3 = Round(rn3, 1)
            End If
            If IsNumeric(rn4) Then
                rn4 = Round(rn4, 1)
            End If
            zfor.initialize rs.Fields("MaterialNumber"), rn3, rn4
            roastingBatches.add zfor, rs.Fields("MaterialNumber")
            rs.MoveNext
        Loop
    End If
    rs.Close
End If

exit_here:
Set rs = Nothing
disconnectScada
Exit Sub

err_trap:
MsgBox "Error in ""getRoastingBatches"" of updater. Error number: " & Err.Number & ", " & Err.Description, vbOKOnly + vbCritical, "Error"
Resume exit_here

End Sub

Private Sub loadVersions()
Dim sql As String
Dim dFrom As Date
Dim dTo As Date
Dim v() As String
Dim dday As Integer
Dim hhour As Integer
Dim i As Integer
Dim rs As ADODB.Recordset

On Error GoTo err_trap

If Me.optWeek = True Then
    updateConnection
    v = split(Me.cmbFirst, " ", , vbTextCompare)
    dday = WeekDayName2Int(v(0))
    hhour = CInt(Left(v(1), 2))
    dFrom = DateAdd("h", hhour, Week2Date(CLng(Me.cmbWeek), CLng(Me.cmbYear), dday, vbFirstFourDays))
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
        For i = Me.cmbVersion.ListCount - 1 To 0 Step -1
            Me.cmbVersion.RemoveItem i
        Next i
        
        Me.cmbVersion.AddItem 0
        Me.cmbVersion.Column(1, 0) = "Najnowsza"
        rs.MoveFirst
        
        i = 1
        
        Do Until rs.EOF
            Me.cmbVersion.AddItem rs.Fields("operDataVer")
            Me.cmbVersion.Column(1, i) = rs.Fields("dayOfWeek")
            Me.cmbVersion.Column(2, i) = rs.Fields("createdOn")
            i = i + 1
            rs.MoveNext
        Loop
        Me.cmbVersion.ColumnWidths = "0 pt;85 pt; 85 pt"
    End If
    
    rs.Close

End If

exit_here:
Set rs = Nothing
closeConnection
Exit Sub

err_trap:
MsgBox "Error in ""loadVersions"" of updater. Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Sub

Private Sub restoreVersion()
 For i = Me.cmbVersion.ListCount - 1 To 0 Step -1
    Me.cmbVersion.RemoveItem i
Next i

Me.cmbVersion.AddItem 0
Me.cmbVersion.Column(1, 0) = "Najnowsza"
Me.cmbVersion = 0
End Sub
