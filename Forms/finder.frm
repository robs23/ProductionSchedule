VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} finder 
   Caption         =   "Szukaj"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7215
   OleObjectBlob   =   "finder.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "finder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private fRecords As New Collection 'collection of filtered records
Private keyWord As String
Private records As New Collection

Private Sub lstResults_Click()
Dim i As Integer
i = Me.lstResults.ListIndex
ThisWorkbook.Sheets(records(Me.lstResults.Column(0, i)).sheet).Select
ThisWorkbook.Sheets(records(Me.lstResults.Column(0, i)).sheet).Range(records(Me.lstResults.Column(0, i)).location.Address).Select
Me.Hide

End Sub

Private Sub txtSearch_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal shift As Integer)
keyWord = Me.txtSearch
If Len(keyWord) >= 1 Then
    Set fRecords = currentSchedule.filterRecordset(keyWord)
    rs2list fRecords
Else
    rs2list records
End If
End Sub

Private Sub UserForm_Activate()
Me.txtSearch.SetFocus
End Sub


Private Sub UserForm_Initialize()
Dim cRecord As clsRecord

Set records = currentSchedule.getRecords
rs2list records

End Sub

Private Sub rs2list(records As Collection)
If records.Count > 0 Then
    With Me.lstResults
        .clear
        i = 0
        For Each cRecord In records
            .AddItem CStr(cRecord.index & "_" & cRecord.sheet)
            .Column(1, i) = cRecord.index
            .Column(2, i) = cRecord.name
            .Column(3, i) = cRecord.sheet
            i = i + 1
        Next cRecord
    End With
End If
End Sub
