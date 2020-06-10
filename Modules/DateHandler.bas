Attribute VB_Name = "DateHandler"
Public Function Week2Date(weekNo As Long, Optional ByVal Yr As Long = 0, Optional ByVal DOW As VBA.VbDayOfWeek = VBA.VbDayOfWeek.vbUseSystemDayOfWeek, Optional ByVal FWOY As VBA.VbFirstWeekOfYear = VBA.VbFirstWeekOfYear.vbUseSystem) As Date
 ' Returns First Day of week
 Dim Jan1 As Date
 Dim Sub1 As Boolean
 Dim ret As Date

 If Yr = 0 Then
   Jan1 = VBA.DateSerial(VBA.year(VBA.Date()), 1, 1)
 Else
   Jan1 = VBA.DateSerial(Yr, 1, 1)
 End If
 Sub1 = (VBA.Format(Jan1, "ww", DOW, FWOY) = 1)
 ret = VBA.DateAdd("ww", weekNo + Sub1, Jan1)
 ret = ret - VBA.weekday(ret, DOW) + 1
 Week2Date = ret
End Function

Public Function IsoWeekNumber(InDate As Date) As Long
    IsoWeekNumber = DatePart("ww", InDate, vbMonday, vbFirstFourDays)
End Function

Public Function WeekDayName2Int(weekday As String) As Integer
Dim temp As Integer

Select Case weekday
    Case Is = "Niedziela"
    temp = 1
    Case Is = "Poniedziałek"
    temp = 2
    Case Is = "Wtorek"
    temp = 3
    Case Is = "Środa"
    temp = 4
    Case Is = "Czwartek"
    temp = 5
    Case Is = "Piątek"
    temp = 6
    Case Is = "Sobota"
    temp = 7
    Case Else
    temp = 0
End Select

WeekDayName2Int = temp
End Function



