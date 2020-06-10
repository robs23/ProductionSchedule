Attribute VB_Name = "PrintHandler"
Option Explicit
Public printersCollection As New Collection
Private Const HKEY_CURRENT_USER As Long = &H80000001
Private Const HKCU = HKEY_CURRENT_USER
Private Const KEY_QUERY_VALUE = &H1&
Private Const ERROR_NO_MORE_ITEMS = 259&
Private Const ERROR_MORE_DATA = 234

' Declaration for the DeviceCapabilities function API call.
Public Declare Function DeviceCapabilities Lib "winspool.drv" _
    Alias "DeviceCapabilitiesA" (ByVal lpsDeviceName As String, _
    ByVal lpPort As String, ByVal iIndex As Long, lpOutput As Any, _
    ByVal lpDevMode As Long) As Long
    
' DeviceCapabilities function constants.
Public Const DC_PAPERNAMES = 16
Public Const DC_PAPERS = 2
Public Const DC_BINNAMES = 12
Public Const DC_BINS = 6
Public Const DEFAULT_VALUES = 0


Private Declare Function RegOpenKeyEx Lib "advapi32" _
    Alias "RegOpenKeyExA" ( _
    ByVal HKey As Long, _
    ByVal lpSubKey As String, _
    ByVal ulOptions As Long, _
    ByVal samDesired As Long, _
    phkResult As Long) As Long

Private Declare Function RegEnumValue Lib "advapi32.dll" _
    Alias "RegEnumValueA" ( _
    ByVal HKey As Long, _
    ByVal dwIndex As Long, _
    ByVal lpValueName As String, _
    lpcbValueName As Long, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    lpData As Byte, _
    lpcbData As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" ( _
    ByVal HKey As Long) As Long

Public Function GetPrinterFullNames() As String()
Dim Printers() As String ' array of names to be returned
Dim PNdx As Long    ' index into Printers()
Dim HKey As Long    ' registry key handle
Dim Res As Long     ' result of API calls
Dim Ndx As Long     ' index for RegEnumValue
Dim ValueName As String ' name of each value in the printer key
Dim ValueNameLen As Long    ' length of ValueName
Dim DataType As Long        ' registry value data type
Dim ValueValue() As Byte    ' byte array of registry value value
Dim ValueValueS As String   ' ValueValue converted to String
Dim CommaPos As Long        ' position of comma character in ValueValue
Dim ColonPos As Long        ' position of colon character in ValueValue
Dim M As Long               ' string index

' registry key in HCKU listing printers
Const PRINTER_KEY = "Software\Microsoft\Windows NT\CurrentVersion\Devices"

PNdx = 0
Ndx = 0
' assume printer name is less than 256 characters
ValueName = String$(256, Chr(0))
ValueNameLen = 255
' assume the port name is less than 1000 characters
ReDim ValueValue(0 To 999)
' assume there are less than 1000 printers installed
ReDim Printers(1 To 1000)

' open the key whose values enumerate installed printers
Res = RegOpenKeyEx(HKCU, PRINTER_KEY, 0&, _
    KEY_QUERY_VALUE, HKey)
' start enumeration loop of printers
Res = RegEnumValue(HKey, Ndx, ValueName, _
    ValueNameLen, 0&, DataType, ValueValue(0), 1000)
' loop until all values have been enumerated
Do Until Res = ERROR_NO_MORE_ITEMS
    M = InStr(1, ValueName, Chr(0))
    If M > 1 Then
        ' clean up the ValueName
        ValueName = Left(ValueName, M - 1)
    End If
    ' find position of a comma and colon in the port name
    CommaPos = InStr(1, ValueValue, ",")
    ColonPos = InStr(1, ValueValue, ":")
    ' ValueValue byte array to ValueValueS string
    On Error Resume Next
    ValueValueS = Mid(ValueValue, CommaPos + 1, ColonPos - CommaPos)
    On Error GoTo 0
    ' next slot in Printers
    PNdx = PNdx + 1
    Printers(PNdx) = ValueName & " on " & ValueValueS
    ' reset some variables
    ValueName = String(255, Chr(0))
    ValueNameLen = 255
    ReDim ValueValue(0 To 999)
    ValueValueS = vbNullString
    ' tell RegEnumValue to get the next registry value
    Ndx = Ndx + 1
    ' get the next printer
    Res = RegEnumValue(HKey, Ndx, ValueName, ValueNameLen, _
        0&, DataType, ValueValue(0), 1000)
    ' test for error
    If (Res <> 0) And (Res <> ERROR_MORE_DATA) Then
        Exit Do
    End If
Loop
' shrink Printers down to used size
ReDim Preserve Printers(1 To PNdx)
Res = RegCloseKey(HKey)
' Return the result array
GetPrinterFullNames = Printers
End Function

Sub Test()
    Dim Printers() As String
    Dim N As Long
    Dim S As String
    Printers = GetPrinterFullNames()
    For N = LBound(Printers) To UBound(Printers)
        S = S & Printers(N) & vbNewLine
    Next N
    MsgBox S, vbOKOnly, "Printers"
End Sub

Public Sub getPrinters()
Dim i As Integer
Dim Printers() As String ' array of names to be returned
Dim PNdx As Long    ' index into Printers()
Dim HKey As Long    ' registry key handle
Dim Res As Long     ' result of API calls
Dim Ndx As Long     ' index for RegEnumValue
Dim ValueName As String ' name of each value in the printer key
Dim ValueNameLen As Long    ' length of ValueName
Dim DataType As Long        ' registry value data type
Dim ValueValue() As Byte    ' byte array of registry value value
Dim ValueValueS As String   ' ValueValue converted to String
Dim CommaPos As Long        ' position of comma character in ValueValue
Dim ColonPos As Long        ' position of colon character in ValueValue
Dim M As Long               ' string index
Dim pnt As clsPrinter

'check if printers exists

If printersCollection.Count = 0 Then

    ' registry key in HCKU listing printers
    Const PRINTER_KEY = "Software\Microsoft\Windows NT\CurrentVersion\Devices"
    
    PNdx = 0
    Ndx = 0
    ' assume printer name is less than 256 characters
    ValueName = String$(256, Chr(0))
    ValueNameLen = 255
    ' assume the port name is less than 1000 characters
    ReDim ValueValue(0 To 999)
    ' assume there are less than 1000 printers installed
    ReDim Printers(1 To 1000)
    
    ' open the key whose values enumerate installed printers
    Res = RegOpenKeyEx(HKCU, PRINTER_KEY, 0&, _
        KEY_QUERY_VALUE, HKey)
    ' start enumeration loop of printers
    Res = RegEnumValue(HKey, Ndx, ValueName, _
        ValueNameLen, 0&, DataType, ValueValue(0), 1000)
    ' loop until all values have been enumerated
    Do Until Res = ERROR_NO_MORE_ITEMS
        M = InStr(1, ValueName, Chr(0))
        If M > 1 Then
            ' clean up the ValueName
            ValueName = Left(ValueName, M - 1)
        End If
        ' find position of a comma and colon in the port name
        CommaPos = InStr(1, ValueValue, ",")
        ColonPos = InStr(1, ValueValue, ":")
        ' ValueValue byte array to ValueValueS string
        On Error Resume Next
        ValueValueS = Mid(ValueValue, CommaPos + 1, ColonPos - CommaPos - 1)
        On Error GoTo 0
        ' next slot in Printers
        Set pnt = New clsPrinter
        pnt.PrinterName = ValueName
        pnt.PrinterPort = ValueValueS
        pnt.initialize
        printersCollection.add pnt, ValueName
        PNdx = PNdx + 1
        Printers(PNdx) = ValueName & " on " & ValueValueS
        ' reset some variables
        ValueName = String(255, Chr(0))
        ValueNameLen = 255
        ReDim ValueValue(0 To 999)
        ValueValueS = vbNullString
        ' tell RegEnumValue to get the next registry value
        Ndx = Ndx + 1
        ' get the next printer
        Res = RegEnumValue(HKey, Ndx, ValueName, ValueNameLen, _
            0&, DataType, ValueValue(0), 1000)
        ' test for error
        If (Res <> 0) And (Res <> ERROR_MORE_DATA) Then
            Exit Do
        End If
    Loop
    ' shrink Printers down to used size
    ReDim Preserve Printers(1 To PNdx)
    Res = RegCloseKey(HKey)
End If

End Sub

Public Function GetBinList(strName As String, portStr As String) As String
' Uses the DeviceCapabilities API function to display a
' message box with the name of the default printer and a
' list of the paper bins it supports.

    Dim lngBinCount As Long
    Dim lngCounter As Long
    Dim hPrinter As Long
    Dim strDeviceName As String
    Dim strDevicePort As String
    Dim strBinNamesList As String
    Dim strBinName As String
    Dim intLength As Integer
    Dim strMsg As String
    Dim aintNumBin() As Integer
    
    On Error GoTo GetBinList_Err
    
    ' Get name and port of the default printer.
    strDeviceName = strName
    strDevicePort = portStr
    
    ' Get count of paper bin names supported by the printer.
    lngBinCount = DeviceCapabilities(lpsDeviceName:=strDeviceName, _
        lpPort:=strDevicePort, _
        iIndex:=DC_BINNAMES, _
        lpOutput:=ByVal vbNullString, _
        lpDevMode:=DEFAULT_VALUES)
    
    ' Re-dimension the array to count of paper bins.
    If lngBinCount > 0 Then
        ReDim aintNumBin(1 To lngBinCount)
        
        ' Pad variable to accept 24 bytes for each bin name.
        strBinNamesList = String(Number:=24 * lngBinCount, Character:=0)
    
        ' Get string buffer of paper bin names supported by the printer.
        lngBinCount = DeviceCapabilities(lpsDeviceName:=strDeviceName, _
            lpPort:=strDevicePort, _
            iIndex:=DC_BINNAMES, _
            lpOutput:=ByVal strBinNamesList, _
            lpDevMode:=DEFAULT_VALUES)
            
        ' Get array of paper bin numbers supported by the printer.
        lngBinCount = DeviceCapabilities(lpsDeviceName:=strDeviceName, _
            lpPort:=strDevicePort, _
            iIndex:=DC_BINS, _
            lpOutput:=aintNumBin(1), _
            lpDevMode:=0)
            
        ' List available paper bin names.
        strMsg = ""
        For lngCounter = 1 To lngBinCount
            
            ' Parse a paper bin name from string buffer.
            strBinName = Mid(String:=strBinNamesList, _
                start:=24 * (lngCounter - 1) + 1, _
                length:=24)
            intLength = VBA.InStr(start:=1, _
                String1:=strBinName, String2:=Chr(0)) - 1
            strBinName = Left(String:=strBinName, _
                    length:=intLength)
    
            ' Add bin name and number to text string for message box.
            strMsg = strMsg & vbCrLf & aintNumBin(lngCounter) _
                & vbTab & strBinName
                
        Next lngCounter
    End If
    GetBinList = strMsg
    ' Show paper bin numbers and names in message box.
    
GetBinList_End:
    Exit Function
GetBinList_Err:
    MsgBox Prompt:=Err.Description, Buttons:=vbCritical & vbOKOnly, _
        title:="Error Number " & Err.Number & " Occurred"
    Resume GetBinList_End
End Function

