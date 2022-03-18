VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pPrinterPort As String
Private pPrinterName As String
Private pTrays As New Collection

Public Property Let PrinterPort(val As String)
pPrinterPort = val
End Property

Public Property Let PrinterName(val As String)
pPrinterName = val
End Property

Public Property Get PrinterPort() As String
PrinterPort = pPrinterPort
End Property

Public Property Get PrinterName() As String
PrinterName = pPrinterName
End Property

Public Function getTrays() As Collection
Set getTrays = pTrays
End Function

Public Function PriterString() As String
Dim defaultPrinter As String

On Error GoTo err_trap

defaultPrinter = Application.ActivePrinter

PriterString = pPrinterName & " na " & pPrinterPort & ":"
Application.ActivePrinter = PriterString

exit_here:
Application.ActivePrinter = defaultPrinter
Exit Function

err_trap:
PriterString = pPrinterName & " on " & pPrinterPort & ":"
Resume exit_here

End Function

Public Sub initialize()
Dim tray As clsTray
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
    strDeviceName = pPrinterName
    strDevicePort = pPrinterPort
    
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
            'strMsg = strMsg & vbCrLf & aintNumBin(lngCounter) _
                & vbTab & strBinName
            Set tray = New clsTray
            tray.TrayId = aintNumBin(lngCounter)
            tray.TrayName = strBinName
            pTrays.add tray, strBinName
        Next lngCounter
    End If

    
GetBinList_End:
    Exit Sub
GetBinList_Err:
    MsgBox Prompt:=Err.Description, Buttons:=vbCritical & vbOKOnly, _
        title:="Error Number " & Err.Number & " Occurred"
    Resume GetBinList_End
End Sub

