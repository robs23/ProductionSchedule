Attribute VB_Name = "Registry"
Public Function AddTrustedLocation(location As String)
On Error GoTo err_proc
'WARNING:  THIS CODE MODIFIES THE REGISTRY
'sets registry key for 'trusted location'

  Dim intLocns As Integer
  Dim i As Integer
  Dim intNotUsed As Integer
  Dim strLnKey As String
  Dim reg As Object
  Dim strPath As String
  Dim strTitle As String
  
  allowNetworkLocations
  
  strTitle = "Add Trusted Location"
  Set reg = CreateObject("wscript.shell")
  strPath = location

  'Specify the registry trusted locations path for the version of Access used
  strLnKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & Application.Version & "\Access\Security\Trusted Locations\Location"

On Error GoTo err_proc0
  'find top of range of trusted locations references in registry
  For i = 999 To 0 Step -1
      reg.RegRead strLnKey & i & "\Path"
      GoTo chckRegPths        'Reg.RegRead successful, location exists > check for path in all locations 0 - i.
checknext:
  Next
  MsgBox "Unexpected Error - No Registry Locations found", vbExclamation
  GoTo exit_proc
  
  
chckRegPths:
'Check if Currentdb path already a trusted location
'reg.RegRead fails before intlocns = i then the registry location is unused and
'will be used for new trusted location if path not already in registy

On Error GoTo err_proc1:
  For intLocns = 1 To i
      reg.RegRead strLnKey & intLocns & "\Path"
      Debug.Print reg.RegRead(strLnKey & intLocns & "\Path")
      'If Path already in registry -> exit
      If InStr(1, reg.RegRead(strLnKey & intLocns & "\Path"), strPath) = 1 Then GoTo exit_proc
NextLocn:
  Next
  
  If intLocns = 999 Then
      MsgBox "Location count exceeded - unable to write trusted location to registry", vbInformation, strTitle
      GoTo exit_proc
  End If
  'if no unused location found then set new location for path
  If intNotUsed = 0 Then intNotUsed = i + 1
  
'Write Trusted Location regstry key to unused location in registry
On Error GoTo err_proc:
  strLnKey = strLnKey & intNotUsed & "\"
  reg.RegWrite strLnKey & "AllowSubfolders", 1, "REG_DWORD"
  reg.RegWrite strLnKey & "Date", Now(), "REG_SZ"
  reg.RegWrite strLnKey & "Description", Application.CurrentProject.name, "REG_SZ"
  reg.RegWrite strLnKey & "Path", strPath, "REG_SZ"
  
exit_proc:
  Set reg = Nothing
  Exit Function
  
err_proc0:
  Resume checknext
  
err_proc1:
  If intNotUsed = 0 Then intNotUsed = intLocns
  Resume NextLocn

err_proc:
  MsgBox Err.Description, , strTitle
  Resume exit_proc
  
End Function


Public Function allowNetworkLocations()
On Error GoTo err_proc
'WARNING:  THIS CODE MODIFIES THE REGISTRY
'sets registry key for 'trusted location'

Dim intLocns As Integer
Dim i As Integer
Dim intNotUsed As Integer
Dim strLnKey As String
Dim reg As Object
Dim strPath As String
Dim strTitle As String

Set reg = CreateObject("wscript.shell")
strLnKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & Application.Version & "\Access\Security\Trusted Locations"
  
On Error GoTo err_proc0
reg.RegRead strLnKey & "\AllowNetworkLocations"
If reg.RegRead(strLnKey & "\AllowNetworkLocations") <> 1 Then
    GoTo err_proc0
End If
GoTo exit_proc

exit_proc:
Set reg = Nothing
Exit Function

err_proc:
MsgBox "Error " & Err.Number & ". " & Err.Description
Resume exit_proc

err_proc0:
On Error GoTo err_proc
strLnKey = strLnKey & "\"
reg.RegWrite strLnKey & "AllowNetworkLocations", 1, "REG_DWORD"
GoTo exit_proc

End Function

Public Sub updateRegistry(key As String, value As Variant)
'key in form e.g. "HKEY_CURRENT_USER\Software\RWsoft\backEndPath"
Dim reg As Object 'registry itself
Dim theType As Variant

On Error GoTo err_trap


Select Case VarType(value)
    Case 0 To 1
    theType = Null
    Case 2
    theType = "REG_DWORD"
    Case 3
    theType = "REG_QWORD"
    Case 7
    value = CLng(value)
    theType = "REG_QWORD"
    Case 8
    theType = "REG_SZ"
    Case 11
    theType = "REG_BINARY"
    Case Else
    theType = Null
End Select

If theType = Null Then
    MsgBox "Type of variable ""value"" passed to ""createRegistryKey"" could not be determined or is unsuported. No key has been created", vbOKOnly + vbExclamation
Else
    Set reg = CreateObject("WScript.Shell")
    reg.RegWrite key, value, theType
End If


Exit_here:
Exit Sub

err_trap:
MsgBox "Error in ""updateRegistry"". Error number: " & Err.Number & ", " & Err.Description
Resume Exit_here

End Sub

Public Function registryKeyExists(key As String) As Variant
Dim bool As Variant
Dim reg As Variant

On Error GoTo err_trap

bool = False

reg = CreateObject("WScript.Shell").RegRead(key)

bool = reg

Exit_here:
registryKeyExists = bool
Exit Function

err_trap:
Resume Exit_here

End Function


