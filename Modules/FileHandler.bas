Attribute VB_Name = "FileHandler"
Function FileExists(ByVal strFile As String, Optional bFindFolders As Boolean) As Boolean
    'Purpose:   Return True if the file exists, even if it is hidden.
    'Arguments: strFile: File name to look for. Current directory searched if no path included.
    '           bFindFolders. If strFile is a folder, FileExists() returns False unless this argument is True.
    'Note:      Does not look inside subdirectories for the file.
    'Author:    Allen Browne. http://allenbrowne.com June, 2006.
    Dim lngAttributes As Long

    'Include read-only files, hidden files, system files.
    lngAttributes = (vbReadOnly Or vbHidden Or vbSystem)

    If bFindFolders Then
        lngAttributes = (lngAttributes Or vbDirectory) 'Include folders as well.
    Else
        'Strip any trailing slash, so Dir does not look inside the folder.
        Do While Right$(strFile, 1) = "\"
            strFile = Left$(strFile, Len(strFile) - 1)
        Loop
    End If

    'If Dir() returns something, the file exists.
    On Error Resume Next
    FileExists = (Len(Dir(strFile, lngAttributes)) > 0)
End Function

Function FolderExists(strPath As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(strPath) And vbDirectory) = vbDirectory)
End Function

Public Function filePicker(multiselect As Boolean, Optional title As Variant, Optional filter As Variant) As Variant
Dim f As Object
Dim temp() As String
Dim varFile As Variant
Set f = Application.FileDialog(3)
f.AllowMultiSelect = multiselect
If Not IsMissing(title) Then
    f.title = title
Else
    f.title = "Wybierz plik"
End If

filePicker = False

With f
    .Filters.clear
    .Filters.add "Pliki bazy danych Ms Access", "*.mdb; *.accdb"
'   ' Show the dialog box. If the .Show method returns True, the '
'   ' user picked at least one file. If the .Show method returns '
'   ' False, the user clicked Cancel. '
   If .Show = True Then

      'Loop through each file selected and add it to our list box. '
      For Each varFile In .SelectedItems
        If isArrayEmpty(temp) Then
           ReDim temp(0) As String
           temp(0) = CStr(varFile)
        Else
            ReDim Preserve temp(UBound(temp) + 1) As String
            temp(UBound(temp)) = CStr(varFile)
        End If
      Next

   End If
End With

If Not isArrayEmpty(temp) Then
    filePicker = temp
End If

End Function
