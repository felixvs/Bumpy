Attribute VB_Name = "Functions"
'Extract the name of the file from the path.
Function extract_file_name(file As Variant) As String
    If file <> False Then
        While InStr(1, file, "\") <> 0
            file = Right(file, Len(file) - InStr(1, file, "\"))
        Wend
        extract_file_name = file
    End If
End Function

'Matching name and phone number between Vivial Force and Galley
Function match_name_phone(search_value_vivialforce As Variant, search_array_galley_all As Range) As Variant
    'If not error run the match
    If Not IsError(Application.Match(search_value_vivialforce, search_array_galley_all, 0)) Then
        match_name_phone = "MATCHED"
    Else
        match_name_phone = "NOT MATCHED"
    End If
End Function
    
'Verify it the file is open
Function IsFileOpen(filename As Variant)
    Dim filenum As Integer, errnum As Integer

    On Error Resume Next   ' Turn error checking off.
    filenum = FreeFile()   ' Get a free file number.
     'Attempt to open the file and lock it.
    Open filename For Input Lock Read As #filenum
    Close filenum          ' Close the file.
    errnum = Err           ' Save the error number that occurred.
    On Error GoTo 0        ' Turn error checking back on.

     'Check to see which error occurred.
    Select Case errnum

         'No error occurred.
         'File is NOT already open by another user.
        Case 0
         IsFileOpen = False

         'Error number for "Permission Denied."
         'File is already opened by another user.
        Case 70
            IsFileOpen = True

         'Another error occurred.
        Case Else
            Error errnum
    End Select
End Function
