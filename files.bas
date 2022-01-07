Attribute VB_Name = "files"
'Get VF files
Public Sub VFiles()
    Dim WS As Worksheet, PR As Workbook, WK As Workbook, files As Variant
    Set WK = ThisWorkbook

    'Get the files
    file = Application.GetOpenFilename(Title:="Choose a file to add", MultiSelect:=False)

    If file <> False Then
    
        'Files names
        Dim search_array_files_names As Range
        Set search_array_files_names = WK.Worksheets("VFile").Range("AV:AV")
        
        'Get file names
        Form.VFTextBox.Value = extract_file_name(file)
        
        'Open the files
        Workbooks.Open filename:=file
        Set PR = Workbooks(Workbooks.count)
        
        'Validating I'ts the correct VF file
        If PR.Worksheets(1).Range("AD1").Value2 = "Name" Then
            
            'Take the last row from the two files
            lastrow = WK.Worksheets("VFile").Range("D" & Rows.count).End(xlUp).Row
            lastrow2 = PR.Worksheets(1).Range("D" & Rows.count).End(xlUp).Row
            
            If lastrow > 1 Then
                
                'Match the name of the file before add it to the sheet
                matched_file = match_name_phone(Form.VFTextBox.Value, search_array_files_names)
                
                If matched_file = "MATCHED" Then
                    MsgBox "This file was added previously:" & " " & Form.VFTextBox.Value & vbCrLf & "Please try with a new file..."
                    PR.Close savechanges = False
                    Exit Sub
                End If
                
                'copy the others files to the working workbook
                PR.Worksheets(1).Range("A1:AP" & lastrow2).Copy Destination:=WK.Worksheets("VFile").Range("A" & lastrow + 1)
                MsgBox "File added : " & Form.VFTextBox.Value
            Else
                'copy the first file to the working workbook
                PR.Worksheets(1).Range("A:AP").Copy Destination:=WK.Worksheets("VFile").Range("A:AP")
            End If
        Else
            MsgBox "You Choose a wrong VF file. Please try again:"
            PR.Close savechanges = False
            Exit Sub
        End If
        
        'Save file name
        lastrow = WK.Worksheets("VFile").Range("AV" & Rows.count).End(xlUp).Row
        WK.Worksheets("VFile").Range("AV" & last_row + 1).Value = Form.VFTextBox.Value

        'Close the VF File
        PR.Close savechanges = False

        'Show image
        Form.Image1.Visible = True
    End If

End Sub

'Get Galley files
Public Sub GalleyFiles()

    Dim WS As Worksheet, PR As Workbook, WK As Workbook, galley_files As Variant
    Set WK = ThisWorkbook

    'Get the files
    galley_files = Application.GetOpenFilename(Title:="Choose a file to add", MultiSelect:=False)

    If galley_files <> False Then
        
        'Files names
        Dim search_array_files_names As Range
        Set search_array_files_names = WK.Worksheets("Galley").Range("AV:AV")

        'Get file names
        Form.galleyTextBox.Value = extract_file_name(galley_files)

        'Open the files
        Workbooks.Open filename:=galley_files
        Set PR = Workbooks(Workbooks.count)

        'Validating I'ts the correct Galley
        If PR.Worksheets(1).Range("I2").Value2 = "Name" Then

            'Take the last row from the two files
            lastrow = WK.Worksheets("Galley").Range("E" & Rows.count).End(xlUp).Row
            lastrow2 = PR.Worksheets(1).Range("D" & Rows.count).End(xlUp).Row

            If lastrow > 1 Then
            
                'Match the name of the file before add it to the sheet
                matched_file = match_name_phone(Form.galleyTextBox.Value, search_array_files_names)
                
                If matched_file = "MATCHED" Then
                    MsgBox "This file was added previously:" & " " & Form.galleyTextBox.Value & vbCrLf & "Please try with a new file..."
                    PR.Close savechanges = False
                    Exit Sub
                End If
                
                'copy the others files to the working workbook
                PR.Worksheets(1).Range("A1:AP" & lastrow2).Copy Destination:=WK.Worksheets("Galley").Range("A" & lastrow + 1)
                MsgBox "File added : " & Form.galleyTextBox.Value
            Else
                'copy the first file to the working workbook
                PR.Worksheets(1).Range("A:AT").Copy Destination:=WK.Worksheets("Galley").Range("A:AT")
            End If
        Else
            MsgBox "You Choose a wrong Galley. Please try again:"
            PR.Close savechanges = False
            Exit Sub
        End If
        
        'Save file name
        lastrow = WK.Worksheets("Galley").Range("AV" & Rows.count).End(xlUp).Row
        WK.Worksheets("Galley").Range("AV" & last_row + 1).Value = Form.galleyTextBox.Value

        'Close the Galley
        PR.Close savechanges = False

        'Show image
        Form.Image2.Visible = True
    End If
    
End Sub
