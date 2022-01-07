Attribute VB_Name = "captions"
Sub Match_Captions()
    Dim RS As Workbook
    Set RS = ThisWorkbook

    'Variables declarations Vivial Force file
    Dim heading As String, name As String, phone As String, concatenated As String, lastrow As Long, listing_type As String, count As Long
    
    ' //Concatenated VF File//
    lastrow = RS.Worksheets("VFile").Range("D" & Rows.count).End(xlUp).Row

    'Concatenated caption lines in the Vivial Force file
    For i = 2 To lastrow

        'Only captions listings are allowed
        If RS.Worksheets("VFile").Range("T" & i).Value <> 0 Then
            
        End If
    Next i
    
End Sub
