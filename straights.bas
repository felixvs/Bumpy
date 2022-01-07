Attribute VB_Name = "straights"
Sub Match_Straights()
    Dim RS As Workbook
    Set RS = ThisWorkbook
    
    'Variables declarations Vivial Force file
    Dim heading As String, name As String, phone As String, concatenated As String, lastrow As Long, listing_type As String, count As Long
    
    ' //Concatenated VF File//
    lastrow = RS.Worksheets("VFile").Range("D" & Rows.count).End(xlUp).Row
    
    'Concatenated straight listings in the Vivial Force file
    For i = 2 To lastrow
    
        'Only straight listings are allowed
        If RS.Worksheets("VFile").Range("S" & i).Value = 0 And RS.Worksheets("VFile").Range("T" & i).Value = 0 Then
            
            N_section = LCase(RS.Worksheets("VFile").Range("N" & i).Value)
            AD_name = LCase(RS.Worksheets("VFile").Range("AD" & i).Value)
            Q_First_Name = LCase(RS.Worksheets("VFile").Range("Q" & i).Value)
            K_Designation = LCase(RS.Worksheets("VFile").Range("K" & i).Value)
            AC_Listing_Street_Number = LCase(RS.Worksheets("VFile").Range("AC" & i).Value)
            AB_Listing_Street = LCase(RS.Worksheets("VFile").Range("AB" & i).Value)
            
            'Run if the check box is selected
            If Form.CheckBox1 = True Then
                W_Listing_City = LCase(RS.Worksheets("VFile").Range("W" & i).Value)
            End If
            
            'Run if the check box is selected
            If Form.CheckBox2 = True Then
                Z_Postal_Code = LCase(RS.Worksheets("VFile").Range("Z" & i).Value)
            End If
            
            I_Cross_Reference = LCase(RS.Worksheets("VFile").Range("I" & i).Value)
            AE_Phone = Replace(Replace(LCase(RS.Worksheets("VFile").Range("AE" & i).Value), "-", ""), " ", "")
            AF_Phone_Override = LCase(RS.Worksheets("VFile").Range("AF" & i).Value)
            
            If AF_Phone_Override <> vbNullString Then
            'DC Added replace to get rid of the special characters from the phone override field 12.10.21
                AE_Phone = Replace(Replace(Replace(Replace(Replace(Replace(LCase(RS.Worksheets("VFile").Range("AF" & i).Value), "/", ""), "\", ""), ")", ""), "(", ""), "-", ""), " ", "")
            End If
            
            '7 digits to compare. Why? Because the galley print with 7 digits some times.
            AE_Phone = Right(AE_Phone, 7)
            
            'Concatenate names, designation, cross reference
            name = Trim(AD_name & Q_First_Name & K_Designation & I_Cross_Reference)

            'Concatenate AC_Listing_Street_Number and AB_Listing_Street
            street = Trim(AC_Listing_Street_Number & AB_Listing_Street)
            
            'Concatenate all data
            concatenate_all = Replace(Trim(N_section & "|" & name & "|" & street & "|" & W_Listing_City & "|" & Z_Postal_Code & "|" & AE_Phone), " ", "")
            'Concatenate data without the street
            concatenate_missing_street = Replace(Trim(N_section & "|" & name & "|" & W_Listing_City & "|" & Z_Postal_Code & "|" & AE_Phone), " ", "")
            'Concatenate data without the name
            concatenate_missing_name = Replace(Trim(N_section & "|" & street & "|" & W_Listing_City & "|" & Z_Postal_Code & "|" & AE_Phone), " ", "")
            'Concatenate data without the listing city
            concatenate_missing_Listing_City = Replace(Trim(N_section & "|" & name & "|" & street & "|" & Z_Postal_Code & "|" & AE_Phone), " ", "")
            'Concatenate data without the phone
            concatenate_missing_phone = Replace(Trim(N_section & "|" & name & "|" & street & "|" & W_Listing_City & "|" & Z_Postal_Code), " ", "")

            'Writing results to the VF sheet
            RS.Worksheets("VFile").Range("AQ" & i).Value = concatenate_all
            RS.Worksheets("VFile").Range("AR" & i).Value = concatenate_missing_street
            RS.Worksheets("VFile").Range("AS" & i).Value = concatenate_missing_name
            RS.Worksheets("VFile").Range("AT" & i).Value = concatenate_missing_Listing_City
            RS.Worksheets("VFile").Range("AU" & i).Value = concatenate_missing_phone
        End If
    Next i
    
    ' //Concatenated Galley//
    lastrow = RS.Worksheets("Galley").Range("E" & Rows.count).End(xlUp).Row
    
    'Concatenated straight listings in the Galley
    For i = 4 To lastrow
        
        'Only straight listings are allowed
        If RS.Worksheets("Galley").Range("E" & i).Value = 0 And RS.Worksheets("Galley").Range("I" & i).Value <> vbNullString And _
           (Len(RS.Worksheets("Galley").Range("P" & i).Value) <> 0 Or RS.Worksheets("Galley").Range("R" & i).Value <> vbNullString) Then
            
            S_section = LCase(RS.Worksheets("Galley").Range("S" & i).Value)
            I_Name = LCase(RS.Worksheets("Galley").Range("I" & i).Value)
            K_Street = LCase(RS.Worksheets("Galley").Range("K" & i).Value)
            
            'Run if the check box is selected
            If Form.CheckBox1 = True Then
                M_City = LCase(RS.Worksheets("Galley").Range("M" & i).Value)
            End If
            
            'Run if the check box is selected
            If Form.CheckBox2 = True Then
                N_Postal_Code = LCase(RS.Worksheets("Galley").Range("N" & i).Value)
            End If
            
            R_Cross_Reference = LCase(RS.Worksheets("Galley").Range("R" & i).Value)
            P_Phone = Replace(Replace(Replace(Replace(LCase(RS.Worksheets("Galley").Range("P" & i).Value), "-", ""), " ", ""), "/", ""), "\", "")
            
            '7 digits to compare. Why? Because the galley print with 7 digits some times.
            P_Phone = Right(P_Phone, 7)
            
            'Concatenate names, cross reference
            name = Trim(I_Name & R_Cross_Reference)
            
            ''Concatenate all data
            concatenate_all = Replace(Trim(S_section & "|" & name & "|" & K_Street & "|" & M_City & "|" & N_Postal_Code & "|" & P_Phone), " ", "")
            'Concatenate data without the street
            concatenate_missing_street = Replace(Trim(S_section & "|" & name & "|" & M_City & "|" & N_Postal_Code & "|" & P_Phone), " ", "")
            'Concatenate data without the name
            concatenate_missing_name = Replace(Trim(S_section & "|" & K_Street & "|" & M_City & "|" & N_Postal_Code & "|" & P_Phone), " ", "")
            'Concatenate data without the listing city
            concatenate_missing_Listing_City = Replace(Trim(S_section & "|" & name & "|" & K_Street & "|" & N_Postal_Code & "|" & P_Phone), " ", "")
            'Concatenate data without the phone
            concatenate_missing_phone = Replace(Trim(S_section & "|" & name & "|" & K_Street & "|" & M_City & "|" & N_Postal_Code), " ", "")
            
            'Writing results to the galley sheet
            RS.Worksheets("Galley").Range("AQ" & i).Value = concatenate_all
            RS.Worksheets("Galley").Range("AR" & i).Value = concatenate_missing_street
            RS.Worksheets("Galley").Range("AS" & i).Value = concatenate_missing_name
            RS.Worksheets("Galley").Range("AT" & i).Value = concatenate_missing_Listing_City
            RS.Worksheets("Galley").Range("AU" & i).Value = concatenate_missing_phone
        End If
    Next i
    
    ' //Match VF File//
    'Columns headers
    RS.Worksheets("Straights").Range("A1").Value = "VF Straight Listings"
    RS.Worksheets("Straights").Range("B1").Value = "Mismatch Type"
    
    'Range variables
    Dim search_array_galley_all As Range, search_value_vivialforce As Variant, pipev As Variant
    Dim search_array_galley_street As Range, search_array_galley_name As Range, search_array_galley_listing_city As Range, search_array_galley_phone As Range
    
    'Set the galley concatenations into an array
    Set search_array_galley_all = RS.Worksheets("Galley").Range("AQ:AQ")
    Set search_array_galley_street = RS.Worksheets("Galley").Range("AR:AR")
    Set search_array_galley_name = RS.Worksheets("Galley").Range("AS:AS")
    Set search_array_galley_listing_city = RS.Worksheets("Galley").Range("AT:AT")
    Set search_array_galley_phone = RS.Worksheets("Galley").Range("AU:AU")
    
    For i = 2 To lastrow
    
        'Setting VF concatenations to seach into the galley array
        search_value_vivialforce = RS.Worksheets("VFile").Range("AQ" & i).Value
        search_value_vivialforce_street = RS.Worksheets("VFile").Range("AR" & i).Value
        search_value_vivialforce_name = RS.Worksheets("VFile").Range("AS" & i).Value
        search_value_vivialforce_listing_city = RS.Worksheets("VFile").Range("AT" & i).Value
        search_value_vivialforce_phone = RS.Worksheets("VFile").Range("AU" & i).Value
        
        If Len(search_value_vivialforce) <> 0 Then
            
            'Calling function match_heading_name_phone() - listing_type is opcional
            matched = match_name_phone(search_value_vivialforce, search_array_galley_all)

            If matched = "NOT MATCHED" Then
                last_row = RS.Worksheets("Straights").Range("A" & Rows.count).End(xlUp).Row
                
                N_section = RS.Worksheets("VFile").Range("N" & i).Value
                AD_name = RS.Worksheets("VFile").Range("AD" & i).Value
                Q_First_Name = RS.Worksheets("VFile").Range("Q" & i).Value
                K_Designation = RS.Worksheets("VFile").Range("K" & i).Value
                AC_Listing_Street_Number = RS.Worksheets("VFile").Range("AC" & i).Value
                AB_Listing_Street = RS.Worksheets("VFile").Range("AB" & i).Value
                W_Listing_City = RS.Worksheets("VFile").Range("W" & i).Value
                Z_Postal_Code = LCase(RS.Worksheets("VFile").Range("Z" & i).Value)
                I_Cross_Reference = RS.Worksheets("VFile").Range("I" & i).Value
                AE_Phone = RS.Worksheets("VFile").Range("AE" & i).Value
                AF_Phone_Override = RS.Worksheets("VFile").Range("AF" & i).Value
                
                If AF_Phone_Override <> vbNullString Then
                    AE_Phone = RS.Worksheets("VFile").Range("AF" & i).Value
                End If
    
                'Concatenate names and designation
                name = Trim(AD_name & " " & Q_First_Name & " " & K_Designation & " " & I_Cross_Reference)
    
                'Concatenate AC_Listing_Street_Number and AB_Listing_Street
                street = Trim(AC_Listing_Street_Number & " " & AB_Listing_Street)
                
                'Formating the phone number
                'AE_Phone = Format(AE_Phone, "000 000-0000")
                
                'Concatenate the rest of the data
                concatenate_columns = Replace(Trim(N_section & " | " & name & " | " & street & " | " & W_Listing_City & " | " & Z_Postal_Code & " | " & AE_Phone), " |  | ", " | ")
                
                'Writing results
                RS.Worksheets("Straights").Range("A" & last_row + 1).Value = concatenate_columns
                
                'Matching all the concatenation
                matched_street = match_name_phone(search_value_vivialforce_street, search_array_galley_street)
                matched_name = match_name_phone(search_value_vivialforce_name, search_array_galley_name)
                matched_listing_city = match_name_phone(search_value_vivialforce_listing_city, search_array_galley_listing_city)
                matched_phone = match_name_phone(search_value_vivialforce_phone, search_array_galley_phone)
                
                'If matched, write the missing part next to the listing
                If matched_street = "MATCHED" Then
                    RS.Worksheets("Straights").Range("B" & last_row + 1).Value = "Address"
                ElseIf matched_name = "MATCHED" Then
                    RS.Worksheets("Straights").Range("B" & last_row + 1).Value = "Name"
                ElseIf matched_listing_city = "MATCHED" Then
                    RS.Worksheets("Straights").Range("B" & last_row + 1).Value = "Community"
                ElseIf matched_phone = "MATCHED" Then
                    RS.Worksheets("Straights").Range("B" & last_row + 1).Value = "Phone"
                Else
                    RS.Worksheets("Straights").Range("B" & last_row + 1).Value = "Insert"
                End If
                
            End If
        End If
    Next i
            
    'Formating columns names
    With RS.Worksheets("Straights").Range("A1:B1")
        .Font.Bold = True
        .EntireColumn.AutoFit
    End With
    
    'Remove helpers columns
    RS.Worksheets("VFile").Range("AQ:AU").Delete
    RS.Worksheets("Galley").Range("AQ:AU").Delete
    
    lasts_row = RS.Worksheets("Straights").Range("A" & Rows.count).End(xlUp).Row
    
    'If nothing found close the macro
    If lasts_row = 1 Then
        MsgBox "Perfect match between straight listings."
        ThisWorkbook.Close savechanges = False
    End If
    
    'AutoFit Columns
    'WK.Worksheets("Straights").Columns("A:B").AutoFit
       
    
'    ' //Match from galley to vivial force//
'    'Range variables
'    Dim search_array_VFile As Range, search_value_galley As Variant
'    Set search_array_VFile = RS.Worksheets("VFile").Range("AQ:AQ")
'
'    For i = 2 To lastrow
'        '//Yello Pages Galley//
'        search_value_galley = RS.Worksheets("Galley").Range("AQ" & i).Value
'
'        'Validating if the cell has a heading
'        If Len(search_value_galley) <> 0 Then
'
'            'Calling function match_heading_name_phone() - listing_type is opcional
'            matched = match_name_phone(search_value_galley, search_array_VFile)
'
'            If matched = "NOT MATCHED" Then
'                last_row = RS.Worksheets("Straights").Range("B" & Rows.count).End(xlUp).Row
'
'                I_Name = RS.Worksheets("Galley").Range("I" & i).Value
'                K_Street = RS.Worksheets("Galley").Range("K" & i).Value
'                M_City = RS.Worksheets("Galley").Range("M" & i).Value
'                P_Phone = Replace(Replace(LCase(RS.Worksheets("Galley").Range("P" & i).Value), "-", ""), " ", "")
'
'                'Concatenate the rest of the data
'                concatenate_columns = Trim(I_Name & " | " & K_Street & " | " & M_City & " | " & P_Phone)
'
'                'Writing results
'                RS.Worksheets("Straights").Range("B" & last_row + 1).Value = concatenate_columns
'            End If
'        End If
'
'
'    Next i

End Sub
