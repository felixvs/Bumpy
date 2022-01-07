VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form 
   Caption         =   "Bumpy"
   ClientHeight    =   1845
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4065
   OleObjectBlob   =   "Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Button_Click()
    Dim WS As Worksheet, PR As Workbook, WK As Workbook
    Set WK = ThisWorkbook

    'White out and clear the screen
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Image2.Visible = True
    
    'Validating that the Text Boxes are empty.
    If Image1.Visible = False Then
        MsgBox "Please provided the VF File"
        Exit Sub
    ElseIf Image2.Visible = False Then
        MsgBox "Please provided the Galley File."
        Exit Sub
    End If

    'Validating if the files are open.
    If IsFileOpen(VFTextBox.Value) Then
        MsgBox "VF File is Open, Please close the file and try again"
        Exit Sub
    ElseIf IsFileOpen(galleyTextBox.Value) Then
        MsgBox "Galley is Open, Please close the file and try again"
        Exit Sub
    End If
    

'    'Validating that the Text Boxes are empty.
'    If VFTextBox.Value = vbNullString Then
'        MsgBox "Please provided the VF File"
'        Exit Sub
'    ElseIf galleyTextBox.Value = vbNullString Then
'        MsgBox "Please provided the Galley File."
'        Exit Sub
'    End If
'
'    'Validating if the files are open.
'    If IsFileOpen(VFTextBox.Value) Then
'        MsgBox "VF File is Open, Please close the file and try again"
'        Exit Sub
'    ElseIf IsFileOpen(galleyTextBox.Value) Then
'        MsgBox "Galley is Open, Please close the file and try again"
'        Exit Sub
'    End If

'    'Open the VF file
'    Workbooks.Open filename:=VFTextBox.Value
'    Set PR = Workbooks(Workbooks.count)
'
'    'Validating I'ts the correct VF file
'    If PR.Worksheets(1).Range("AD1").Value2 = "Name" Then
'        'copy it to the working workbook
'        PR.Worksheets(1).Range("A:AP").Copy Destination:=WK.Worksheets("VFile").Range("A:AP")
'    Else
'        MsgBox "You Choose a wrong VF file. Please try again:"
'        PR.Close savechanges = False
'        Exit Sub
'    End If
'
'    'Close the VF File
'    PR.Close savechanges = False
    
'    'Open the Galley
'    Workbooks.Open filename:=galleyTextBox.Value
'    Set PR = Workbooks(Workbooks.count)
'
'    'Validating I'ts the correct Galley file
'    If PR.Worksheets(1).Range("I2").Value2 = "Name" Then
'        'copy it to the working workbook
'        PR.Worksheets(1).Range("A:AT").Copy Destination:=WK.Worksheets("Galley").Range("A:AT")
'    Else
'        MsgBox "You Choose a wrong Galley. Please try again:"
'        PR.Close savechanges = False
'        Exit Sub
'    End If
'
'    'Close the Galley
'    PR.Close savechanges = False
    
    'Hide the form
    Form.Hide
    
    'Calls Here
    Call Match_Straights

    'Get directory name
    Dim directory_name As String
    If InStr(VFTextBox.Value, "_") Then
        position1 = InStr(1, VFTextBox.Value, "_")
        position2 = InStr(position1 + 1, VFTextBox.Value, "_")
        position3 = position2 - position1
        position3 = Abs(position3)
        directory_name = Mid(VFTextBox.Value, position1 + 1, position3 - 1)
    End If
    
    'Save file in the same directory of the macro was placed.
    WK.SaveAs ThisWorkbook.Path & "\" & directory_name & " - Straight Report", FileFormat:=51
    
    'Macro finish massage.
    MsgBox "The Report has been created." & vbCr & "It's in the folder where the macro was placed."
    
    'Destroy form
    Unload Me
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    
    'Close excel, without saving.
    ThisWorkbook.Close savechanges = False
    
End Sub

''Get VF file
'Private Sub VFButton_Click()
'    Dim VF_file As Variant
'
'    VF_file = Application.GetOpenFilename(Title:="Choose Neustar File to add Section Code", MultiSelect:=False)
'    If VF_file <> False Then VFTextBox.Value = extract_file_name(VF_file)
'
'    'Show image
'    If VF_file <> False Then
'        Image1.Visible = True
'    End If
'
'End Sub

'Get Galley file
'Private Sub GalleyButton_Click()
'    Dim galley_file As Variant
'
'    galley_file = Application.GetOpenFilename(Title:="Choose Neustar File to add Section Code", MultiSelect:=False)
'    If galley_file <> False Then galleyTextBox.Value = extract_file_name(galley_file)
'
'    'Show image
'    If galley_file <> False Then
'        Image2.Visible = True
'    End If
'End Sub


'Get VF file
Private Sub VFButton_Click()
    Call VFiles
End Sub

'Get Galley file
Private Sub GalleyButton_Click()
    Call GalleyFiles
End Sub

''Get VF file
'Private Sub VFButton_Click()
'    Dim WS As Worksheet, PR As Workbook, WK As Workbook, files As Variant
'    Set WK = ThisWorkbook
'
'    'Get the files
'    file = Application.GetOpenFilename(Title:="Choose a file to add", MultiSelect:=False)
'
'    If file <> False Then
'
'        'Get file names
'        VFTextBox.Value = extract_file_name(file)
'
'        'Open the files
'        Workbooks.Open filename:=file
'        Set PR = Workbooks(Workbooks.count)
'
'        'Validating I'ts the correct VF file
'        If PR.Worksheets(1).Range("AD1").Value2 = "Name" Then
'
'            'Take the last row from the two files
'            lastrow = WK.Worksheets("VFile").Range("D" & Rows.count).End(xlUp).Row
'            lastrow2 = PR.Worksheets(1).Range("D" & Rows.count).End(xlUp).Row
'
'            If lastrow > 1 Then
'                'copy the others files to the working workbook
'                PR.Worksheets(1).Range("A1:AP" & lastrow2).Copy Destination:=WK.Worksheets("VFile").Range("A" & lastrow + 1)
'                MsgBox "File added : " & VFTextBox.Value
'            Else
'                'copy the first file to the working workbook
'                PR.Worksheets(1).Range("A:AP").Copy Destination:=WK.Worksheets("VFile").Range("A:AP")
'            End If
'        Else
'            MsgBox "You Choose a wrong VF file. Please try again:"
'            PR.Close savechanges = False
'            Exit Sub
'        End If
'
'        'Close the VF File
'        PR.Close savechanges = False
'
'        'Show image
'        Image1.Visible = True
'    End If
'
'End Sub


''Get Galley file
'Private Sub GalleyButton_Click()
'
'    Dim WS As Worksheet, PR As Workbook, WK As Workbook, galley_files As Variant
'    Set WK = ThisWorkbook
'
'    'Get the files
'    galley_files = Application.GetOpenFilename(Title:="Choose a file to add", MultiSelect:=False)
'
'    If galley_files <> False Then
'
'        'Get file names
'        galleyTextBox.Value = extract_file_name(galley_files)
'
'        'Open the files
'        Workbooks.Open filename:=galley_files
'        Set PR = Workbooks(Workbooks.count)
'
'        'Validating I'ts the correct Galley
'        If PR.Worksheets(1).Range("I2").Value2 = "Name" Then
'
'            'Take the last row from the two files
'            lastrow = WK.Worksheets("Galley").Range("E" & Rows.count).End(xlUp).Row
'            lastrow2 = PR.Worksheets(1).Range("D" & Rows.count).End(xlUp).Row
'
'            If lastrow > 1 Then
'                'copy the others files to the working workbook
'                PR.Worksheets(1).Range("A1:AP" & lastrow2).Copy Destination:=WK.Worksheets("Galley").Range("A" & lastrow + 1)
'                MsgBox "File added : " & galleyTextBox.Value
'            Else
'                'copy the first file to the working workbook
'                PR.Worksheets(1).Range("A:AT").Copy Destination:=WK.Worksheets("Galley").Range("A:AT")
'            End If
'        Else
'            MsgBox "You Choose a wrong Galley. Please try again:"
'            PR.Close savechanges = False
'            Exit Sub
'        End If
'
'        'Close the Galley
'        PR.Close savechanges = False
'
'        'Show image
'        Image2.Visible = True
'    End If
'
'End Sub

'Close the form
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ThisWorkbook.Close savechanges = True
End Sub

Private Sub CheckBox1_Click()
 If CheckBox1.Value = False Then
    CheckBox1.Caption = "Community OFF"
 ElseIf CheckBox1.Value = True Then
    CheckBox1.Caption = "Community ON"
 End If
End Sub

Private Sub CheckBox2_Click()
 If CheckBox2.Value = False Then
    CheckBox2.Caption = "Zip Code OFF"
 ElseIf CheckBox1.Value = True Then
    CheckBox2.Caption = "Zip Code ON"
 End If
End Sub
