Attribute VB_Name = "Module1"
Option Explicit


Sub Commitee_Applications_Reading_Lists()

Dim File_Name As String
Dim FSO
Dim File_Location As String
Dim Folder_Destination As String
Dim File_Destination As String
Dim Last_Name As String
Dim First_Name As String
Dim Reading_Range As Range
Dim Dept As String
Dim Unit As String
Dim Total_Score As String
Dim Reading_List As Worksheet
Dim Cell As Range
Dim Ranking As String
Dim Level_Study As String
Dim Reviewer_Department_Range As Range
Dim Workbook_Name As String
Dim Worksheet_Name As String
Dim Reading_List_Template As Range
Dim RowHght2 As Long
Dim RowHght3 As Long
Dim RowHght5 As Long
Dim Files_Location As String
Dim YearX As String
Dim YearC As String
Dim Workbook As Object


Workbook_Name = ActiveWorkbook.Name
Worksheet_Name = ActiveSheet.Name

'This is to make sure that the reading lists look legible when copied over for each reviewer

RowHght2 = Workbooks(Workbook_Name).Worksheets(Worksheet_Name).Range("A2").EntireRow.Height
RowHght3 = Workbooks(Workbook_Name).Worksheets(Worksheet_Name).Range("A3").EntireRow.Height
RowHght5 = Workbooks(Workbook_Name).Worksheets(Worksheet_Name).Range("A5").EntireRow.Height

YearC = Cells(1, 8).Value
YearX = Left(Cells(1, 8), 4)

File_Location = Application.InputBox("Enter the folder location of the Applications. Make sure this address includes a backward slash \ at the end.", _
Type:=2)

Files_Location = Application.InputBox("Enter the folder location of the Commitee Files (guidelines, score sheet, etc). Make sure this address includes a backward slash \ at the end.", _
Type:=2)

Workbooks(Workbook_Name).Worksheets(Worksheet_Name).Activate

'Files_Location is where the additional pdfs such as score sheet, normalization info, and commitee guidelines are located


Set Reading_List_Template = Workbooks(Workbook_Name).Worksheets(Worksheet_Name).Range("A1:H5")



Set Reading_Range = Application.InputBox("Select the range of the reading assignments for all members. These are all the 1's that you've entered. Do not include the header.", _
Type:=8)

ActiveWorkbook.Names.Add _
            Name:="Committee_Reading", _
            RefersTo:=Reading_Range
            
Set Reviewer_Department_Range = Application.InputBox("Select the range of the Committee Member Units.", _
Type:=8)

ActiveWorkbook.Names.Add _
            Name:="Reviewer_Unit", _
            RefersTo:=Reviewer_Department_Range


Workbooks(Workbook_Name).Worksheets(Worksheet_Name).Activate


Folder_Destination = Application.InputBox("Enter the folder location of where the review file folders should go. This is probably just your OneDrive folder. Make sure this address includes a backward slash \ at the end.", _
Type:=2)


Set FSO = CreateObject("Scripting.FileSystemObject")

'Creates reading list sheets for each reviewer

For Each Cell In Workbooks(Workbook_Name).Worksheets(Worksheet_Name).Range("Reviewer_Unit")

Sheets.Add(After:=Sheets(Sheets.Count)).Name = ("Reading_List - " & Cell.Value)
Reading_List_Template.Copy
With Workbooks(Workbook_Name).Worksheets("Reading_List - " & Cell.Value).Range("A1")
        .Cells(1).PasteSpecial xlPasteColumnWidths
        .Cells(1).PasteSpecial xlPasteValues
        .Cells(1).PasteSpecial xlPasteFormats
        
Workbooks(Workbook_Name).Worksheets("Reading_List - " & Cell.Value).Range("A2").RowHeight = RowHght2
Workbooks(Workbook_Name).Worksheets("Reading_List - " & Cell.Value).Range("A3").RowHeight = RowHght3
Workbooks(Workbook_Name).Worksheets("Reading_List - " & Cell.Value).Range("A5").RowHeight = RowHght5

End With
        
Cells(1, 8).Value = Cell.Value

'Creates a folder for each member in the specific Folder Destination (Onedrive)

MkDir (Folder_Destination & YearC & " CIHR CGS D Committee Files - " & Cell.Value)

Next Cell

Workbooks(Workbook_Name).Worksheets(Worksheet_Name).Activate

'Populates the reading lists and saves the applications to the reviewer's folder

For Each Cell In Workbooks(Workbook_Name).Worksheets(Worksheet_Name).Range("Committee_Reading")

If Cell.Value = "1" Then

Cell.Activate
Last_Name = Cells(ActiveCell.Row, 1).Value
First_Name = Cells(ActiveCell.Row, 2).Value
Dept = Cells(ActiveCell.Row, 3).Value
Level_Study = Cells(ActiveCell.Row, 4).Value
Unit = Cells(5, ActiveCell.Column).Value

Set Reading_List = Sheets("Reading_List - " & Unit)

File_Destination = (Folder_Destination & YearC & " CIHR CGS D Committee Files - " & Unit & "\")

FSO.CopyFile (File_Location & Last_Name & ", " & First_Name & ", CIHRDoc2021.pdf"), File_Destination, True

Reading_List.Range("A999").End(xlUp).Offset(1, 0).Value = Last_Name
Reading_List.Range("B999").End(xlUp).Offset(1, 0).Value = First_Name
Reading_List.Range("C999").End(xlUp).Offset(1, 0).Value = Dept
Reading_List.Range("D999").End(xlUp).Offset(1, 0).Value = Level_Study


Else
End If
Next Cell

Workbooks(Workbook_Name).Worksheets(Worksheet_Name).Activate

'Saves the Reading Lists and Additional PDFs to the reviewer's folder

For Each Cell In Workbooks(Workbook_Name).Worksheets(Worksheet_Name).Range("Reviewer_Unit")

Workbooks(Workbook_Name).Worksheets("Reading_List - " & Cell.Value).Activate

Workbooks(Workbook_Name).Worksheets("Reading_List - " & Cell.Value).Range(Range("A1").End(xlToRight), Range("A1").End(xlDown)).Borders.LineStyle = xlContinuous

Workbooks(Workbook_Name).Worksheets("Reading_List - " & Cell.Value).Copy Before:=Workbooks.Add.Sheets(1)

Application.ActiveWorkbook.SaveAs Filename:=Folder_Destination & YearC & " CIHR CGS D Committee Files - " & Cell.Value & "\" & "1. CIHR Doc Reading List - " & Cell.Value & ".xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
ActiveWorkbook.Close

FSO.CopyFile (Files_Location & "2. Score Sheet - CGS Doctoral Awards.pdf"), Folder_Destination & YearC & " CIHR CGS D Committee Files - " & Cell.Value & "\", True
FSO.CopyFile (Files_Location & "3. SGS Awards Committee Guidelines.pdf"), Folder_Destination & YearC & " CIHR CGS D Committee Files - " & Cell.Value & "\", True
FSO.CopyFile (Files_Location & "4. Normalisation for Awards Adjudication.pdf"), Folder_Destination & YearC & " CIHR CGS D Committee Files - " & Cell.Value & "\", True

Next Cell

Application.ActiveWorkbook.Names("Committee_Reading").Delete
Application.ActiveWorkbook.Names("Reviewer_Unit").Delete

Workbooks(Workbook_Name).Worksheets(Worksheet_Name).Activate

MsgBox ("Complete! Reading lists and applications have been successfully saved in " & Folder_Destination & ".")

End Sub



Sub Test_Application_Names()

Dim File_Name As String
Dim FSO
Dim File_Location As String
Dim File_Destination As String
Dim Last_Name As String
Dim First_Name As String
Dim Applicant_Range As Range
Dim Dept As String
Dim Unit As String
Dim Reading_List As Worksheet
Dim Cell As Range
Dim Ranking As String
Dim Level_Study As String
Dim Workbook_Name As String
Dim Worksheet_Name As String
Dim YearX As String
Dim YearC As String
Dim Count_Highlight As Integer

MsgBox ("This code checks that there is a PDF file in the format *Last Name, First Name - UNIT* in a specific folder. There will be some questions asked now. If you're unsure about what to enter, go bug Olivera.")
Workbook_Name = ActiveWorkbook.Name
Worksheet_Name = ActiveSheet.Name

Set Applicant_Range = Application.InputBox("Select the range of Last Names only. Do not include headers.", _
Type:=8)

ActiveWorkbook.Names.Add _
            Name:="Applicant", _
            RefersTo:=Applicant_Range

Set FSO = CreateObject("Scripting.FileSystemObject")

YearC = Cells(1, 8).Value
YearX = Left(Cells(1, 8), 4)


File_Location = Application.InputBox("Enter the folder location of the Applications. Make sure this address includes a backward slash \ at the end.", _
Type:=2)

Count_Highlight = 0

For Each Cell In Workbooks(Workbook_Name).Worksheets(Worksheet_Name).Range("Applicant")

Cell.Activate
Last_Name = ActiveCell.Value
First_Name = Cells(ActiveCell.Row, 2).Value
Dept = Cells(ActiveCell.Row, 3).Value
Level_Study = Cells(ActiveCell.Row, 4).Value


If Not FSO.FileExists(File_Location & Last_Name & ", " & First_Name & ", CIHRDoc2021.pdf") Then
Cell.Activate
ActiveCell.Interior.ColorIndex = 35
Count_Highlight = Count_Highlight + 1

Else

End If
Next Cell

Application.ActiveWorkbook.Names("Applicant").Delete

MsgBox ("There are a total of " & Count_Highlight & " application(s) with mismatching names in " & File_Location & ". See name(s) highlighted in green.")

End Sub


