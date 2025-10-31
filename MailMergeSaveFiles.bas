Attribute VB_Name = "MailMergeSaveFiles"

Option Explicit

Sub LoopThroughData()

'compiled October 2020

'This sub is a "mail merge" example that converts related data entries in multiple columns for the same identifier into an individual pdf with the data arranged in a desired format, e.g. quiz answers from a student into a more reader friendly format.
'bulk formula changes to cells in a worksheet tab containing existing data in a known format, e.g. if calibration constants or property formulas must be updated. The constants are stored in a separate tab in the file ("Constants").
'In the example below, the experimental data ("RootNum") are used to calculate various properties such as their square, square root, cube, etc.

'The code belows is for example data where there are five pieces of data per row. The working spreadsheet has the following tabs:
'  "Master" - contains 3 buttons; a button executes "ChooseFolderPath" (output as file path in cell F8), a second button executes "LoopThroughData" (output as number of successful executions in F15) and a third button executes "ClearData" (see ClearResults.bas) to clear previous outputs
'xxx'of data: first row is blank, cells B2:G2 have headings (Number, Square root, Straight line, Square, Cube, Cubic root), data block 1 in B3:G7, data block 2 in B10:G14 and data block 3 in B17:G21; the integers 1 - 15 are entered in B3:B7, B10:B14 and B17:B21
'  Sheet4 ("Constants") - contains constants to be used in the formulas, in cells B2:C6; headings in column B (Square root, Line, Square, Cube, Cubic root) and values in column C (0.5, 1, 2, 3, 1/3)


Dim saveLocation, PDFFileName As String
Dim SuccessCount, RowCount, rCounter As Integer
Dim DataCheck As Range

'***Edit/update this block***
'The following lists the variables that will hold data entered by students, e.g. StudentNum, Surname, Name, Q1, Q2, etc.
Dim DataID, Position, Alphabet, Greek, RandomSent As String

Application.ScreenUpdating = False 'Stops the screen from flashing while the files are being processed
saveLocation = Sheets("Master").Range("F8").Value 'Reads off path of output folder as previously chosen
SuccessCount = 0 'counts the number of PDFs successfully saved

'Verify that an output folder has been selected
If saveLocation = "" Then
    MsgBox "Please choose folder first!", vbCritical
    Exit Sub
End If

Sheets("Data").Activate 'Go to tab with data from SUNLearn
Range("A1").Select 'Find first cell on tab

'Verify that the tab is not empty
Set DataCheck = ActiveCell
If DataCheck Is Nothing Then
    MsgBox "Student data sheet is empty!", vbCritical
    Exit Sub
End If

RowCount = Sheets("Data").Range("A1", Range("A1").End(xlDown)).Rows.Count 'Count number of rows with entries in column A
rCounter = 2 'Indicates row currently being processed; start processing in the second row, assuming the first row contains headings

Do While rCounter <= RowCount
    Application.ScreenUpdating = False 'Stops the screen from flashing while the files are being processed
    Application.DisplayAlerts = False 'Disables alerts while the files are being processed
    
'***Edit/update this block***
'The following should reflect the variables as listed earlier, e.g. StudentNum, Surname, Name, Q1, Q2, etc.
    'Read off the data in the row into appropriate variables
    DataID = Sheets("Data").Cells(rCounter, 1).Value
    Position = Sheets("Data").Cells(rCounter, 2).Value
    Alphabet = Sheets("Data").Cells(rCounter, 3).Value
    Greek = Sheets("Data").Cells(rCounter, 4).Value
    RandomSent = Sheets("Data").Cells(rCounter, 5).Value
    
'***Edit/update this block***
'The following should reflect the variables as listed earlier, e.g. StudentNum, Surname, Name, Q1, Q2, etc.
    'Read off the data from variables into appropriate cells on the "Template" tab
    Sheets("Template").Activate
    Range("C1").Value = DataID
    Range("C2").Value = Position
    Range("B4").Value = Alphabet
    Range("B5").Value = Greek
    Range("B6").Value = RandomSent
     
    'Save the PDF
    Rows.VerticalAlignment = xlVAlignTop
    Rows.WrapText = True
    
'***Edit/update this block***
'The following should reflect the filenames to use when saving each PDF, e.g. PDFFileName = Surname & " " & StudentNum
    PDFFileName = Position & " " & DataID 'Determine the filename of the PDF corresponding to the current data
    
    Sheets("Template").Activate
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=saveLocation & PDFFileName, Quality:=xlQualityStandard
    SuccessCount = SuccessCount + 1
    rCounter = rCounter + 1 'Move to the next row of student data
    Application.ScreenUpdating = True
Loop

Sheets("Master").Range("B15").Value = "Number of files saved successfully:"
Sheets("Master").Range("F15").Value = SuccessCount
Sheets("Master").Activate
Application.ScreenUpdating = True
    
End Sub
