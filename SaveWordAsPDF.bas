Attribute VB_Name = "SaveWordAsPDF"

Sub ConvertMultipleWordToPDF()

'This sub converts MS Word files (.docx) in the input folder specified (in "OpenSourceFolder"/"SelectedWordFilesFolder") to PDF, saving the PDF files in the output folder specified (in "OpenTargetFolder"/"SelectedPdfFilesFolder")

'The .xlsm file contains only a single tab with a button to execute sub (placed in cells B2:D3); output from the sub is recorded as stats in cells B5:B7

Dim OpenSourceFolder As Object, OpenTargetFolder As Object
Dim InputWordFile, OutputPdfFile, SelectedPdfFilesFolder, SelectedWordFilesFolder As String
Dim objWordApp As Word.Application
Dim objMyWordFile As Word.Document
Dim SuccessCount As Integer

SuccessCount = 0
Range("B5").Value = ""
Range("B6").Value = ""
Range("B7").Value = ""

Application.ScreenUpdating = False 'Stops the screen from flashing while the files are being processed
Application.DisplayAlerts = False

Set objWordApp = CreateObject("Word.Application") 'opens MS Word program

'Select input file folder; stops executing if the folder is not selected
MsgBox ("Select input folder where word files are stored")
Set OpenSourceFolder = Application.FileDialog(msoFileDialogFolderPicker)
If OpenSourceFolder.Show = -1 Then
	SelectedWordFilesFolder = OpenSourceFolder.SelectedItems(1)
End If
If SelectedWordFilesFolder = "" Then
	MsgBox "No input folder selected. Code will exit"
	Exit Sub
End If

AppActivate Application.Caption

'Select output file folder; stops executing if the folder is not selected
MsgBox ("Select output folder where PDF files are stored")
Set OpenTargetFolder = Application.FileDialog(msoFileDialogFolderPicker)
If OpenTargetFolder.Show = -1 Then
	SelectedPdfFilesFolder = OpenTargetFolder.SelectedItems(1)
End If
If SelectedPdfFilesFolder = "" Then
	MsgBox "No output folder selected, code will exit"
	Exit Sub
End If

'For the selected input file folder, works through individual files in alphabetical order, opening only the MS Word (.docx) files and saving them as PDF in the selected output file folder
InputWordFile = Dir(SelectedWordFilesFolder & "\*.docx")
While InputWordFile <> ""  'Loop will end if either the input folder is empty or if the last Word file in the folder has been processed
	Set objMyWordFile = objWordApp.Documents.Open(SelectedWordFilesFolder & "\" & InputWordFile)
	objWordApp.Visible = True
	OutputPdfFile = SelectedPdfFilesFolder & "\" & Replace(objMyWordFile.Name, "docx", "pdf")
	objWordApp.ActiveDocument.ExportAsFixedFormat OutputFileName:=OutputPdfFile, ExportFormat:=wdExportFormatPDF 'saves PDF file with the same name as the Word file
	objMyWordFile.Close
	InputWordFile = Dir 'move on to next file in the folder
	SuccessCount = SuccessCount + 1
Wend

objWordApp.Documents.Application.Quit 'closes MS Word program

'Reports the stats on the main Excel tab once the code has been executed
Range("B5").Value = "Number of files successfully converted and saved: " & SuccessCount
Range("B6").Value = "Input folder (Word): " & SelectedWordFilesFolder
Range("B7").Value = "Output folder (PDF): " & SelectedPdfFilesFolder

End Sub
