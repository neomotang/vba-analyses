Attribute VB_Name = "ChooseFolder"
'Option Explicit

Sub ChooseFolderPath()
    
Dim fldr As FileDialog
Dim strPath As String
Dim FolderPath As String

Set fldr = Application.FileDialog(msoFileDialogFolderPicker) 'choose folder where student marksheet files are saved
fldr.Title = "Select the folder containing the marksheets"
fldr.AllowMultiSelect = False
fldr.InitialFileName = strPath

If fldr.Show <> -1 Then
    MsgBox "Folder was not chosen. Please try again.", vbCritical
    Else:
        FolderPath = fldr.SelectedItems(1) 'retrieves path name
        Sheets("Main").Range("B8").Value = "Folder path chosen:"
        Sheets("Main").Range("F8").Value = FolderPath & "\"
    End If

Sheets("Main").Activate

End Sub

