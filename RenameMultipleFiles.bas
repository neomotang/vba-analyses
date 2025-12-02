Attribute VB_Name = "RenameFiles"

Sub RenameMultipleFiles()

'compiled November 2020

'This sub renames multiple files in a selected directory. The original file names are entered in column B, and the new file names are in column D.

    With Application.FileDialog(msoFileDialogFolderPicker) 'prompt to select a directory/folder
        .AllowMultiSelect = False
        If .Show = -1 Then
            selectDirectory = .SelectedItems(1)
            dFileList = Dir(selectDirectory & Application.PathSeparator & "*")'retrieves the names of all files in the selected folder
        
            Do Until dFileList = ""
                curRow = 0
                On Error Resume Next
                curRow = Application.Match(dFileList, Range("B:B"), 0) 'for an individual file, identifies the row entry in the spreadsheet containing its current filename
                If curRow > 0 Then'>0 if the current filename does appear in spreadsheet; changes the filename to the entry in column D on the same row
                    Name selectDirectory & Application.PathSeparator & dFileList As _
                    selectDirectory & Application.PathSeparator & Cells(curRow, "D").Value
                End If
        
                dFileList = Dir 'moves on the next file on the list, removing the filename that has just been processed
            Loop
        End If
    End With
    
End Sub
