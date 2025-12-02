Attribute VB_Name = "CreateFolders"

Sub MakeFolders()

'compiled November 2020

'This sub creates multiple subfolders in the same folder as the current file, and gives them names as specified in column F.

Dim Rng As Range
Dim maxRows, maxCols, r, c As Integer

Set Rng = Sheet1.Range("F1:F90") 'names of the new subfolders are retrieved
'The next two lines of code are redundant for the specific example, but allow expansion to an unknown amount of subfolders if e.g. the line above is: Set Rng = Sheet1.Range("F1", Range("I1").End(xlDown))
maxRows = Rng.Rows.Count 
maxCols = Rng.Columns.Count
For c = 1 To maxCols
    r = 1
    Do While r <= maxRows
        If Len(Dir(ActiveWorkbook.Path & "\" & Rng(r, c), vbDirectory)) = 0 Then 'confirms that a subfolder with the potential new name does not already exist
            MkDir (ActiveWorkbook.Path & "\" & Rng(r, c))
            On Error Resume Next 'if the subfolder already exists, carry on and try to create the next subfolder
        End If
        
        r = r + 1
    Loop
Next c
    
End Sub
