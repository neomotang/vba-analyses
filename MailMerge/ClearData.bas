Attribute VB_Name = "ClearResults"

Option Explicit

Sub ClearData()
    
'compiled October 2020

'This sub clears the temporary data in cells F8 & F15 on the "Master" tab, as well as all data on the "Template" tab.
'To use, update the "Template" tab, and then edit the section of code below indicated by: ***Edit/update from this point onwards***

Sheets("Master").Range("B8").Value = ""
Sheets("Master").Range("F8").Value = ""
Sheets("Master").Range("B15").Value = ""
Sheets("Master").Range("F15").Value = ""

'***Edit/update from this point onwards***
Sheets("Template").Range("C1").Value = ""
Sheets("Template").Range("C2").Value = ""
Sheets("Template").Range("B4").Value = ""
Sheets("Template").Range("B5").Value = ""
Sheets("Template").Range("B6").Value = ""
    
End Sub
