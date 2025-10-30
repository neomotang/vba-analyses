Attribute VB_Name = UpdateFormulas"
'Option Explicit

Sub OneSheet()

'compiled August 2020

'This sub applies bulk formula changes to cells in a worksheet tab containing existing data in a known format, e.g. if calibration constants or property formulas must be updated. The constants are stored in a separate tab in the file ("Constants").
'In the example below, the experimental data ("RootNum") are used to calculate various properties such as their square, square root, cube, etc.

'The code belows is for example data where the spreadsheet has the following tabs:
'  Sheet1 ("Singles") - contains 3 blocks of data: first row is blank, cells B2:G2 have headings (Number, Square root, Straight line, Square, Cube, Cubic root), data block 1 in B3:G7, data block 2 in B10:G14 and data block 3 in B17:G21; the integers 1 - 15 are entered in B3:B7, B10:B14 and B17:B21
'  Sheet4 ("Constants") - contains constants to be used in the formulas, in cells B2:C6; headings in column B (Square root, Line, Square, Cube, Cubic root) and values in column C (0.5, 1, 2, 3, 1/3)

    Sheet1.Range("B1").End(xlDown).Offset(1, 0).Select 'find the beginning of the first data block, assuming the first row is blank (how I usually keep them) and the next row has headings
    ff = ActiveCell.Value 'intialize flag that should be raised if the end of the data block is reached, signified by an empty cell
    blockcounter = 0 'initialize counter, for me to see how many data blocks the loop goes through
    
    Do While Not IsEmpty(ff)
        RowID = ActiveCell.Row 'gets the row ID of the beginning of the block
        ColID = ActiveCell.Column 'gets the column ID of the beginning of the block
        counter = 0 'initialize counter, for me to see how many times the loop runs
        RootNum = ActiveCell.Value 'reads the number from the "Number" column, B
     
        Do While Not IsEmpty(RootNum) 'loop condition to check if the cell in column B is empty or not; run only if not empty
            ConstID = 2 'initialize row number of power in the "Constants" sheet
    
            Do While ColID <= 6 'because I know how many columns there are in the tab Sheet1
                Sheet1.Cells(RowID, ColID + 1).Value = "=" & "B" & RowID & "^Constants!$C$" & ConstID 'enters the formula of the type "=B3^Constants!$C$2" for RowID = 3 and ConstID = 2
                ColID = ColID + 1
                ConstID = ConstID + 1
                Loop 'the loop enters the formulas in the rest of the columns for the current row in the data block
            
            ColID = ActiveCell.Column 'intialize counter for the next loop iteration, since the selected cell is still the beginning of the current data block
            ConstID = 2 'intialize counter for the next loop iteration
            
            RowID = RowID + 1 'will move to the next row to check if empty in the next line and loop condition
            RootNum = Sheet1.Cells(RowID, ColID).Value
            counter = counter + 1 'just a normal loop counter
            Loop
    
        ActiveCell.End(xlDown).Offset(0, 0).Select 'find the end of the current block of numbers
        ActiveCell.End(xlDown).Offset(0, 0).Select 'find the beginning of the next block of numbers (assuming no headings repeated) or end of column if all number blocks have been identified
        'NB: ActiveCell.End(xlDown).Offset(0, 0).Select is executed twice because I usually leave two empty rows between data blocks, to allow for calculation of averages of each column
        
        ff = ActiveCell.Value 'if empty will be the end of the data on the tab
        blockcounter = blockcounter + 1
        Loop
    
    Sheet1.Range("B1").End(xlDown).Offset(1, 0).Select 'goes back to the beginning of the first data block before execution of code is complete
    
End Sub

Sub MultipleSheets()

'compiled August 2020

'This sub applies bulk formula changes to cells across multiple tabs in a worksheet containing existing data in a known format, e.g. if calibration constants or property formulas must be updated. The constants are stored in a separate tab in the file ("Constants").
'In the example below, the experimental data ("RootNum") are used to calculate various properties such as their square, square root, cube, etc.

'The code belows is for example data where the spreadsheet has the following tabs:
'  Sheet1 ("Singles") - contains 3 blocks of data: first row is blank, cells B2:G2 have headings (Number, Square root, Straight line, Square, Cube, Cubic root), data block 1 in B3:G7, data block 2 in B10:G14 and data block 3 in B17:G21; the integers 1 - 15 are entered in B3:B7, B10:B14 and B17:B21
'  Sheet2 ("Doubles") - contains 3 blocks of data: first row is blank, cells B2:G2 have headings (Number, Square root, Straight line, Square, Cube, Cubic root), data block 1 in B3:G7, data block 2 in B10:G14 and data block 3 in B17:G21; the even integers 2 - 30 are entered in B3:B7, B10:B14 and B17:B21
'  Sheet3 ("Odds") - contains 3 blocks of data: first row is blank, cells B2:G2 have headings (Number, Square root, Straight line, Square, Cube, Cubic root), data block 1 in B3:G7, data block 2 in B10:G14 and data block 3 in B17:G21; the odd integers 1 - 29 are entered in B3:B7, B10:B14 and B17:B21
'  Sheet4 ("Constants") - contains constants to be used in the formulas, in cells B2:C6; headings in column B (Square root, Line, Square, Cube, Cubic root) and values in column C (0.5, 1, 2, 3, 1/3)
    
    SheetID = 1
    SheetTotal = Sheets.Count
    
    Do While SheetID <= SheetTotal - 1
        Sheets(SheetID).Activate
        
        Sheets(SheetID).Range("B1").End(xlDown).Offset(1, 0).Select 'find the beginning of the block of numbers, assuming the first row is blank (how I usually keep them) and the next row has headings
        ff = ActiveCell.Value 'intialize flag that should be raised if the end is reached, signified by an empty cell
        'blockcounter = 0 'initialize counter, for me to see how many blocks of numbers the loop goes through
    
        Do While Not IsEmpty(ff)
            RowID = ActiveCell.Row 'gets the row ID of the beginning of the block
            ColID = ActiveCell.Column 'gets the column ID of the beginning of the block
            'counter = 0 'initialize counter, for me to see how many times the loop runs
            RootNum = ActiveCell.Value 'reads the number from the "Number" column, B
     
            Do While Not IsEmpty(RootNum) 'loop condition to check if the cell in column B is empty or not; run only if not empty
                'Sheet1.Cells(RowID, ColID - 1).Value = RootNum 'pseudo-code - prints the current value of RootNum in col A
            
                ConstID = 2 'initialize row number of power in the "Constants" sheet
    
                Do While ColID <= 6 'because I know how many columns there are in each tab
                    Sheets(SheetID).Cells(RowID, ColID + 1).Value = "=" & "B" & RowID & "^Constants!$C$" & ConstID 'enters the formula of the type "=B3^Constants!$C$2" for RowID = 3 and ConstID = 2
                    ColID = ColID + 1
                    ConstID = ConstID + 1
                    Loop 'the loop enters the formulas in the rest of the columns for the current row in the data block
            
                ColID = ActiveCell.Column 'intialize counter for the next loop iteration
                ConstID = 2 'intialize counter for the next loop iteration
            
                RowID = RowID + 1 'will move to the next row to check if empty in the next line and loop condition
                RootNum = Sheet1.Cells(RowID, ColID).Value
                'counter = counter + 1 'just a normal loop counter
                Loop
    
            ActiveCell.End(xlDown).Offset(0, 0).Select 'find the end of the current block of numbers
            ActiveCell.End(xlDown).Offset(0, 0).Select 'find the beginning of the next block of numbers (assuming no headings repeated) or end of column if all number blocks have been identified
            'NB: ActiveCell.End(xlDown).Offset(0, 0).Select is executed twice because I usually leave two empty rows between data blocks, to allow for calculation of averages of each column
             
            ff = ActiveCell.Value 'if empty will be the end of the data on the tab
            'blockcounter = blockcounter + 1
            Loop
    
        Sheets(SheetID).Range("B1").End(xlDown).Offset(1, 0).Select 'goes back to the beginning of the first data block on current tab before moving on to the next tab
        SheetID = SheetID + 1
        Loop
    
    Sheets(1).Activate
    Sheets(1).Range("B1").End(xlDown).Offset(1, 0).Select 'goes back to the beginning of the first data block on first tab before execution of code is complete
    MsgBox "All done!"
    
End Sub
