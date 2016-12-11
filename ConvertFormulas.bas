Attribute VB_Name = "ConvertFormula"
'Converts all selected cells formulas into the actual value
Sub ConvertFormulas()

    'Step 1: Declare variables
    Dim my_range As Range
    Dim my_cell As Range
    
    'Step 2: Save the workbook before changing cells
    Select Case MsgBox("Can't undo this action. " & "Would you like to continue?", vbYesNoCancel)
    
        Case Is = vbYes
            ThisWorkbook.Save
        
        Case Is = vbCancel
            Exit Sub
    End Select
    
    'Step 3: Define the target range
    Set my_range = Selection
    
    'Step 4: Start looping through the range
    For Each my_cell In my_range
    
    'Step 5: If cell has formula, set to the value shown.
    If my_cell.HasFormula Then
        my_cell.Formula = my_cell.Value
    End If
    
    'Step 6: Get the next cell in range
    Next my_cell

End Sub
'Unhide all columns
Sub UnhideAllColumns()
    Columns.EntireColumn.Hidden = False
End Sub
