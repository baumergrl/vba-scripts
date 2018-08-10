Sub RowToColumn()
'
' RowToColumn Macro
' Select the first cell of a row that you wish to turn into a column.
'

'
    Dim actCol As Integer
    Dim actRow As Integer
    Dim newCol As Integer
    Dim newRow As Integer

    actCol = ActiveCell.Column
    actRow = ActiveCell.row
    newCol = ActiveCell.Column
    newRow = ActiveCell.row
    
    Do While IsEmpty(Cells(actRow, newCol + 1)) = False
        newCol = newCol + 1
        newRow = newRow + 1
        
        Cells(actRow, newCol).Cut Cells(newRow, actCol)
        
    Loop
    
End Sub
