Attribute VB_Name = "MExample"
Option Explicit

Sub ExampleWorkTable()
    Dim table       As CTable
    Dim wsMain      As Worksheet
    Dim oColumns    As Object
    Dim i           As Integer
    
    Set wsMain = ThisWorkbook.Sheets(MConstants.SHEET_MAIN)

    'init table
    Set table = New CTable
    table.InitClass wsMain, "headerCell"
    
    If Not table.IsAccesTable Then ' Table exists?
        Exit Sub
    End If
    
    If Not table.IsEmptyTable Then ' Table empty?
        Exit Sub
    End If
    
    'Allows you to determine the column number taking into account changing labels
    Set oColumns = table.GetColumns

    table.ToString 'info table

    With wsMain
        
        .Cells(7, oColumns("C16")).Value = vbNullString

        Debug.Print .Cells(4, oColumns("C10")).Value ' -> Exp_1
        Debug.Print .Cells(5, oColumns("C12")).Value ' -> val_2
        
        '-> Exp_1, exp_2 ...
        For i = table.GetFirstRow + 1 To table.GetLastRow
            Debug.Print .Cells(i, table.GetFirstColumn).Value
        Next i
        
        .Cells(7, oColumns("C16")).Value = "1" 'Write Value
        
    End With
    
    Debug.Print table.GetFirstRow '-> 3
    
    table.SetCellUpperBound 5, 5 'New position for first row
    Debug.Print table.GetFirstRow '-> 5 - New position


End Sub


