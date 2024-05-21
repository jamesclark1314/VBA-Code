Attribute VB_Name = "Module1"
Sub copy()
    ' Create a copy of sheet and rename
    Sheets("Percentile Rankings").copy After:=Sheets(3)
    Sheets("Percentile Rankings (2)").Name = "Copy"
End Sub

Sub copy_paste()
    ' Copy the most recent spread from Sec Credit Data
    ' and paste to corresponding column in Percentile Rankings

    Dim sec_credit_col As Range
    Dim perc_col As Range

    'Iterate through each column in Sec Credit Data sheet
    For Each sec_credit_col In Range("A:AB").Columns

        'Save the value in last row of column to a variable
        last_row = Worksheets("Sec Credit Data").Cells(Rows.Count, sec_credit_col.column).End(xlUp).Value
        
        'Save the vlue of column header in variable
        If IsEmpty(Worksheets("Sec Credit Data").Cells(4, sec_credit_col.column)) = False Then
            col_header = Worksheets("Sec Credit Data").Cells(4, sec_credit_col.column).Value
'            Debug.Print col_header
        Else
        End If
'        Debug.Print last_row

        'Paste value of last_row to corresponding column in Percentile Rankings sheet

        ' Find the date column and paste to corresponding date column for each subsector in PR sheet
        If Worksheets("Sec Credit Data").Cells(1, sec_credit_col.column).Value = 1 Then
                
            For Each perc_col In Range("A:CD").Columns
            
                If IsEmpty(Worksheets("Copy").Cells(5, perc_col.column)) = False Then
                    Worksheets("Copy").Cells(Rows.Count, perc_col.column).End(xlUp).Offset(1, 0) = last_row
                End If
            
            Next perc_col

        Else

            For Each perc_col In Worksheets("Copy").Range("A:CD").Columns

                If IsEmpty(Worksheets("Sec Credit Data").Cells(4, sec_credit_col.column)) = False And Worksheets("Copy").Cells(5, perc_col.column).Value = col_header Then
                    Worksheets("Copy").Cells(Rows.Count, perc_col.column).End(xlUp).Offset(0, 1).Value = last_row
'                    Debug.Print (last_row)
                Else
                End If

            Next perc_col

        End If

    Next sec_credit_col

End Sub

Sub highlight()
    ' Remove all existing highlights from page
    Worksheets("Copy").Cells.Interior.ColorIndex = 0
    
    ' Highlight the most recent entry for each subsector
    Dim subsector As Range
    
    For Each subsector In Range("A:CD").Columns
        If IsEmpty(Worksheets("Copy").Cells(5, subsector.column)) = False Then
            Worksheets("Copy").Cells(Rows.Count, subsector.column).End(xlUp).Offset(0, 1).Interior.Color = RGB(255, 255, 0)
            Worksheets("Copy").Cells(Rows.Count, subsector.column).End(xlUp).Offset(0, 0).Interior.Color = RGB(255, 255, 0)
            
            ' Center the spread
            Worksheets("Copy").Cells(Rows.Count, subsector.column).End(xlUp).Offset(0, 1).HorizontalAlignment = xlCenter
        Else
        End If
    Next subsector
End Sub

Sub sort()
    Dim subsector As Range
    
    ' Sort each subsector by spread in ascending order
    For Each subsector In Range("A:CD").Columns
        If Worksheets("Copy").Cells(12, subsector.column) = "Date" Then
            Worksheets("Copy").Cells(12, subsector.column).CurrentRegion.sort Key1:=Cells(12, subsector.column).Offset(0, 1), Order1:=xlAscending, Header:=xlYes
        End If
    Next subsector
End Sub

Sub rank1()
    Dim subsector As Range
    
    ' Delete everything in the rank column then re-rank
    For Each subsector In Range("A:CF").Columns
        If Worksheets("Copy").Cells(12, subsector.column) = "Rank" Then
            Worksheets("Copy").Range(Worksheets("Copy").Cells(12, subsector.column).Offset(1, 0), Worksheets("Copy").Cells(12, subsector.column).End(xlDown)).ClearContents
        End If
    Next subsector
End Sub

Sub rank2()
    Dim column As Range
    Dim row As Range
    Dim row_count As Long
    Dim iteration As Long
    
    ' Fill in the rank column
    iteration = 1
    For Each column In Range("A:CF").Columns
    
            row_count = Worksheets("Copy").Cells(Rows.Count, iteration).End(xlUp).row
            Debug.Print row_count
            
        If Worksheets("Copy").Cells(12, column.column).Offset(0, 1) = "Rank" Then
            Worksheets("Copy").Cells(13, column.column).Offset(0, 1) = "1"
            
            For Each row In Range(Cells(14, column.column).Offset(0, 1), Cells(row_count, column.column).Offset(0, 1))
                row.Value = Cells(row.row - 1, column.column).Offset(0, 1).Value + 1
            Next row
    
        End If
        iteration = iteration + 1
        
    Next column
End Sub

Sub all_subs()
    Call copy
    Call rank1
    Call copy_paste
    Call highlight
    Call sort
    Call rank2
End Sub

