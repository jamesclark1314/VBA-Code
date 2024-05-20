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
                    Debug.Print (last_row)
                Else
                End If

            Next perc_col

        End If

    Next sec_credit_col

End Sub

Sub all_subs()
    Call copy
    Call copy_paste
End Sub


