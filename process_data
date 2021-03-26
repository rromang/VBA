Sub process_data()

'Sorting all data
Application.ScreenUpdating = False 'worksheet is not visible when running the script
Application.DisplayAlerts = False 'alerts are not displayed


Dim wsh As Worksheet 'establishes a variable for worksheet type

For Each wsh In Worksheets 'runs the subroutine to sort all data in all sheets
    wsh.Select
    
    Dim lastrow As Long
    
    'count all rows in ticker column
    lastrow = WorksheetFunction.CountA(Range("A:A"))
    'MsgBox lastrow
    
       
    'sorts based on first column in ascending order
    With ActiveSheet.Sort
         .SortFields.Add Key:=Range("A1"), Order:=xlAscending
         .SortFields.Add Key:=Range("B1"), Order:=xlAscending
         .SetRange Range("A1", "G" & lastrow)
         .Header = xlYes
         .Apply
    End With
Next
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Finds unique ticker types
For Each wsh In Worksheets 'runs the subroutine to remove duplicate values in ticker for all sheets
    wsh.Select
    Dim ticker_rng As Range
    
    ActiveSheet.Range("A:A").Copy Range("J:J") 'copy range with repeated tickers into new column
    
    Range("J1").Value = "Ticker" 'places title for ticker in summary table
    ActiveSheet.Range("J2", Range("J2").End(xlDown)).Select 'selects the newly pasted ticker column in summary table
    Set ticker_rng = Selection 'sets ticker range to selected cells
    With ticker_rng
        .RemoveDuplicates Columns:=1, Header:=xlNo 'open remove duplicates windows w/o,_
        'headear selection(alert window still opens in my excel for Mac)
    End With
    
    Columns(10).EntireColumn.AutoFit
Next
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Finds percent change
For Each wsh In Worksheets 'runs the subroutine to remove duplicate values in ticker for all sheets
    wsh.Select
    Dim lastrow_ticker As Long
    Dim lastrow_summary As Long
    Dim summary_rng As Range
    
    
    lastrow_ticker = WorksheetFunction.CountA(Range("A:A"))
    Set ticker_rng = Range("A2", "A" & lastrow_ticker) 'establishes the range of ticker
    
    lastrow_summary = WorksheetFunction.CountA(Range("J:J"))
    Set summary_rng = Range("J2", "J" & lastrow_summary) 'establishes the range of summary table
    
    startrow = 2
    'loop for each ticker value to find percent change
    For j = 2 To lastrow_summary
        'loop for each row in ticker block to find open value and close value
        For i = startrow To lastrow_ticker
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then 'checks if the next row is not equal to previous row
                endrow = Cells(i, 1).Row                                'establishes the last row of ticker of the same kind
                open_value = Cells(startrow, 3).Value            'gets the value of the first cell in the block of tickers of the same kind (only possible if values are sorted)
                close_value = Cells(endrow, 6).Value             'gets the value of the first cell in the block of tickers of the same kind (only possible if values are sorted)
                startrow = endrow + 1                                  'resets the start row to start at the beginning of the next ticker block
                i = lastrow_ticker                                          'forces vba to set i to the last loop condition once the rest of the statements were run so it get out of the loop
            End If
            
        Next i
        If open_value = 0 Then 'traps error if open value is 0
            percentchange = 0
        Else
        percentchange = (close_value - open_value) / open_value 'sets equation for % change
        Cells(j, 11).Value = percentchange
        End If
    Next j
    
    Range("K1") = "Percent Change"
    Range("K2", "K" & lastrow_summary).NumberFormat = "0.00%"
    
    Columns(11).EntireColumn.AutoFit
Next
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Finds yearly change
For Each wsh In Worksheets 'runs the subroutine to remove duplicate values in ticker for all sheets
    wsh.Select
    
    lastrow_ticker = WorksheetFunction.CountA(Range("A:A"))
    Set ticker_rng = Range("A2", "A" & lastrow_ticker) 'establishes the range of ticker
    
    lastrow_summary = WorksheetFunction.CountA(Range("J:J"))
    Set summary_rng = Range("J2", "J" & lastrow_summary) 'establishes the range of summary table
    
    startrow = 2
    'loop for each ticker value to find percent change
    For j = 2 To lastrow_summary
        'loop for each row in ticker block to find open value and close value
        For i = startrow To lastrow_ticker
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then 'checks if the next row is not equal to previous row
                endrow = Cells(i, 1).Row                                'establishes the last row of ticker of the same kind
                open_value = Cells(startrow, 3).Value            'gets the value of the first cell in the block of tickers of the same kind (only possible if values are sorted)
                close_value = Cells(endrow, 6).Value             'gets the value of the first cell in the block of tickers of the same kind (only possible if values are sorted)
                startrow = endrow + 1                                  'resets the start row to start at the beginning of the next ticker block
                i = lastrow_ticker                                          'forces vba to set i to the last loop condition once if statement is met it gets out of the loop
            End If
            
        Next i
      
        Change = (close_value - open_value)
        Cells(j, 13).Value = Change
    
        If Cells(j, 13).Value > 0 Then   'color codes if values are positive or negative
        Cells(j, 13).Interior.ColorIndex = 4
        Else
        Cells(j, 13).Interior.ColorIndex = 3
        End If
    Next j
    
    Range("M1") = "Yearly Change"
    Columns(13).EntireColumn.AutoFit
    
    
    Columns(13).EntireColumn.AutoFit
Next
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Finds total stock volume
For Each wsh In Worksheets 'runs the subroutine to remove duplicate values in ticker for all sheets
    wsh.Select
    
    Dim total_Vol As Variant
    
    lastrow_ticker = WorksheetFunction.CountA(Range("A:A"))
    Set ticker_rng = Range("A2", "A" & lastrow_ticker) 'establishes the range of ticker
    
    lastrow_summary = WorksheetFunction.CountA(Range("J:J"))
    Set summary_rng = Range("J2", "J" & lastrow_summary) 'establishes the range of summary table
    
    startrow = 2
    'loop for each ticker value to find percent change
    For j = 2 To lastrow_summary
        'loop for each row in ticker block to find open value and close value
        For i = startrow To lastrow_ticker
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then 'checks if the next row is not equal to previous row
                endrow = Cells(i, 1).Row                               'establishes the last row of ticker of the same kind
                vol_rng = Range("G" & startrow, "G" & endrow) 'creates a range for volume
                total_Vol = WorksheetFunction.Sum(vol_rng)    'sums all volumes per range category
                
                startrow = endrow + 1                                  'resets the start row to start at the beginning of the next ticker block
                i = lastrow_ticker                                          'forces vba to set i to the last loop condition once,_
                                                                                    'condition is met so it gets out of the loop
            End If
            
        Next i
     
        Cells(j, 12).Value = total_Vol
    
    Next j
    
    Range("L1") = "Total Stock Volume"
    Columns(12).EntireColumn.AutoFit
Next
''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Finds greatest increase, decrease and total volume
For Each wsh In Worksheets 'runs the subroutine to remove duplicate values in ticker for all sheets
    wsh.Select

    Dim percent_rng As Range
    Dim totalstock_rng As Range
    Dim prct_hi, prct_hi_ticker, prct_lo, prct_lo_ticker, stock_vol_max, stockmax_ticker As Variant
    
    
    lastrow_ticker = WorksheetFunction.CountA(Range("A:A"))
    Set ticker_rng = Range("A2", "A" & lastrow_ticker) 'establishes the range of ticker for all data
    
    lastrow_summary = WorksheetFunction.CountA(Range("J:J"))
    Set percent_rng = Range("K2", "K" & lastrow_summary) 'establishes the range of percent values
    Set totalstock_rng = Range("L2", "L" & lastrow_summary) 'establishes the range of stock vol
    
    startrow = 2
        
        For i = startrow To lastrow_ticker
            If Cells(i, 11).Value = Application.WorksheetFunction.Max(percent_rng) Then 'checks if row is equal to the max in the,_
            'designated range (% change)
                prct_hi = Cells(i, 11).Value
                prct_hi_ticker = Cells(i, 10).Value
            End If
            If Cells(i, 11).Value = Application.WorksheetFunction.Min(percent_rng) Then 'checks if row is equal to the min in the<_
            'designated range (% change)
                prct_lo = Cells(i, 11).Value
                prct_lo_ticker = Cells(i, 10)
            End If
             If Cells(i, 12).Value = Application.WorksheetFunction.Max(totalstock_rng) Then 'checks if row is equal to the min in the,_
             'designated range (total stock vol)
                stock_vol_max = Cells(i, 12).Value
                stockmax_ticker = Cells(i, 10).Value
            End If
        Next i
    
        Cells(2, 16).Value = prct_hi
        Cells(3, 16).Value = prct_lo
        Cells(4, 16).Value = stock_vol_max
        Cells(2, 17).Value = prct_hi_ticker
        Cells(3, 17).Value = prct_lo_ticker
        Cells(4, 17).Value = stockmax_ticker
        
    
    Range("O2") = "Greatest % Increase"
    Range("O3") = "Greatest %Decrease"
    Range("O4") = "Greatest Total Volume"
    Range("P1") = "Value"
    Range("Q1") = "Ticker"
    Range("P2").NumberFormat = "0.00%"
    Range("P3").NumberFormat = "0.00%"
    Columns(15).EntireColumn.AutoFit
    Columns(16).EntireColumn.AutoFit
    Columns(17).EntireColumn.AutoFit
Next


Application.ScreenUpdating = True
Application.DisplayAlerts = True


End Sub