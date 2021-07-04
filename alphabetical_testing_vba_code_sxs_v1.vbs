Sub stonks()
    'populate the row headers for summary data
    Range("i1").Value = "Ticker"
    Range("j1").Value = "Yearly_Change"
    Range("k1").Value = "Percent_Change"
    Range("l1").Value = "Total_Stock_Volume"
    
    'set variable to hold stock ticker
    Dim stock_ticker As String
    'set variable to hold opening price - may not need
    Dim stock_open As Double
    'set variable to hold closing price - may not need
    Dim stock_close As Double
    'set variable to hold total volume
    Dim total_vol As Double
    'set total_vol to reset each time
    total_vol = 0
    'set variable for summary table row
    Dim summary_table_row As Integer
    'set summary table row to start at row 2
    summary_table_row = 2
    'set variable to find last row
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    
    'loop thru data to pull each unique ticker and total vol
        For i = 2 To last_row
        'store the first open in stock open and last close in stock close
            If Cells(i, 2).Value = "20160101" Then
            stock_open = Cells(i, 3).Value
            ElseIf Cells(i, 2).Value = "20161230" Then
            stock_close = Cells(i, 6).Value
            End If

            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            'store ticker in stock_ticker
            stock_ticker = Cells(i, 1).Value
            'sum total volume
            total_vol = total_vol + Cells(i, 7).Value
            'send ticker to the summary table
            Range("I" & summary_table_row).Value = stock_ticker
            'send total vol to the summary table
            Range("L" & summary_table_row).Value = total_vol
            'send annual price change to summary table
            Range("J" & summary_table_row).Value = stock_close - stock_open
            'send annual % change to summary table
            Range("K" & summary_table_row).Value = (stock_close / stock_open) - 1
            'add one to summary table row so it starts the next line
            summary_table_row = summary_table_row + 1
            'reset total volume sum
            total_vol = 0
            
            'if the next ticker row is the same
            Else
            'sum total volume
            total_vol = total_vol + Cells(i, 7).Value
            
            End If
        Next i
        
    'evaluate results to apply conditional formatting
    For j = 2 To 4
        For k = 10 To 11
            'if value >= 0 then green
            If Cells(j, k).Value >= 0 Then
            Cells(j, k).Interior.ColorIndex = 4
            Else
            'if value <= 0 then red
            Cells(j, k).Interior.ColorIndex = 3
            End If
        Next k
        
    
    Next j
    
    
    
    'autofit column width for summary data
    Worksheets("test data").Columns("I:L").AutoFit
    'apply currency format to yearly_change
    Worksheets("test data").Columns("J:J").Style = "Currency"
    'apply percent format to percent_change
    Worksheets("test data").Columns("K:K").Style = "Percent"
    'apply number with commas format to total_stock_volume
    Worksheets("test data").Columns("L:L").NumberFormat = "#,##0"
End Sub
