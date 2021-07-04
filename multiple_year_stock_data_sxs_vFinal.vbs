Sub stonks()
    Dim w As Worksheet
    
    For Each w In ActiveWorkbook.Worksheets
    w.Activate
    
    'populate the row headers for summary data
    Range("i1").Value = "Ticker"
    Range("j1").Value = "Yearly_Change"
    Range("k1").Value = "Percent_Change"
    Range("l1").Value = "Total_Stock_Volume"
    
    'set variable to hold stock ticker
    Dim stock_ticker As String
    'set variable to hold opening price
    Dim stock_open As Double
    stock_open = Cells(2, 3).Value
    'set variable to hold closing price
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
                If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                'store ticker in stock_ticker
                stock_ticker = Cells(i, 1).Value
                'store close price
                stock_close = Cells(i, 6).Value
                'sum total volume
                total_vol = total_vol + Cells(i, 7).Value
                'send ticker to the summary table
                Range("I" & summary_table_row).Value = stock_ticker
                'send total vol to the summary table
                Range("L" & summary_table_row).Value = total_vol
                'send annual price change to summary table
                Range("J" & summary_table_row).Value = stock_close - stock_open
                'send annual % change to summary table
                    If stock_open = 0 Then
                    Range("K" & summary_table_row).Value = 1
                    Else
                    Range("K" & summary_table_row).Value = (stock_close / stock_open) - 1
                    End If
                'add one to summary table row so it starts the next line
                summary_table_row = summary_table_row + 1
                'reset total volume sum
                total_vol = 0
                'set open for next ticker
                stock_open = Cells(i + 1, 3)
                'if the next ticker row is the same
                Else
                'sum total volume
                total_vol = total_vol + Cells(i, 7).Value
                End If
                
            Next i
            
        'define new last row
        new_last_row = Cells(Rows.Count, 9).End(xlUp).Row
        'evaluate results to apply conditional formatting
        For j = 2 To new_last_row
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
    Columns("I:L").AutoFit
    'apply currency format to yearly_change
    Columns("J:J").Style = "Currency"
    'apply percent format to percent_change
    Columns("K:K").NumberFormat = "0.00%"
    'apply number with commas format to total_stock_volume
    Columns("L:L").NumberFormat = "#,##0"
        
    Next w
    
  
  
  
  
End Sub



