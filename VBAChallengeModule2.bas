Attribute VB_Name = "Module1"
 
Sub ticker()

'establish ws
Dim ws As Worksheet
   
    'establish Min_Max table elements
    Dim prange As Range
    Dim vrange As Range
    

For Each ws In Worksheets

    'establish ticker variable
    Dim ticker_symbol As String
    
    'establish and define stock volume
    Dim stock_volume As Double
    stock_volume = 0
    
    'establish and define table row
    Dim results_table_row As Integer
    results_table_row = 2
    
    'establish open and close date values
    Dim opendateval As Double
    Dim closedateval As Double
    Dim changevalue As Double
    Dim percentchange As Double
    
    
    'define open date value
    opendateval = ws.Cells(2, 3)

    'print new column labels
    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Yearly Change"
    ws.Cells(1, 11) = "Percentage Change"
    ws.Cells(1, 12) = "Total Stock Volume"
    ws.Cells(1, 15) = "Ticker"
    ws.Cells(1, 16) = "Value"
    
    'set range for Min_Max table elements
    Set vrange = ws.Range("L:L")
    Set prange = ws.Range("K:K")
    
    'establish last row count
    Dim LR As String
    LR = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
        
    For i = 2 To LR
    
        'argument - if current cell does not equal next cell then
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'define ticker symbol
            ticker_symbol = ws.Cells(i, 1).Value
            
            'define value of close date, value of close date and open date difference, and add the last stock volume to running sum
            closedateval = ws.Cells(i, 6).Value
            changevalue = closedateval - opendateval
            stock_volume = stock_volume + ws.Cells(i, 7).Value
            
            'print ticker symbol, change value, cell color, change percent, and stock volume
            ws.Range("I" & results_table_row).Value = ticker_symbol
            ws.Range("J" & results_table_row).Value = changevalue
            If ws.Range("J" & results_table_row).Value < 0 Then
                    ws.Range("J" & results_table_row).Interior.ColorIndex = 3
                    ws.Range("K" & results_table_row).Interior.ColorIndex = 3
                Else: ws.Range("J" & results_table_row).Interior.ColorIndex = 4
                      ws.Range("K" & results_table_row).Interior.ColorIndex = 4
                End If
            ws.Range("K" & results_table_row).Value = FormatPercent((changevalue / opendateval))
            ws.Range("L" & results_table_row).Value = stock_volume
                            
            'define the new table row
            results_table_row = results_table_row + 1
                        
            'reset stock volume and define the new open date value
            stock_volume = 0
            opendateval = ws.Cells(i + 1, 3)
            
            
            
        Else
            
            'add to running sum of stock volume
            stock_volume = stock_volume + ws.Cells(i, 7).Value
        
        End If
        
        
    Next i
    
'Find greatest increase, print the label, ticker and result
   For i = 2 To LR
        If ws.Cells(i, 11) = Application.WorksheetFunction.Max(prange) Then
            ws.Cells(2, 14).Value = "Greatest % Increase"
            ws.Cells(2, 15).Value = ws.Cells(i, 9)
            ws.Cells(2, 16).Value = Application.WorksheetFunction.Max(prange)
        Exit For
        
        End If
   Next i

'Find greatest decrease, print the label, ticker and result
      For i = 2 To LR
        If ws.Cells(i, 11) = Application.WorksheetFunction.Min(prange) Then
            ws.Cells(3, 14).Value = "Greatest % Decrease"
            ws.Cells(3, 15).Value = ws.Cells(i, 9)
            ws.Cells(3, 16).Value = Application.WorksheetFunction.Min(prange)
        Exit For
        
        End If
        
'Find greatest total volume, print the label, Ticker and result
   Next i
      For i = 2 To LR
        If ws.Cells(i, 12) = Application.WorksheetFunction.Max(vrange) Then
            ws.Cells(4, 14).Value = "Greatest Total Volume"
            ws.Cells(4, 15).Value = ws.Cells(i, 9)
            ws.Cells(4, 16).Value = Application.WorksheetFunction.Max(vrange)
        Exit For
        
        End If
   Next i
   

Next ws

  
End Sub
