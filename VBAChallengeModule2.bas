Attribute VB_Name = "Module1"
'cell/range.Interior.ColorIndex = 3
'colorindex 3 = Red
'colorindex 4 = green

Sub ticker()

Dim ws As Worksheet
      
    
    
    'Min_Max table
    Dim prange As Range
    'Dim maxp As Double
    'Dim minp As Double
    
    Dim vrange As Range
    'Dim maxvol As Double

For Each ws In Worksheets

    'ticker symbol
    Dim ticker_symbol As String
    
    'stock volume
    Dim stock_volume As Double
    stock_volume = 0
    
    'Table Row Tracker
    Dim results_table_row As Integer
    results_table_row = 2
    
    'Open and Close Date Values
    Dim opendateval As Double
    Dim closedateval As Double
    Dim changevalue As Double
    Dim percentchange As Double
    
    
    'open date value
    opendateval = ws.Cells(2, 3)

    'new column labels
    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Yearly Change"
    ws.Cells(1, 11) = "Percentage Change"
    ws.Cells(1, 12) = "Total Stock Volume"
    ws.Cells(1, 15) = "Ticker"
    ws.Cells(1, 16) = "Value"
    
    'set range
    Set vrange = ws.Range("L:L")
    Set prange = ws.Range("K:K")
    
    maxp = Application.WorksheetFunction.Max(prange)
    minp = Application.WorksheetFunction.Min(prange)
    maxvol = Application.WorksheetFunction.Max(vrange)
    
    Dim LR As String
    LR = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
        
    For i = 2 To LR
    
                               
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            ticker_symbol = ws.Cells(i, 1).Value
            
            closedateval = ws.Cells(i, 6).Value
            changevalue = closedateval - opendateval
            stock_volume = stock_volume + ws.Cells(i, 7).Value
            
            ws.Range("I" & results_table_row).Value = ticker_symbol
            ws.Range("J" & results_table_row).Value = changevalue
            If ws.Range("J" & results_table_row).Value < FormatPercent(0) Then
                    ws.Range("J" & results_table_row).Interior.ColorIndex = 3
                Else: ws.Range("J" & results_table_row).Interior.ColorIndex = 4
                End If
            ws.Range("K" & results_table_row).Value = FormatPercent((changevalue / opendateval))
            ws.Range("L" & results_table_row).Value = stock_volume
                            
            results_table_row = results_table_row + 1
                        
            stock_volume = 0
            opendateval = ws.Cells(i + 1, 3)
            
            
            
        Else
            
            stock_volume = stock_volume + ws.Cells(i, 7).Value
        
        End If
        
        
    Next i
'Find Max increase and Ticker
   For i = 2 To LR
        If ws.Cells(i, 11) = Application.WorksheetFunction.Max(prange) Then
            ws.Cells(2, 14).Value = "Greatest % Increase"
            ws.Cells(2, 15).Value = ws.Cells(i, 9)
            ws.Cells(2, 16).Value = Application.WorksheetFunction.Max(prange)
        Exit For
        
        End If
   Next i

'Find max decrerase and Ticker
      For i = 2 To LR
        If ws.Cells(i, 11) = Application.WorksheetFunction.Min(prange) Then
            ws.Cells(3, 14).Value = "Greatest % Decrease"
            ws.Cells(3, 15).Value = ws.Cells(i, 9)
            ws.Cells(3, 16).Value = Application.WorksheetFunction.Min(prange)
        Exit For
        
        End If
'Find Max Vol and Ticker
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
