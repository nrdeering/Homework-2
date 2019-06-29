Attribute VB_Name = "Module1"
Sub Stock_Market()

Dim Ticker As String

Dim Stock_Volume As Double

Dim Stock_Totals As Integer
Stock_Totals = 2

Dim Open_Stock As Double
Open_Stock = 0

Dim Closing_Stock As Double
Close_Stock = 0

Dim TickerCount As Long
TickerCount = 0

Dim stock_max As Long

Dim stock_min As Long

Dim stock_ticker As Long


    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
       
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
        'Loop through each row
        For i = 2 To LastRow
            
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                'Ticker output
                
                Ticker = Cells(i, 1).Value
                Range("I" & Stock_Totals).Value = Ticker
                
                
                'Yearly change between opening and closing stocks output
             
                Open_Stock = Cells(i - TickerCount, 3).Value
                Close_Stock = Cells(i, 6).Value
                Range("J" & Stock_Totals).Value = Close_Stock - Open_Stock
                Range("J" & Stock_Totals).NumberFormat = "#.000000000"
             
                    'Yearly Change Format
                    If Range("J" & Stock_Totals).Value > 0 Then
                        Range("J" & Stock_Totals).Interior.ColorIndex = 4
                        
                    Else
                        Range("J" & Stock_Totals).Interior.ColorIndex = 3
                    End If
                     
                    'Percentage change output
                    If Open_Stock <> 0 Then
                        Range("K" & Stock_Totals).Value = ((Close_Stock - Open_Stock) / Open_Stock) * 1
                        Range("K" & Stock_Totals).NumberFormat = "0.00%"
                        
                    Else
                        Range("K" & Stock_Totals).Value = "0"
                    End If
                
                    'Stock_Volume Output
                    Stock_Volume = Stock_Volume + Cells(i, 7).Value
                    Range("L" & Stock_Totals).Value = Stock_Volume
                    Stock_Totals = Stock_Totals + 1
                    
                    
                    'Setting Stock_Volume and TickCount back to 0
                    Stock_Volume = 0
                    TickerCount = 0
            Else
                Stock_Volume = Stock_Volume + Cells(i, 7).Value
                TickerCount = TickerCount + Cells(i, 1).Count
            End If
                 
        Next i
        
        'Column Headers
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        
        'Find min/max of percentage increae/decrease and the max stock volume
        max_value = Application.WorksheetFunction.Max(Range("K:K"))
        min_value = Application.WorksheetFunction.Min(Range("K:K"))
        greatest_volume = Application.WorksheetFunction.Max(Range("L:L"))
        
        
        'for loop that loops through to find min/max of % increase/decrease and max stock volume
        For g = 2 To LastRow
            If Cells(g, 11) = max_value Then
                Cells(2, 16).Value = Cells(g, 9)
                Cells(2, 17).Value = max_value
                Cells(2, 17).NumberFormat = "0.00%"
        
            End If
            
            If Cells(g, 11) = min_value Then
                Cells(3, 16).Value = Cells(g, 9)
                Cells(3, 17).Value = min_value
                Cells(3, 17).NumberFormat = "0.00%"
        
            End If
            
            If Cells(g, 12) = greatest_volume Then
                Cells(4, 16).Value = Cells(g, 9)
                Cells(4, 17).Value = greatest_volume
        
            End If
        
        Next g
        
        'Auto fit ranges
        Range("I:L").EntireColumn.AutoFit
        Range("O:Q").EntireColumn.AutoFit
    
End Sub

