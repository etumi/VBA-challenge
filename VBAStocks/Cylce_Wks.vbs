Sub Stocks_Tracker()

    Dim Summary_Table_row, LastRow, LastCol, TickerCount, TransYear, TransYear2 As Long
    Dim StocksVolume, First, Last, YearlyChange, PerChange As Double
    Dim Ticker As String

    'loop through each worksheet
    For Each ws In Worksheets
        
        ' Get the last row and last column in the worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        LastCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        
        
        'Name summary table
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ' First row to use in summary table
        Summary_Table_row = 2
    
        'Set the first price of each ticker to 0
        First = 0
    
        'loop through all cells
        For i = 2 To LastRow
        
            'Compare year or each cell to make sure you are within the same year
            TransYear = Left$(Str(ws.Cells(i, 2).Value), 5) 'compare year or each cell to make sure you are within the same year
            TransYear2 = Left$(Str(ws.Cells(i + 1, 2).Value), 5)
    
           ' For cells with the same ticker
           If (ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value) And (TransYear = TransYear2) Then
           
                If First = 0 Then
                    First = ws.Cells(i, 3).Value
                End If
                
                StocksVolume = StocksVolume + ws.Cells(i, 7).Value 'Cummulate stock volume
                
            ElseIf (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
                
                StocksVolume = StocksVolume + ws.Cells(i, 7).Value 'Cummulate stock volume
    
                Last = ws.Cells(i, 6).Value 'Assign last price for ticker in the year
                
                YearlyChange = Last - First
                
                'Calculate the percentage change and account for zero price
                
                If First = 0 Then
                    PercentChange = 0
                Else
                    PercentChange = ((Last / First) - 1)
                End If
                
                'Populate summary table
                
                ws.Cells(Summary_Table_row, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(Summary_Table_row, 10).Value = YearlyChange
                ws.Cells(Summary_Table_row, 11).Value = PercentChange
                ws.Cells(Summary_Table_row, 12).Value = StocksVolume
                
                'Change number formats
                
                ws.Cells(Summary_Table_row, 10).NumberFormat = "0.00"
                ws.Cells(Summary_Table_row, 11).NumberFormat = "0.00%"
                ws.Cells(Summary_Table_row, 12).NumberFormat = "###,###,###,###"
                
                'Conditional format for yearly change column
                
                If ws.Cells(Summary_Table_row, 10).Value <= 0 Then
                    ws.Cells(Summary_Table_row, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(Summary_Table_row, 10).Interior.ColorIndex = 4
                End If
                
                'Increase summary table row
                
                Summary_Table_row = Summary_Table_row + 1
                
                'Reset storage variables
                
                StocksVolume = 0
                First = 0
                Last = 0
                
            End If
    
        Next i

    Next ws

End Sub
