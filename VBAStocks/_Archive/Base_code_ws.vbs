Sub Stocks_Tracker()

    Dim Summary_Table_row, LastRow, LastCol, TickerCount, TransYear, TransYear2 As Long
    Dim StocksVolume, First, Last, YearlyChange, PerChange As Double
    Dim Ticker As String
    Dim TickerCol, YearlyChangeCol, PercentChangeCol, StocksVolumeCol as Integer

    'Assign variables for summary table
    TickerCol = 9
    YearlyChangeCol = 10
    PercentChangeCol = 11
    StocksVolumeCol = 12

    'loop through each worksheet
    For Each ws In Worksheets
        
        ' Get the last row and last column in the worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        LastCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        
        
        'Name headers in summary table
        
        ws.Cells(1,TickerCol).Value = "Ticker"
        ws.Cells(1,YearlyChangeCol).Value = "Yearly Change"
        ws.Cells(1,PercentChangeCol).Value = "Percent Change"
        ws.Cells(1,StocksVolumeCol).Value = "Total Stock Volume"
        
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
                
                ws.Cells(Summary_Table_row, TickerCol).Value = ws.Cells(i, 1).Value
                ws.Cells(Summary_Table_row, YearlyChange).Value = YearlyChange
                ws.Cells(Summary_Table_row, PercentChangeCol).Value = PercentChange
                ws.Cells(Summary_Table_row, StocksVolumeCol).Value = StocksVolume
                
                'Change number formats
                
                ws.Cells(Summary_Table_row, YearlyChangeCol).NumberFormat = "0.00"
                ws.Cells(Summary_Table_row, PercentChangeCol).NumberFormat = "0.00%"
                ws.Cells(Summary_Table_row, StocksVolumeCol).NumberFormat = "###,###,###,###"
                
                'Conditional format for yearly change column
                
                If ws.Cells(Summary_Table_row, YearlyChangeCol).Value <= 0 Then
                    ws.Cells(Summary_Table_row, YearlyChangeCol).Interior.ColorIndex = 3
                Else
                    ws.Cells(Summary_Table_row, YearlyChangeCol).Interior.ColorIndex = 4
                End If
                
                'Increase summary table row
                
                Summary_Table_row = Summary_Table_row + 1
                
                'Reset storage variables
                
                StocksVolume = 0
                First = 0
                Last = 0
                
            End If
    
        Next i

        ' Cycle through summary table to obtain data

        ' Create variable for summary table
        LastRowTable = Cells(Rows.Count, TickerCol).End(xlUp).Row
        Dim MaxPerIncr, MaxPerDecr, MaxTotStockVol As Double
        Dim CellTracker, CellTracker2, CellTracker3  As Long
        
        MaxPerIncr = 0
        MaxPerDecr = 0
        MaxTotStockVol = 0

        'Find the largerst percent increase
        For j = 2 To LastRowTable
            If Cells(j, YearlyChangeCol).Value > MaxPerIncr Then
                MaxPerIncr = Cells(j, YearlyChangeCol).Value
                CellTracker = j
            End If
        Next j
        
        'Find the largest percent decrease
        For j = 2 To LastRowTable
            If Cells(j, YearlyChangeCol).Value < MaxPerDecr Then
                MaxPerDecr = Cells(j, YearlyChangeCol).Value
                CellTracker2 = j
            End If
        Next j
        'Find the largest percent decrease
        For j = 2 To LastRowTable
            If Cells(j, StocksVolumeCol).Value > MaxTotStockVol Then
                MaxTotStockVol = Cells(j, StocksVolumeCol).Value
                CellTracker3 = j
            End If
        Next j
        
        'Make second summary or results
        '-------------------------------------------
        'Row and column headings
        Range("N2").Value = "Greatest % Increase"
        Range("N3").Value = "Greatest % Decrease"
        Range("N4").Value = "Greatest Total Stock Volume"
        Range("O1").Value = "Ticker"
        Range("P1").Value = "Value"

        'Populate table
        'Greatest increase
        Range("O2").Value = Cells(CellTracker, TickerCol).Value
        Range("P2").Value = MaxPerIncr
        'Greatest decrease
        Range("O3").Value = Cells(CellTracker2, TickerCol).Value
        Range("P3").Value = MaxPerDecr
        'Greatest increase
        Range("O4").Value = Cells(CellTracker3, TickerCol).Value
        Range("P4").Value = MaxTotStockVol

    Next ws

End Sub
