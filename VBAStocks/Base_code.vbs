Sub Stocks_Tracker()

    Dim Summary_Table_row, LastRow, LastCol, TickerCount, TransYear, TransYear2 As Integer
    Dim StocksVolume, First, Last, YearlyChange, PerChange As Double
    Dim Ticker As String

    'Get the last row and last column in the worksheet
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    LastCol = Cells(1, Columns.Count).End(xlToLeft).Column
    
    'MsgBox ("Last row " + Str(LastRow) + " Last Column " + Str(LastCol))
    
    'Name summary table
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    'first row to use in summary table
    Summary_Table_row = 2

    'Set the first price of each ticker to 0
    First = 0

    'loop through all cells
    For i = 2 To LastRow
        'compare year or each cell to make sure you are within the same year
        TransYear = Left$(Str(Cells(i, 2).Value), 5) 'compare year or each cell to make sure you are within the same year
        TransYear2 = Left$(Str(Cells(i + 1, 2).Value), 5)
        'MsgBox ("TransYear " + TransYear)
        'MsgBox ("TransYear2 " + Str(TransYear2))

       'for cells with the same ticker
       If (Cells(i, 1).Value = Cells(i + 1, 1).Value) And (TransYear = TransYear2) Then
        '    MsgBox ("Year " + TransYear)
        '    MsgBox ("Next row year " + Str(TransYear2))
        '    MsgBox ("Start stocks volume " + Str(StocksVolume))
            If First = 0 Then
                First = Cells(i, 3).Value
            End If
            
            'MsgBox ("First vol " + Str(First))
            
        '    MsgBox ("Stocks Volume total " + Str(StocksVolume))
            
            StocksVolume = StocksVolume + Cells(i, 7).Value
            
            'MsgBox ("First vol " + Str(First))
            
        ElseIf (Cells(i, 1).Value <> Cells(i + 1, 1).Value) Then
            
            StocksVolume = StocksVolume + Cells(i, 7).Value 'Cummulate stock volume

            Last = Cells(i, 6).Value 'assign last cell price
            
            YearlyChange = Last - First

            PerChange = ((Last / First) - 1)
            
            'Populate summary table
            Cells(Summary_Table_row, 9).Value = Cells(i, 1).Value
            Cells(Summary_Table_row, 10).Value = YearlyChange
            Cells(Summary_Table_row, 11).Value = PerChange
            'MsgBox ("total vol " + Str(StocksVolume))
            Cells(Summary_Table_row, 12).Value = StocksVolume
            
            'change number formats
            Cells(Summary_Table_row, 10).NumberFormat = "0.00"
            Cells(Summary_Table_row, 11).NumberFormat = "0.00%"
            Cells(Summary_Table_row, 12).NumberFormat = "###,###,###,###"
            
            'conditional format for yearly change column
            If Cells(Summary_Table_row, 10).Value <= 0 Then
            
                Cells(Summary_Table_row, 10).Interior.ColorIndex = 3
                
            Else
            
                Cells(Summary_Table_row, 10).Interior.ColorIndex = 4
                
            End If
            
            'Increase summary table row
            Summary_Table_row = Summary_Table_row + 1
            
            'reset storage variables
            StocksVolume = 0
            First = 0
            Last = 0
            
         End If

Next i

End Sub

