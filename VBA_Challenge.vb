Sub WednesdayAllStocksAnalysisRefactored()

'this code works.

    Dim startTime As Single
    Dim endTime As Single

    yearValue = InputBox("What year would you like to run your analysis on?")
    
    startTime = Timer
    
    'Format the All Stocks Analysis Worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks(" + yearValue + ")"
    
    'Create header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows of stock data
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create tickerIndex to access array of tickers
    Dim tickerIndex As Integer
    tickerIndex = 0
    
    '1b) Create output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    '2a) Set initial volume to zero for all tickers
    For i = 0 To 11
        
        tickerVolumes(i) = 0
        
    Next i
    
    '2b) Loop over all rows in the yearValue spreadsheet
    Worksheets(yearValue).Activate
    For i = 2 To RowCount
    
        '3a) Increase total volume for current ticker
        If Cells(i, 1).Value = tickers(tickerIndex) Then
            
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            
        End If
        
        '3b) Check if current row is the first row for selected ticker to find startingPrices
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If
        
        '3c) Check if current row is last row for selected ticker to find endingPrices
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        'If the next row's ticker doesn't match, increase the tickerIndex
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            
            '3d) Increase the tickerIndex
            tickerIndex = tickerIndex + 1
        
        End If
    
    Next i
    
    '4) Loop through the output arrays to find the Ticker, Total Daily Volume, and Return
    For i = 0 To 11
    
        Worksheets("All Stocks Analysis").Activate
        'use tickerIndex to access each position in the output array
        tickerIndex = i
        Cells(4 + i, 1).Value = tickers(tickerIndex)
        Cells(4 + i, 2).Value = tickerVolumes(tickerIndex)
        Cells(4 + i, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
    
    Next i
    
    
    'Formatting the output sheet
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:c3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("b4:b15").NumberFormat = "#,##0"
    Range("c4:c15").NumberFormat = "0.0%"
    Columns("b").AutoFit
    
    dataRowStart = 4
    dataRowEnd = 15
    
    For i = dataRowStart To dataRowEnd
    
        If Cells(i, 3) > 0 Then
        
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
        
        End If
    
    Next i
    
    endTime = Timer
    MsgBox ("This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue))

End Sub
