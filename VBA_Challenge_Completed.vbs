
Sub AllStocksAnalysisRefactored()

    outputSheet = "All Stocks Refactored"
    
    yearValue = InputBox("What year would you like to run the analysis on?")
    
    Dim startTime As Single
    Dim endTime  As Single

    startTime = Timer
    
    Worksheets(outputSheet).Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    'Initialize array of all tickers
    Dim tickers(11) As String

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
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    'This assumes the data is sorted by ticker
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(11) As Long
    Dim tickerStartingPrices(11) As Single
    Dim tickerEndingPrices(11) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To UBound(tickers)
        tickerVolumes(i) = 0
    Next i
    
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        'get values for logic
        curRow = Cells(i, 1).Value
        nextRow = Cells(i + 1, 1).Value
        prevRow = Cells(i - 1, 1).Value
        volume = Cells(i, 8).Value
        price = Cells(i, 6).Value
                
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + volume
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If curRow = tickers(tickerIndex) And curRow <> prevRow Then
        
            tickerStartingPrices(tickerIndex) = price
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        If curRow <> nextRow Then

            tickerEndingPrices(tickerIndex) = price
            
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To UBound(tickers)
                
        Worksheets(outputSheet).Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
    'Formatting
    Worksheets(outputSheet).Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

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
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
