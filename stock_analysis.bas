Attribute VB_Name = "Module1"
Sub AllStocksAnalysis()
    yearValue = InputBox("What year would you like to run the analysis on?")

    Worksheets("All Stocks Analysis").Activate

    Range("A1").Value = "All Stocks (" + yearValue + ")"

    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    'Create tickerIndex and tickerCount variable
    tickerIndex = 0
    tickerCount = 0
    ticker = " "

    Worksheets(yearValue).Activate

    'get the number of rows to loop over, thanks StackOverflow!
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    'Count the number of unique stock tickers
    For i = 2 To RowCount
        
        If ticker <> Cells(i, 1).Value Then
            
            tickerCount = tickerCount + 1
            ticker = Cells(i, 1).Value
            
        End If
        
    Next i
        
    'Arrays can only be ReDimmed with a variable length so first the must be dimmed as 0 length arrays
    Dim tickers() As String
    Dim startingPrices() As Single
    Dim endingPrices() As Single
    ReDim tickers(tickerCount)
    ReDim startingPrices(tickerCount)
    ReDim endingPrices(tickerCount)
    
    'Initialize the totalVolumes array with values of 0
    Dim totalVolumes() As Single
    ReDim totalVolumes(tickerCount)
    For i = 1 To tickerCount
        
        totalVolumes(i - 1) = 0
        
    Next i
    
    'Reset ticker variable
    ticker = " "
    
    'Loop through entire stock data set with any number of tickers to track volume and return info
    For i = 2 To RowCount
    
        'If the ticker value does not match then move to the next ticker entry
        If ticker <> Cells(i, 1).Value Then
            ticker = Cells(i, 1).Value
            tickers(tickerIndex) = Cells(i, 1).Value
            startingPrices(tickerIndex) = Cells(i, 6).Value
            totalVolumes(tickerIndex) = Cells(i, 8).Value
            tickerIndex = tickerIndex + 1
            
        'tickerindex - 1 is used because tickerIndex is already incremented
        Else
            endingPrices(tickerIndex - 1) = Cells(i, 6).Value
            totalVolumes(tickerIndex - 1) = totalVolumes(tickerIndex - 1) + Cells(i, 8).Value
            
        End If
        
    Next i
    
    'Record our stock data in the analysis sheet
    Worksheets("All Stocks Analysis").Activate
        
    'Loop through our data arrays to record each value
    For i = 0 To tickerCount - 1

        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = totalVolumes(i)
        Cells(4 + i, 3).Value = endingPrices(i) / startingPrices(i) - 1
        
    Next i

    'Formatting
    Worksheets("All Stocks Analysis").Activate
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

End Sub


