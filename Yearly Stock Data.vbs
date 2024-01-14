Sub YearlyStockData()

'Loop through worksheets
For Each ws In Worksheets

    'Determine the Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Create a variable to hold the ticker value
    Dim ticker As String
    'Create a variable to hold the total_volume value
    Dim total_volume As Double
    'Create a variable to hold the summary_table value
    Dim summary_table As Integer
    'Create a variable to hold the yearly change
    Dim yearlyChange As Double
    'Create a variable to hold the opening price value
    Dim openPrice As Double
    'Create a variable to hold the closing price value
    Dim endPrice As Double
    'Create a variable to hold the yearly percentage change value
    Dim percentageChange As Double

    'Assign the variables total_volume, summary_table & Opening Price to a value
    total_volume = 0
    summary_table = 2
    yearlyChange = 0
    openPrice = ws.Cells(2, 3).Value

'Create a loop to go through the tables to the Last Row
For i = 2 To LastRow

'Condition to check if the value has changed
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'Assign the closing price value
            endPrice = ws.Cells(i, 6).Value
            
            'make the math to calculate the yearly change & yearly percenatge change
            yearlyChange = endPrice - openPrice
            percentageChange = (yearlyChange / openPrice) * 100
            
            'Print the values into cells
            ws.Range("J" & summary_table).Value = yearlyChange
            ws.Range("K" & summary_table).Value = Format(percentageChange, "0.00") & "%"
            
            'Setting the opening price value to the next ticker
            openPrice = ws.Cells(i + 1, 3).Value
            
            'Assign the value we got from the loop to the values
            ticker = ws.Cells(i, 1).Value
            total_volume = total_volume + ws.Cells(i, 7).Value
            
            'Name the first row of the summary table
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"
            
            'Insert the name of the value into the first row of the summary table
            ws.Range("I" & summary_table).Value = ticker
            ws.Range("L" & summary_table).Value = total_volume
            
            
            
            'Add one to the summary_table
            summary_table = summary_table + 1
            
            'Reset the total_volume
            total_volume = 0
            yearlyChange = 0
            
            
        
        Else
              total_volume = total_volume + ws.Cells(i, 7).Value
     End If
                 
  Next i



    'Create a variable to hold the first value of the row value
    Dim CurrentVolume As Double
    'Create a variable to hold the max value
    Dim maxVolume As Double
    ''Assign the the higher value
    maxVolume = ws.Cells(2, 12).Value
    'Create a variable to hold the ticker value
    Dim currenTicker As String
    'Create a variable to hold the ticker for the higher volume value
    Dim maxTickerV As String
    ''Assign the first value's ticker
    maxTickerV = ws.Cells(2, 12).Value
    
    'Create a variable to hold the first value of the percentage change value
    Dim currentGreater As Double
    'Create a variable to hold the greatest increase value
    Dim maxGreater As Double
    ' Assign the first the percentage value as the greatest
    maxGreater = ws.Cells(2, 11).Value
    'Create a variable to hold the percentage change ticker value
    Dim percentageTickerIncrease As String
    'Assign the first ticker value as the greatest increase
    percentageTickerIncrease = ws.Cells(2, 9).Value
    
    'Create a variable to hold the first value of the percentage change value
    Dim currentDecrease As Double
    'Create a variable to hold the greatest decrease value
    Dim maxDecrease As Double
    ' Assign the first the percentage value as the greatest decrease
    maxDecrease = ws.Cells(2, 11).Value
    'Create a variable to hold the percentage change ticker value
    Dim percentageTickerDecrease As String
    'Assign the first ticker value as the greatest decrease
    percentageTickerDecrease = ws.Cells(2, 9).Value

    'Assign the first ticker as the current value
    currenTicker = ws.Cells(2, 9).Value

'Create a loop to go through the tables to the Last Row
    For i = 2 To LastRow
    
        'Condition to check if the value is bigger than 0, and if it's true We want to color the cell with Green,
        'if it's less than 0, We want to color it with Red, and if it's = 0, we will color it Grey
        If (ws.Cells(i, 11).Value > 0) Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
                ElseIf ws.Cells(i, 11).Value < 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 3
                Else
                ws.Cells(i, 10).Interior.ColorIndex = 15
        End If
            
        ' Assign the first volume value as the current value
         CurrentVolume = ws.Cells(i, 12).Value
        
        'Condition to check if the current value is greater than the next value, if so, then
        'we will hold it as the max value untill we find the greatest volume
        If (CurrentVolume > maxVolume) Then
                maxVolume = CurrentVolume
                maxTickerV = ws.Cells(i, 9).Value
        End If
        
        ' Assign the first volume value as the current value
        currentGreater = ws.Cells(i, 11).Value
        
        'Condition to check if the current value is greater than the next value, if so, then
        'we will hold it as the max value untill we find the greatest percentage increase
        If (currentGreater > maxGreater) Then
                maxGreater = currentGreater
                percentageTickerIncrease = ws.Cells(i, 9).Value
        End If
        
        ' Assign the first volume value as the current value
        currentDecrease = ws.Cells(i, 11).Value
        
        'Condition to check if the current value is greater than the next value, if so, then
        'we will hold it as the max value untill we find the greatest percentage decrease
        If (currentDecrease < maxDecrease) Then
                maxDecrease = currentDecrease
                percentageTickerDecrease = ws.Cells(i, 9).Value
        End If
        
    'Printing all the results after finding the values.
    
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(2, 16).Value = percentageTickerIncrease
    ws.Cells(2, 17).Value = Format(maxGreater, "0.00%")
                
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(3, 16).Value = percentageTickerDecrease
    ws.Cells(3, 17).Value = Format(maxDecrease, "0.00%")
    
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(4, 16).Value = maxTickerV
    ws.Cells(4, 17).Value = maxVolume

    Next i

 Next ws

End Sub




