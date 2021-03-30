Sub StockMarket()

'Create a script that will loop through all the stocks for one year and output the following information.

'The ticker symbol.

'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

'The total stock volume of the stock.


For Each ws In Worksheets

    'Variables needed for calculating yearly change and percent change
    Dim openingAmount As Double
    Dim closingAmount As Double
    
    Dim yearlyChange As Double

    Dim percentChange As Double

    'This variable is the counter for I through L
    Dim SummaryTableRow As Long
        SummaryTableRow = 2
        
    'This variable stores the added Stock Volume after each consecutive for loop
    Dim tickerCounter As Double
        tickerCounter = 0
    
    'For finding out the last row of column A
    Dim lastrow As Long
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    

    'This opening value is being assigned outside the forloop because its the first opening value
    openingAmount = ws.Cells(2, 3).Value

    For i = 2 To lastrow

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'Variable for last closing amount at the end of year
            closingAmount = ws.Cells(i, 6).Value
        
            'Difference between last closing amount and first opening amount
            yearlyChange = closingAmount - openingAmount
            
            'This if statement handles the error caused by dividing by 0 in the percentChange formula
            If openingAmount > 0 Then
                percentChange = yearlyChange / openingAmount
            Else
                ws.Cells(SummaryTableRow, 11).Value = 0
            End If
            
            'This openingAmount variable is for all the opening amounts at the beginning of the year but starts at AA ticker
            openingAmount = ws.Cells(i + 1, 3).Value
    
            tickerCounter = tickerCounter + ws.Cells(i, 7).Value
       
            'SummaryTableRow formula allows for loop to input a value
            ws.Cells(SummaryTableRow, 9).Value = ws.Cells(i, 1).Value
            ws.Cells(SummaryTableRow, 10).Value = yearlyChange
            ws.Cells(SummaryTableRow, 11).Value = percentChange
            ws.Cells(SummaryTableRow, 12).Value = tickerCounter
             
            'Color index for the Yearlychange
            If ws.Cells(SummaryTableRow, 10).Value < 0 Then
            
                 ws.Cells(SummaryTableRow, 10).Interior.ColorIndex = 3
            
            ElseIf ws.Cells(SummaryTableRow, 10).Value >= 0 Then
        
                 ws.Cells(SummaryTableRow, 10).Interior.ColorIndex = 4
            
            End If
            
            'Ticker count reset
            tickerCounter = 0
            
            'Counter for columns I through L
            SummaryTableRow = SummaryTableRow + 1
           
           'Counter continues for all rows when (i,1) = (i+1,1)
        Else
            tickerCounter = tickerCounter + ws.Cells(i, 7).Value
        
        End If

       

    Next i
    

' Bonus
'Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume"

    'Formula/variable for finding last row of column I
    Dim lastrowBonus As Long
        lastrowBonus = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
    'Variables
    Dim percentMax As Double
    Dim percentMin As Double
    Dim volumeTotal As Double

    'Variables with a min max function to find greatest increase/decrease
    percentMax = WorksheetFunction.Max(ws.Range("K2:K9999"))
    percentMin = WorksheetFunction.Min(ws.Range("K2:K9999"))
    volumeTotal = WorksheetFunction.Max(ws.Range("L2:L9999"))
    
    'Assign the values above into the appropriate cell
    ws.Cells(2, 17).Value = percentMax
    ws.Cells(3, 17).Value = percentMin
    ws.Cells(4, 17).Value = volumeTotal

    For x = 2 To lastrowBonus
    
    'Look for the value that matches the max
    If ws.Cells(2, 17).Value = ws.Cells(x, 11).Value Then
        ws.Cells(2, 16).Value = ws.Cells(x, 9).Value
    'Look for the value that matches the min
    ElseIf ws.Cells(3, 17).Value = ws.Cells(x, 11).Value Then
        ws.Cells(3, 16).Value = ws.Cells(x, 9).Value
    'Look for the value that matches the max volume
    ElseIf ws.Cells(4, 17).Value = ws.Cells(x, 12).Value Then
        ws.Cells(4, 16).Value = ws.Cells(x, 9).Value
    End If

    Next x
    
    
Next ws


End Sub
