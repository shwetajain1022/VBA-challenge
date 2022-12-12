Attribute VB_Name = "Module1"
Sub DisplayTickerAndCalculation():
    'Delcare variables to be used
    Dim cntWS, k As Integer
    Dim openPrice, closePrice, greatestVolume, minChangePercent, maxChangePercent, totalVolume, noofrows As Double
    Dim tickerName, minChangePercentTicker, maxChangePercentTicker, greatestVolumeTicker As String
    'iterate through the worksheets
    'Question 1 - Solution
    cntWS = ActiveWorkbook.Worksheets.Count
    For i = 1 To cntWS
        noofrows = ActiveWorkbook.Worksheets(i).Cells(Rows.Count, "A").End(xlUp).Row
        k = 2
        ActiveWorkbook.Worksheets(i).Cells(1, 9).Value = "Ticker"
        ActiveWorkbook.Worksheets(i).Cells(1, 10).Value = "Yearly Change"
        ActiveWorkbook.Worksheets(i).Cells(1, 11).Value = "Percent Change"
        ActiveWorkbook.Worksheets(i).Cells(1, 12).Value = "Total Stock Volume"
        tickerName = ActiveWorkbook.Worksheets(i).Cells(2, 1).Value
        openPrice = ActiveWorkbook.Worksheets(i).Cells(2, 3).Value
        closePrice = ActiveWorkbook.Worksheets(i).Cells(2, 6).Value
        totalVolume = ActiveWorkbook.Worksheets(i).Cells(2, 7).Value
        'Change the number format to percent for K column (Percent change column)
        ActiveWorkbook.Worksheets(i).Range("K:K").NumberFormat = "0.00%"
        'Iterate through rows in the worksheet
        For j = 3 To noofrows
            'check if the previous row value for the Ticker Name is same
            If (ActiveWorkbook.Worksheets(i).Cells(j, 1).Value = ActiveWorkbook.Worksheets(i).Cells(j - 1, 1).Value) Then
                'if yes then add the volume from the row to totalvolume variable and update the closeprice with the close price from the row 
                totalVolume = totalVolume + ActiveWorkbook.Worksheets(i).Cells(j, 7).Value                
                closePrice = ActiveWorkbook.Worksheets(i).Cells(j, 6).Value
            Else
                'if no then display the calculated values using tickerName,closePrice,openPrice stored in into column I,J,K 
                ActiveWorkbook.Worksheets(i).Cells(k, 9).Value = tickerName
                ActiveWorkbook.Worksheets(i).Cells(k, 10).Value = (closePrice - openPrice)
                'Change the color of the Yearly Change cell (column J) to red if difference of  closePrice and openPrice is red else green
                If ((closePrice - openPrice) < 0) Then
                    ActiveWorkbook.Worksheets(i).Cells(k, 10).Interior.ColorIndex = 3
                Else
                    ActiveWorkbook.Worksheets(i).Cells(k, 10).Interior.ColorIndex = 4
                End If
                'display percent change of the yearly open price with yealry close price for the ticker name
                ActiveWorkbook.Worksheets(i).Cells(k, 11).Value = ActiveWorkbook.Worksheets(i).Cells(k, 10).Value / openPrice
                'display the total stock volumen for the ticker name
                ActiveWorkbook.Worksheets(i).Cells(k, 12).Value = totalVolume
                'update counter k
                k = k + 1
                'initiate values for the next ticker name
                tickerName = ActiveWorkbook.Worksheets(i).Cells(j, 1).Value
                openPrice = ActiveWorkbook.Worksheets(i).Cells(j, 3).Value
                closePrice = ActiveWorkbook.Worksheets(i).Cells(j, 6).Value
                totalVolume = ActiveWorkbook.Worksheets(i).Cells(j, 7).Value
            End If
        Next j
        'Question 2 - solution
        'Declare the columns for displaying the information
        ActiveWorkbook.Worksheets(i).Cells(1, 15).Value = "Ticker"
        ActiveWorkbook.Worksheets(i).Cells(1, 16).Value = "Value"
        ActiveWorkbook.Worksheets(i).Cells(2, 14).Value = "Greatest % increase"
        ActiveWorkbook.Worksheets(i).Cells(3, 14).Value = "Greatest % decrease"
        ActiveWorkbook.Worksheets(i).Cells(4, 14).Value = "Greatest total volume"
        'assign initial values to the variables - minChangePercent,maxChangePercent,greatestVolume,minChangePercentTicker,maxChangePercentTicker and greatestVolumeTicker
        minChangePercent = ActiveWorkbook.Worksheets(i).Cells(2, 11).Value
        maxChangePercent = ActiveWorkbook.Worksheets(i).Cells(2, 11).Value
        greatestVolume = ActiveWorkbook.Worksheets(i).Cells(2, 12).Value
        minChangePercentTicker = ActiveWorkbook.Worksheets(i).Cells(2, 9).Value
        maxChangePercentTicker = ActiveWorkbook.Worksheets(i).Cells(2, 9).Value
        greatestVolumeTicker = ActiveWorkbook.Worksheets(i).Cells(2, 9).Value
        'Change number format for the greatest minimum change percent and greatest maximum change percent columns
        ActiveWorkbook.Worksheets(i).Range("P2:P3").NumberFormat = "0.00%"
        For l = 3 To k
            'Calculation for getting the greatest minimum change percent
            If (minChangePercent > ActiveWorkbook.Worksheets(i).Cells(l, 11).Value) Then
                minChangePercent = ActiveWorkbook.Worksheets(i).Cells(l, 11).Value
                minChangePercentTicker = ActiveWorkbook.Worksheets(i).Cells(l, 9).Value
            End If
            'Calculation for getting the greatest maximum change percent
            If (maxChangePercent < ActiveWorkbook.Worksheets(i).Cells(l, 11).Value) Then
                maxChangePercent = ActiveWorkbook.Worksheets(i).Cells(l, 11).Value
                maxChangePercentTicker = ActiveWorkbook.Worksheets(i).Cells(l, 9).Value
            End If
            'Calculation for getting the greatest volume change
            If (greatestVolume < ActiveWorkbook.Worksheets(i).Cells(l, 12).Value) Then
                greatestVolume = ActiveWorkbook.Worksheets(i).Cells(l, 12).Value
                greatestVolumeTicker = ActiveWorkbook.Worksheets(i).Cells(l, 9).Value
            End If
        Next l
        'Display the Ticker names for the greatest minimum change percent, greatest maximum change percent and greatest volume change
        ActiveWorkbook.Worksheets(i).Cells(2, 15).Value = maxChangePercentTicker
        ActiveWorkbook.Worksheets(i).Cells(3, 15).Value = minChangePercentTicker
        ActiveWorkbook.Worksheets(i).Cells(4, 15).Value = greatestVolumeTicker
        'Display the values for the greatest minimum change percent, greatest maximum change percent and greatest volume change
        ActiveWorkbook.Worksheets(i).Cells(2, 16).Value = maxChangePercent
        ActiveWorkbook.Worksheets(i).Cells(3, 16).Value = minChangePercent
        ActiveWorkbook.Worksheets(i).Cells(4, 16).Value = greatestVolume
    Next i
    
    
End Sub
