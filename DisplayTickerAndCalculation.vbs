Attribute VB_Name = "Module1"
Sub DisplayTickerAndCalculation():
    Dim cntWS, k As Integer
    Dim openPrice, closePrice, greatestVolume, minChangePercent, maxChangePercent, totalVolume, noofrows As Double
    Dim tickerName, minChangePercentTicker, maxChangePercentTicker, greatestVolumeTicker As String
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
        ActiveWorkbook.Worksheets(i).Range("K:K").NumberFormat = "0.00%"
        For j = 3 To noofrows
            If (ActiveWorkbook.Worksheets(i).Cells(j, 1).Value = ActiveWorkbook.Worksheets(i).Cells(j - 1, 1).Value) Then
                totalVolume = totalVolume + ActiveWorkbook.Worksheets(i).Cells(j, 7).Value
                closePrice = ActiveWorkbook.Worksheets(i).Cells(j, 6).Value
            Else
                ActiveWorkbook.Worksheets(i).Cells(k, 9).Value = tickerName
                ActiveWorkbook.Worksheets(i).Cells(k, 10).Value = (closePrice - openPrice)
                If ((closePrice - openPrice) < 0) Then
                    ActiveWorkbook.Worksheets(i).Cells(k, 10).Interior.ColorIndex = 3
                Else
                    ActiveWorkbook.Worksheets(i).Cells(k, 10).Interior.ColorIndex = 4
                End If
                ActiveWorkbook.Worksheets(i).Cells(k, 11).Value = ActiveWorkbook.Worksheets(i).Cells(k, 10).Value / openPrice
                ActiveWorkbook.Worksheets(i).Cells(k, 12).Value = totalVolume
                k = k + 1
                tickerName = ActiveWorkbook.Worksheets(i).Cells(j, 1).Value
                openPrice = ActiveWorkbook.Worksheets(i).Cells(j, 3).Value
                closePrice = ActiveWorkbook.Worksheets(i).Cells(j, 6).Value
                totalVolume = ActiveWorkbook.Worksheets(i).Cells(j, 7).Value
            End If
        Next j
        ActiveWorkbook.Worksheets(i).Cells(1, 15).Value = "Ticker"
        ActiveWorkbook.Worksheets(i).Cells(1, 16).Value = "Value"
        ActiveWorkbook.Worksheets(i).Cells(2, 14).Value = "Greatest % increase"
        ActiveWorkbook.Worksheets(i).Cells(3, 14).Value = "Greatest % decrease"
        ActiveWorkbook.Worksheets(i).Cells(4, 14).Value = "Greatest total volume"
        minChangePercent = ActiveWorkbook.Worksheets(i).Cells(2, 11).Value
        maxChangePercent = ActiveWorkbook.Worksheets(i).Cells(2, 11).Value
        greatestVolume = ActiveWorkbook.Worksheets(i).Cells(2, 12).Value
        minChangePercentTicker = ActiveWorkbook.Worksheets(i).Cells(2, 9).Value
        maxChangePercentTicker = ActiveWorkbook.Worksheets(i).Cells(2, 9).Value
        greatestVolumeTicker = ActiveWorkbook.Worksheets(i).Cells(2, 9).Value
        ActiveWorkbook.Worksheets(i).Range("P2:P3").NumberFormat = "0.00%"
        For l = 3 To k
            If (minChangePercent > ActiveWorkbook.Worksheets(i).Cells(l, 11).Value) Then
                minChangePercent = ActiveWorkbook.Worksheets(i).Cells(l, 11).Value
                minChangePercentTicker = ActiveWorkbook.Worksheets(i).Cells(l, 9).Value
            End If
            If (maxChangePercent < ActiveWorkbook.Worksheets(i).Cells(l, 11).Value) Then
                maxChangePercent = ActiveWorkbook.Worksheets(i).Cells(l, 11).Value
                maxChangePercentTicker = ActiveWorkbook.Worksheets(i).Cells(l, 9).Value
            End If
            If (greatestVolume < ActiveWorkbook.Worksheets(i).Cells(l, 12).Value) Then
                greatestVolume = ActiveWorkbook.Worksheets(i).Cells(l, 12).Value
                greatestVolumeTicker = ActiveWorkbook.Worksheets(i).Cells(l, 9).Value
            End If
        Next l
        ActiveWorkbook.Worksheets(i).Cells(2, 15).Value = maxChangePercentTicker
        ActiveWorkbook.Worksheets(i).Cells(3, 15).Value = minChangePercentTicker
        ActiveWorkbook.Worksheets(i).Cells(4, 15).Value = greatestVolumeTicker
        
        ActiveWorkbook.Worksheets(i).Cells(2, 16).Value = maxChangePercent
        ActiveWorkbook.Worksheets(i).Cells(3, 16).Value = minChangePercent
        ActiveWorkbook.Worksheets(i).Cells(4, 16).Value = greatestVolume
    Next i
    
    
End Sub
