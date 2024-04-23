Attribute VB_Name = "Module1"
Sub stockanalysis1()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, rowOutput As Long
    Dim ticker As String
    Dim volumeTotal As Double, openPrice As Double, closePrice As Double
    Dim startYr As Double, endYr As Double, yearlyChange As Double, percentChange As Double
    Dim greatestIncrease As Double, greatestDecrease As Double, greatestVolume As Double
    Dim incTicker As String, decTicker As String, volTicker As String
    
    greatestIncrease = -1E+308  ' Initialize to very small number
    greatestDecrease = 1E+308   ' Initialize to very large number
    greatestVolume = 0

    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        rowOutput = 2 ' Start output from row 2 to avoid headers

        ' Define headers for output
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' More headers for greatest values
        ws.Cells(1, 14).Value = "Metric"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"

        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"

        openPrice = ws.Cells(2, 3).Value ' Assuming column 3 is Open Price
        volumeTotal = 0
        startYr = ws.Cells(2, 2).Value ' Assuming column 2 is Date and starts with the year

        For i = 2 To lastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                closePrice = ws.Cells(i, 6).Value ' Assuming column 6 is Close Price
                endYr = ws.Cells(i, 2).Value ' Assuming end year is on the last row of the ticker for the year
                volumeTotal = volumeTotal + ws.Cells(i, 7).Value ' Assuming column 7 is Volume

                yearlyChange = closePrice - openPrice
                If openPrice <> 0 Then
                    percentChange = yearlyChange / openPrice
                Else
                    percentChange = 0
                End If

                ' Output data to worksheet
                ws.Cells(rowOutput, 9).Value = ticker
                ws.Cells(rowOutput, 10).Value = yearlyChange
                ws.Cells(rowOutput, 11).Value = Format(percentChange, "0.00%")
                ws.Cells(rowOutput, 12).Value = volumeTotal
                
                ' Apply Conditional Formatting
                If yearlyChange > 0 Then
                    ws.Cells(rowOutput, 10).Interior.Color = RGB(0, 255, 0) ' Green
                ElseIf yearlyChange < 0 Then
                    ws.Cells(rowOutput, 10).Interior.Color = RGB(255, 0, 0) ' Red
                End If

                ' Check for greatest metrics
                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    incTicker = ticker
                End If

                If percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    decTicker = ticker
                End If

                If volumeTotal > greatestVolume Then
                    greatestVolume = volumeTotal
                    volTicker = ticker
                End If

                rowOutput = rowOutput + 1
                If i + 1 <= lastRow Then
                    openPrice = ws.Cells(i + 1, 3).Value
                    volumeTotal = 0
                End If
            Else
                volumeTotal = volumeTotal + ws.Cells(i, 7).Value
            End If
        Next i

        ' Output greatest values
        ws.Cells(2, 15).Value = incTicker
        ws.Cells(2, 16).Value = Format(greatestIncrease, "0.00%")
        ws.Cells(3, 15).Value = decTicker
        ws.Cells(3, 16).Value = Format(greatestDecrease, "0.00%")
        ws.Cells(4, 15).Value = volTicker
        ws.Cells(4, 16).Value = greatestVolume
    Next ws

    
End Sub


