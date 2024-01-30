Option Explicit

Sub StockAnalysisForAllSheets()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim volume As Double
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryTableRow As Integer

    ' Loop through all worksheets
    For Each ws In Worksheets
        ' Initialize variables for summary table
        summaryTableRow = 2
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Initialize variables for tracking greatest % increase, % decrease, and total volume
        Dim greatestIncrease As Double
        Dim greatestDecrease As Double
        Dim greatestVolume As Double
        Dim greatestIncreaseTicker As String
        Dim greatestDecreaseTicker As String
        Dim greatestVolumeTicker As String
        
        ' Initialize variables for tracking current row
        Dim i As Long
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Loop through each row in the worksheet
        For i = 2 To lastRow
            ' Check if the current row's ticker is different from the previous row
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                ' If it's a new ticker, set opening price
                openPrice = ws.Cells(i, 3).Value
            End If
            
            ' Add the stock volume to the total
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            ' Check if the current row's ticker is different from the next row or it's the last row
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Or i = lastRow Then
                ' If it's the last row for the current ticker, set closing price
                closePrice = ws.Cells(i, 6).Value
                
                ' Calculate yearly change and percent change
                yearlyChange = closePrice - openPrice
                percentChange = IIf(openPrice = 0, 0, yearlyChange / openPrice)
                
                ' Output results to summary table
                ws.Cells(summaryTableRow, 9).Value = ws.Cells(i, 1).Value ' Ticker
                ws.Cells(summaryTableRow, 10).Value = yearlyChange ' Yearly Change
                ws.Cells(summaryTableRow, 11).Value = percentChange ' Percent Change
                ws.Cells(summaryTableRow, 12).Value = totalVolume ' Total Stock Volume
                
                ' Apply conditional formatting based on positive/negative yearly change
                If yearlyChange >= 0 Then
                    ws.Cells(summaryTableRow, 10).Interior.Color = RGB(0, 255, 0) ' Green for positive change
                Else
                    ws.Cells(summaryTableRow, 10).Interior.Color = RGB(255, 0, 0) ' Red for negative change
                End If
                
                ' Update tracking variables for greatest % increase, % decrease, and total volume
                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    greatestIncreaseTicker = ws.Cells(i, 1).Value
                End If
                
                If percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    greatestDecreaseTicker = ws.Cells(i, 1).Value
                End If
                
                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    greatestVolumeTicker = ws.Cells(i, 1).Value
                End If
                
                ' Reset variables for the next ticker
                totalVolume = 0
                summaryTableRow = summaryTableRow + 1
            End If
        Next i
        
        ' Output the greatest % increase, % decrease, and total volume to the worksheet
        ws.Cells(2, 16).Value = "Greatest % Increase"
        ws.Cells(3, 16).Value = "Greatest % Decrease"
        ws.Cells(4, 16).Value = "Greatest Total Volume"
        
        ws.Cells(2, 17).Value = greatestIncreaseTicker
        ws.Cells(3, 17).Value = greatestDecreaseTicker
        ws.Cells(4, 17).Value = greatestVolumeTicker
        
        ws.Cells(2, 18).Value = greatestIncrease * 100 & "%"
        ws.Cells(3, 18).Value = greatestDecrease * 100 & "%"
        ws.Cells(4, 18).Value = greatestVolume
    Next ws
End Sub
