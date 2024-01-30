Option Explicit

Sub CalculateGreatestValues()
    Dim ws As Worksheet
    Dim lastRow As Long

    ' Loop through all worksheets
    For Each ws In Worksheets
        ' Initialize variables for tracking current row
        Dim i As Long
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Initialize variables for tracking greatest % increase, % decrease, and total volume
        Dim greatestIncrease As Double
        Dim greatestDecrease As Double
        Dim greatestVolume As Double
        Dim greatestIncreaseTicker As String
        Dim greatestDecreaseTicker As String
        Dim greatestVolumeTicker As String
        Dim yearlyChange As Double
        Dim percentChange As Double
        Dim totalVolume As Double

        ' Loop through each row in the worksheet
        For i = 2 To lastRow
            ' Retrieve data from each row
            yearlyChange = ws.Cells(i, 10).Value ' Yearly change
            percentChange = ws.Cells(i, 11).Value ' Percent change
            totalVolume = ws.Cells(i, 12).Value ' Total stock volume

            ' Update tracking variables for greatest % increase, % decrease, and total volume
            If i = 2 Then
                greatestIncrease = percentChange
                greatestDecrease = percentChange
                greatestVolume = totalVolume
                greatestIncreaseTicker = ws.Cells(i, 9).Value ' Ticker symbol
                greatestDecreaseTicker = ws.Cells(i, 9).Value ' Ticker symbol
                greatestVolumeTicker = ws.Cells(i, 9).Value ' Ticker symbol
            End If

            If percentChange > greatestIncrease Then
                greatestIncrease = percentChange
                greatestIncreaseTicker = ws.Cells(i, 9).Value ' Ticker symbol
            End If

            If percentChange < greatestDecrease Then
                greatestDecrease = percentChange
                greatestDecreaseTicker = ws.Cells(i, 9).Value ' Ticker symbol
            End If

            If totalVolume > greatestVolume Then
                greatestVolume = totalVolume
                greatestVolumeTicker = ws.Cells(i, 9).Value ' Ticker symbol
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
