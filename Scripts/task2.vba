Option Explicit

Sub CreateColumns()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim volume As Double
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double

    ' Loop through all worksheets
    For Each ws In Worksheets
        ' Initialize variables for tracking current row
        Dim i As Long
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Create columns for additional data
        ws.Cells(1, 8).Value = "Ticker Symbol"
        ws.Cells(1, 9).Value = "Total Stock Volume"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"

        ' Loop through each row in the worksheet
        For i = 2 To lastRow
            ' Retrieve data from each row
            ticker = ws.Cells(i, 1).Value ' Ticker symbol
            volume = ws.Cells(i, 7).Value ' Volume of stock
            openPrice = ws.Cells(i, 3).Value ' Open price
            closePrice = ws.Cells(i, 6).Value ' Close price
            
            ' Calculate yearly change and percent change
            yearlyChange = closePrice - openPrice
            percentChange = IIf(openPrice = 0, 0, yearlyChange / openPrice)
            
            ' Output data to the new columns
            ws.Cells(i, 8).Value = ticker ' Ticker symbol
            ws.Cells(i, 9).Value = volume ' Total stock volume
            ws.Cells(i, 10).Value = yearlyChange ' Yearly change
            ws.Cells(i, 11).Value = percentChange ' Percent change
        Next i
    Next ws
End Sub
