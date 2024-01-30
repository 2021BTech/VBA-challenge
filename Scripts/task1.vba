Option Explicit

Sub RetrieveStockData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim volume As Double
    Dim openPrice As Double
    Dim closePrice As Double

    ' Loop through all worksheets
    For Each ws In Worksheets
        ' Initialize variables for tracking current row
        Dim i As Long
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Loop through each row in the worksheet
        For i = 2 To lastRow
            ' Retrieve data from each row
            ticker = ws.Cells(i, 1).Value ' Ticker symbol
            volume = ws.Cells(i, 7).Value ' Volume of stock
            openPrice = ws.Cells(i, 3).Value ' Open price
            closePrice = ws.Cells(i, 6).Value ' Close price
            
            ' You can do something with the retrieved data, e.g., store in arrays, output to another sheet, etc.
            ' For now, let's just print the values in the Immediate Window for demonstration purposes
            Debug.Print "Ticker: " & ticker & ", Volume: " & volume & ", Open Price: " & openPrice & ", Close Price: " & closePrice
        Next i
    Next ws
End Sub
