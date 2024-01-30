Option Explicit

Sub ApplyConditionalFormatting()
    Dim ws As Worksheet
    Dim lastRow As Long

    ' Loop through all worksheets
    For Each ws In Worksheets
        ' Initialize variables for tracking current row
        Dim i As Long
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Apply conditional formatting to "Yearly Change" column (Column 10)
        ws.Range(ws.Cells(2, 10), ws.Cells(lastRow, 10)).FormatConditions.Delete ' Clear existing formatting
        With ws.Range(ws.Cells(2, 10), ws.Cells(lastRow, 10)).FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0")
            .Interior.Color = RGB(0, 255, 0) ' Green for positive change
        End With
        With ws.Range(ws.Cells(2, 10), ws.Cells(lastRow, 10)).FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
            .Interior.Color = RGB(255, 0, 0) ' Red for negative change
        End With

        ' Apply conditional formatting to "Percent Change" column (Column 11)
        ws.Range(ws.Cells(2, 11), ws.Cells(lastRow, 11)).FormatConditions.Delete ' Clear existing formatting
        With ws.Range(ws.Cells(2, 11), ws.Cells(lastRow, 11)).FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0")
            .Interior.Color = RGB(0, 255, 0) ' Green for positive change
        End With
        With ws.Range(ws.Cells(2, 11), ws.Cells(lastRow, 11)).FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
            .Interior.Color = RGB(255, 0, 0) ' Red for negative change
        End With

    Next ws
End Sub
