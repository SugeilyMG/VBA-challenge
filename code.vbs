Sub Summary():

On Error Resume Next

Dim ws_count As Integer
Dim ws As Worksheet
Dim r As Range
Dim ticker As String
Dim ticker_count As Integer

For Each ws In ActiveWorkbook.Worksheets
    max_value = 0
    total_value = 0
    min_value = 0
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
        
    last_r = Cells(Rows.Count, "A").End(xlUp).Row
    j = 2
    Total = 0
    open_value = ws.Cells(2, 3).Value
    For i = 2 To last_r
        ticker = ws.Cells(i, 1).Value
        Total = ws.Cells(i, 7).Value + Total
        If ticker <> ws.Cells(i + 1, 1).Value Then
            ws.Cells(j, 9).Value = ticker
            ws.Cells(j, 12).Value = Total
            close_value = ws.Cells(i, 6).Value
            y_change = close_value - open_value
            p_change = Round(((close_value - open_value) / open_value) * 100, 2)
            ws.Cells(j, 10).Value = y_change
            ws.Cells(j, 11).Value = p_change
                If y_change > 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4
                Else
                ws.Cells(j, 10).Interior.ColorIndex = 3
                End If
            j = j + 1
            ticker_count = ticker_count + 1
            Total = 0
            open_value = ws.Cells(i + 1, 3).Value
        End If
    Next i
        
        last_row = Cells(Rows.Count, "I").End(xlUp).Row
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ws.Cells(2, 15).Value = "Greatest % increase"
        max_value = WorksheetFunction.Max(ws.Range(Cells(2, 11), Cells(last_row, 11)))
        max_address = WorksheetFunction.Match(max_value, ws.Range(Cells(2, 11), Cells(last_row, 11)), 0)
        ws.Cells(2, 16).Value = ws.Cells(max_address + 1, 9)
        ws.Cells(2, 17).Value = max_value

        ws.Cells(4, 15).Value = "Greatest total volume"
        total_value = WorksheetFunction.Max(ws.Range(Cells(2, 12), Cells(last_row, 12)))
        total_address = WorksheetFunction.Match(max_value, ws.Range(Cells(2, 12), Cells(last_row, 12)), 0)
        ws.Cells(4, 16).Value = ws.Cells(total_address + 1, 9)
        ws.Cells(4, 17).Value = total_value
        
        ws.Cells(3, 15).Value = "Greatest % decrease"
        min_value = WorksheetFunction.Min(ws.Range(Cells(2, 11), Cells(last_row, 11)))
        min_address = WorksheetFunction.Match(min_value, ws.Range(Cells(2, 11), Cells(last_row, 11)), 0)
        ws.Cells(3, 16).Value = ws.Cells(min_address + 1, 9)
        ws.Cells(3, 17).Value = min_value
    

Next
End Sub