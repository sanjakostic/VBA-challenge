Sub CalculateSummary()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim summaryRow As Long
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim openPrice As Double
    Dim closePrice As Double
    Dim totalVolume As Double
    Dim ticker As String
    
    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        summaryRow = 2
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Loop through each row to calculate yearly change, percent change, and total volume
        For i = 2 To lastRow
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                openPrice = ws.Cells(i, 3).Value
                closePrice = ws.Cells(i, 6).Value
                yearlyChange = closePrice - openPrice
                percentChange = yearlyChange / openPrice
                totalVolume = Application.Sum(ws.Range(ws.Cells(summaryRow, 7), ws.Cells(i, 7)))
                
                ws.Cells(summaryRow, 9).Value = ticker
                ws.Cells(summaryRow, 10).Value = yearlyChange
                ws.Cells(summaryRow, 11).Value = percentChange
                ws.Cells(summaryRow, 12).Value = totalVolume
                
                ' Format Yearly Change cells
                If yearlyChange >= 0 Then
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(0, 255, 0) ' Green
                Else
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(255, 0, 0) ' Red
                End If
                
                summaryRow = summaryRow + 1
            End If
        Next i
        
        ' Find Greatest % Increase, Greatest % Decrease, and Greatest Total Volume
        ws.Cells(2, 15).Value = "Ticker"
        ws.Cells(2, 16).Value = "Value"
        ws.Cells(3, 14).Value = "Greatest % Increase"
        ws.Cells(4, 14).Value = "Greatest % Decrease"
        ws.Cells(5, 14).Value = "Greatest Total Volume"
        
        ws.Cells(3, 16).Value = WorksheetFunction.Max(ws.Range("K:K"))
        ws.Cells(4, 16).Value = WorksheetFunction.Min(ws.Range("K:K"))
        ws.Cells(5, 16).Value = WorksheetFunction.Max(ws.Range("L:L"))
        
        ws.Cells(3, 15).Value = ws.Cells(Application.Match(ws.Cells(3, 16).Value, ws.Range("K:K"), 0) + 1, 9).Value
        ws.Cells(4, 15).Value = ws.Cells(Application.Match(ws.Cells(4, 16).Value, ws.Range("K:K"), 0) + 1, 9).Value
        ws.Cells(5, 15).Value = ws.Cells(Application.Match(ws.Cells(5, 16).Value, ws.Range("L:L"), 0) + 1, 9).Value
    Next ws
End Sub


