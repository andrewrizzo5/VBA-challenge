Sub stock_analysis()
    Dim ws As Worksheet
    Dim row_count As Long
    Dim i As Long
    Dim output_row As Long
    Dim stock_ticker As String
    Dim opening_price As Double
    Dim closing_price As Double
    Dim change_over_year As Double
    Dim percent_change As Double
    Dim volume_total As Double
    Dim max_increase As Double
    Dim max_decrease As Double
    Dim max_volume As Double
    Dim max_increase_ticker As String
    Dim max_decrease_ticker As String
    Dim max_volume_ticker As String

    For Each ws In ThisWorkbook.Worksheets
        output_row = 2
        volume_total = 0
        opening_price = ws.Cells(2, 3).Value
        max_increase = 0
        max_decrease = 0
        max_volume = 0

        row_count = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        For i = 2 To row_count
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                stock_ticker = ws.Cells(i, 1).Value
                volume_total = volume_total + ws.Cells(i, 7).Value
                closing_price = ws.Cells(i, 6).Value

                change_over_year = closing_price - opening_price
                If opening_price <> 0 Then
                    percent_change = (change_over_year / opening_price) * 100
                Else
                    percent_change = 0
                End If

                If percent_change > max_increase Then
                    max_increase = percent_change
                    max_increase_ticker = stock_ticker
                End If

                If percent_change < max_decrease Then
                    max_decrease = percent_change
                    max_decrease_ticker = stock_ticker
                End If

                If volume_total > max_volume Then
                    max_volume = volume_total
                    max_volume_ticker = stock_ticker
                End If

                ws.Cells(output_row, 9).Value = stock_ticker
                ws.Cells(output_row, 10).Value = change_over_year
                ws.Cells(output_row, 11).Value = percent_change
                ws.Cells(output_row, 12).Value = volume_total

                output_row = output_row + 1
                volume_total = 0

                If i + 1 <= row_count Then
                    opening_price = ws.Cells(i + 1, 3).Value
                End If
            Else
                volume_total = volume_total + ws.Cells(i, 7).Value
            End If
        Next i

        ' GREATEST SECTION '
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(2, 16).Value = max_increase_ticker
        ws.Cells(3, 16).Value = max_decrease_ticker
        ws.Cells(4, 16).Value = max_volume_ticker
        ws.Cells(2, 17).Value = max_increase
        ws.Cells(3, 17).Value = max_decrease
        ws.Cells(4, 17).Value = max_volume
    Next ws
End Sub
