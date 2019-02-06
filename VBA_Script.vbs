Sub StockMarket():

    Dim ws As Worksheet

    For Each ws In Worksheets

        ' find the last row of each
        Dim lRow As Double
            
        lRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' set dimensions
        Dim ticker As String
        Dim total_volume As Double
        total_volume = 0

        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2

        ' set loop instructions
        For i = 2 To lRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = (ws.Cells(i, 1).Value)
                total_volume = total_volume + ws.Cells(i, 7).Value
                ws.Range("J" & Summary_Table_Row).Value = ticker
                ws.Range("K" & Summary_Table_Row).Value = total_volume
                Summary_Table_Row = Summary_Table_Row + 1
                total_volume = 0
            Else
                total_volume = total_volume + Cells(i, 7)
            End If
        Next i
       
    Next ws

End Sub


