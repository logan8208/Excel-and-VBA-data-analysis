Sub StockDataChecker()


Dim ws as Worksheet 
Dim StockName as String
Dim StockNameTotal as Double
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

Dim lastrow as Long


For Each ws in Worksheets
    

    ws.Cells(1,9).Value = "Ticker"
    ws.Cells(1,10).Value = "Tot. Volume"

    lastrow = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 to lastrow

        If ws.Cells(i,1).Value <> ws.Cells(i+1,1).Value Then

            StockName = ws.Cells(i,1).Value

            StockNameTotal = StockNameTotal + ws.Cells(i,6).Value

            ws.Range("I" & Summary_Table_Row).Value = StockName
            ws.Range("J" & Summary_Table_Row).Value = StockNameTotal

            Summary_Table_Row = Summary_Table_Row + 1

            StockNameTotal = 0

        Else

            StockNameTotal = StockNameTotal + ws.Cells(i,6).Value


        End If

    Next i

    Summary_Table_Row = 2


Next ws


End Sub