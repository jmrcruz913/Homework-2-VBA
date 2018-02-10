Sub Homeworktwo()

    
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
         
    Dim ticker As String

    Dim total_stock_volume As Double
    total_stock_volume = 0

    Dim Summary_Table_Row As Integer

    Summary_Table_Row = 2
    
    Dim LastRow As Long
    
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To LastRow


        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        ticker = Cells(i, 1).Value

        total_stock_volume = total_stock_volume + Cells(i, 7).Value


            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Total Stock Volume"


            ws.Range("I" & Summary_Table_Row).Value = ticker

            ws.Range("J" & Summary_Table_Row).Value = total_stock_volume

        Summary_Table_Row = Summary_Table_Row + 1

        total_stock_volume = 0

    Else

    total_stock_volume = total_stock_volume + Cells(i, 7).Value

End If

Next i

         
    Next ws

End Sub
