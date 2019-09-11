
Sub StockSummary():
    
    Dim ws As Worksheet
    Dim i As Long
    Dim RowNum As Long
    Dim Vol As Double
    Dim Summary_Table As Integer

    Summary_Table = 2

    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim PctChg As Double
    Dim FirstTime As Integer

    FirstTime = 0

    Dim YrChg As Double

    For Each ws In Worksheets

        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"

        RowNum = ws.Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To RowNum
            If ws.Cells(i, 1) = ws.Cells(i, 1).Offset(1, 0) Then
                FirstTime = FirstTime + 1
                Vol = Vol + ws.Cells(i, 7)
                If FirstTime = 1 Then
                    OpenPrice = ws.Cells(i, 3)
                Else
                End If
            Else

                Vol = Vol + ws.Cells(i, 7)

                ws.Cells(Summary_Table, 9) = ws.Cells(i, 1)
                ws.Cells(Summary_Table, 12) = Vol
                
                ClosePrice = ws.Cells(i, 6)

                If OpenPrice <> 0 Then
                    PctChg = ((ClosePrice - OpenPrice) / OpenPrice)
                    YrChg = ClosePrice - OpenPrice
                Else
                    PctChg = 0
                    YrChg = 0
                End If

                ws.Cells(Summary_Table, 11) = PctChg
                ws.Cells(Summary_Table, 11).NumberFormat = "0.00%"
                ws.Cells(Summary_Table, 10) = YrChg
                    If ws.Cells(Summary_Table, 10).Value > 0 Then
                        ws.Cells(Summary_Table, 10).Interior.ColorIndex = 4
                    Else
                        ws.Cells(Summary_Table, 10).Interior.ColorIndex = 3
                    End If

                Vol = 0

                Summary_Table = Summary_Table + 1

                FirstTime = 0
            End If
        Next i

    Summary_Table = 2

    Next ws
End Sub

