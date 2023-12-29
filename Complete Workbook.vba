Sub Yearly_Change_Workbook()
    Dim ws As Worksheet
    Dim ticker As String
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Stock_Volume As Double
    Dim i As Long
    Dim lastrow As Long
    Dim Summary_Table_Row As Integer
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Price_Row As Long
    
    For Each ws In ThisWorkbook.Worksheets
        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Summary_Table_Row = 2
        Price_Row = 2

        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        For i = 2 To lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

                ws.Cells(Summary_Table_Row, 9).Value = ticker
                ws.Cells(Summary_Table_Row, 12).Value = Total_Stock_Volume

                Open_Price = ws.Cells(Price_Row, 3).Value
                Close_Price = ws.Cells(i, 6).Value
                Yearly_Change = Close_Price - Open_Price

                If Open_Price = 0 Then
                    Percent_Change = 0
                Else
                    Percent_Change = Yearly_Change / Open_Price
                End If

                ws.Cells(Summary_Table_Row, 10).Value = Yearly_Change
                ws.Cells(Summary_Table_Row, 11).Value = Percent_Change
                ws.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"

                If Yearly_Change > 0 Then
                    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
                End If

                Summary_Table_Row = Summary_Table_Row + 1
                Price_Row = i + 1
                Total_Stock_Volume = 0
            Else
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
            End If
        Next i
        
        Dim maxPercentIncrease As Double
        Dim maxPercentChangeTicker As String
        Dim minPercentDecrease As Double
        Dim minPercentChangeTicker As String
        Dim maxTotalVolume As Double
        Dim maxTotalVolumeTicker As String
        
        maxPercentIncrease = Application.WorksheetFunction.Max(ws.Range(ws.Cells(2, 11), ws.Cells(Summary_Table_Row - 1, 11)).Value)
        minPercentDecrease = Application.WorksheetFunction.min(ws.Range(ws.Cells(2, 11), ws.Cells(Summary_Table_Row - 1, 11)).Value)
        maxTotalVolume = Application.WorksheetFunction.Max(ws.Range(ws.Cells(2, 12), ws.Cells(Summary_Table_Row - 1, 12)).Value)
        
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"

        
        For i = 2 To Summary_Table_Row - 1
            If ws.Cells(i, 11).Value = maxPercentIncrease Then
            maxPercentChangeTicker = ws.Cells(i, 9).Value
            Exit For
            End If
            
        Next i
        
        For i = 2 To Summary_Table_Row - 1
            If ws.Cells(i, 11).Value = minPercentDecrease Then
            minPercentChangeTicker = ws.Cells(i, 9).Value
            Exit For
            End If
            
        Next i
        
        For i = 2 To Summary_Table_Row - 1
        If ws.Cells(i, 12).Value = maxTotalVolume Then
        maxTotalVolumeTicker = ws.Cells(i, 9).Value
        Exit For
        End If
    Next i
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 16).Value = maxPercentChangeTicker
        ws.Cells(2, 17).Value = maxPercentIncrease
        ws.Cells(3, 16).Value = minPercentChangeTicker
        ws.Cells(3, 17).Value = minPercentDecrease
        ws.Cells(2, 15).Value = "Greatest % increase"
        ws.Cells(3, 15).Value = "Greatest % decrease"
        ws.Cells(4, 15).Value = "Greatest total volume"
        ws.Cells(4, 16).Value = maxTotalVolumeTicker
        ws.Cells(4, 17).Value = maxTotalVolume
        
    Next ws
End Sub

