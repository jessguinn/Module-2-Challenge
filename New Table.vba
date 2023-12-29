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
