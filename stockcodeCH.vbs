Sub stock_test()

    ' loop for each sheet
    For Each ws In Worksheets
    
            Dim i As Long
            Dim stock_index As Integer
            Dim total_volume As Double
            
            total_volume = 0
            stock_index = 2
            
            Dim start As Long
            start = 2
            
            ' iterate through all rows
            For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
                
                total_volume = total_volume + ws.Cells(i, 7).Value
                   
                ' check if last record for current stock
                 If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    ' if so, log out desired stock metrics into secondary table
                    ws.Cells(stock_index, 9).Value = ws.Cells(i, 1).Value  ' output current ticker
                    ws.Cells(stock_index, 10).Value = ws.Cells(i, 6).Value - ws.Cells(start, 3).Value ' yearly change
                    ws.Cells(stock_index, 11).Value = (ws.Cells(i, 6).Value - ws.Cells(start, 3).Value) / ws.Cells(start, 3).Value ' percent change
                    ws.Cells(stock_index, 12).Value = total_volume   ' total stock volume
                    
                    ' reset counters
                    total_volume = 0
                    stock_index = stock_index + 1
                    start = i + 1
                    
                End If
            
            Next i
        
        ' define lastrow of secondary table for future use
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, 11).End(xlUp).Row
        
        ' formatting adjustments
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Range("K2:K" & lastRow).NumberFormat = "0.00%"
        
        ' find and output values from secondary table
            Dim j As Integer
            Dim maxVal As Double
            Dim minVal As Double
            Dim maxVolume As Double
            
            Dim maxTicker As String
            Dim minTicker As String
            Dim volTicker As String
            
            maxVal = -1000
            minVal = 10000
            maxVolume = 0
            
            For j = 2 To lastRow
                    
                    ' greatest % inc
                    If ws.Cells(j, 11).Value > maxVal Then
                    maxVal = ws.Cells(j, 11).Value
                    maxTicker = ws.Cells(j, 9).Value
                    ws.Cells(2, 16).Value = maxTicker
                    ws.Cells(2, 17).Value = maxVal
                    End If
                    
                    ' greatest % dec
                    If ws.Cells(j, 11).Value < minVal Then
                    minVal = ws.Cells(j, 11).Value
                    minTicker = ws.Cells(j, 9).Value
                    ws.Cells(3, 16).Value = minTicker
                    ws.Cells(3, 17).Value = minVal
                    End If
                    
                    ' greatest total volume
                    If ws.Cells(j, 12).Value > maxVolume Then
                    maxVolume = ws.Cells(j, 12).Value
                    volTicker = ws.Cells(j, 9).Value
                    ws.Cells(4, 16).Value = volTicker
                    ws.Cells(4, 17).Value = maxVolume
                    End If
                    
                    ' apply green/red background to yearly changes if +/-
                    If ws.Cells(j, 10).Value > 0 Then
                    ws.Cells(j, 10).Interior.ColorIndex = 4
                    Else
                    ws.Cells(j, 10).Interior.ColorIndex = 3
                    End If
        
            Next j
            
            'formatting adjustments
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"
            ws.Range("Q2:Q3").NumberFormat = "0.00%"
            ws.Cells.EntireColumn.AutoFit
        
    Next ws
End Sub


