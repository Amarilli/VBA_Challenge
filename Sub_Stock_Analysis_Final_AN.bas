Attribute VB_Name = "Module1"
Sub StockAnalysis()

        ' Set boundaries and locations for variables
        Dim WorksheetName As String
        Dim i As Long
        Dim j As Long
        Dim TableRow As Long
        Dim LastRow As Long
        
        
        ' Define all variables, data types, and values
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim PercentMin As Double
        Dim PercentMax As Double
        Dim VolumeMax As Double
        Dim TotalStockVolume As Double
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim TickerName As String
        

        ' Loop through all worksheets
    For Each ws In Worksheets
    
        ' WorksheetName
        WorksheetName = ws.Name
        
        ' Add titles to columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ' Set starting point
        TableRow = 2
        j = 2
        
        ' Define lastrow of worksheet column A
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Do loop of the current worksheet to Lastrow
        For i = 2 To LastRow
        
            ' Ticker symbol output column nine I
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    ws.Cells(TableRow, 9).Value = ws.Cells(i, 1).Value
                    
                ' Calculate Yearly Change and save it in column ten J
                
                    YearlyChange = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                    ws.Cells(TableRow, 10).Value = YearlyChange
                ' Conditional formatting for YearlyChange: green for positive and red for negative
                    If ws.Cells(TableRow, 10).Value > 0 Then
                        ws.Cells(TableRow, 10).Interior.ColorIndex = 4
                    ElseIf ws.Cells(TableRow, 10).Value <= 0 Then
                        ws.Cells(TableRow, 10).Interior.ColorIndex = 3
                    End If
                ' Calculate Percent Change
                    If ws.Cells(j, 3).Value <> 0 Then
                        PercentChange = (YearlyChange / ws.Cells(j, 3).Value)
                 'Percent sign
                        ws.Cells(TableRow, 11).Value = Format(PercentChange, "0.00%")
                    Else
                        ws.Cells(TableRow, 11).Value = Format(0, "0.00%")
                    End If
                'Conditional formatting for PercentChange: red for negative and green for positive
                    If ws.Cells(TableRow, 11).Value > 0 Then
                        ws.Cells(TableRow, 11).Interior.ColorIndex = 4
                    ElseIf ws.Cells(TableRow, 11).Value <= 0 Then
                        ws.Cells(TableRow, 11).Interior.ColorIndex = 3
                    End If
                'Write the Total Stock Volume in the table row column twelve L
                    
    
                ws.Cells(TableRow, 12).Value = WorksheetFunction.Sum(ws.Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                ' Make the loop work by adding 1
                    
                TableRow = TableRow + 1
                j = i + 1
            End If
        Next i
        
        ' Second part - hard solution
        
        ' Define lastrow of worksheet column I
        LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        ' Extrapolate percentage and tickers
        PercentMax = ws.Cells(2, 11).Value
        PercentMin = ws.Cells(2, 11).Value
        VolumeMax = ws.Cells(2, 12).Value
        
        For i = 2 To LastRow
                
                ' Greatest increase
            If ws.Cells(i, 11).Value > PercentMax Then
                PercentMax = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            Else
                PercentMax = PercentMax
            End If
        
            ' Greatest decrease
            If ws.Cells(i, 11).Value < PercentMin Then
                PercentMin = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            Else
                PercentMin = PercentMin
            End If
               
               'TotalVolume
            If ws.Cells(i, 12).Value > TotalVolume Then
               TotalVolume = ws.Cells(i, 12).Value
               ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            Else
          
               TotalVolume = TotalVolume
               
            End If
               
                
             'Results in cells
            ws.Cells(2, 17).Value = Format(PercentMax, "0.00%")
            ws.Cells(3, 17).Value = Format(PercentMin, "0.00%")
            ws.Cells(4, 17).Value = VolumeMax
            ws.Columns("Q").EntireColumn.AutoFit
            
        
        Next i
    
    Next ws
    
    End Sub
