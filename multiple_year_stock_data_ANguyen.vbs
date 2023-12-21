Attribute VB_Name = "Module1"
Sub StockData()

    'source: https://support.microsoft.com/en-au/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0
    'source: https://stackoverflow.com/questions/21918166/excel-vba-for-each-worksheet-loop
    For Each ws In Worksheets
    
        'establish variables
        Dim Ticker As String
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim TotalVolume As Double
        Dim SummaryTableRow As Integer
    
        'print categories
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
    
        'set intial value
        TotalVolume = 0
        
        'keep track of location in summary table
        SummaryTableRow = 2
    
        'set open price
        OpenPrice = ws.Cells(2, 3).Value
    
        'source: https://www.excelcampus.com/vba/find-last-row-column-cell/
        'counts the rows in first column
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'create a loop, similar to credit charges activity
        For i = 2 To lastrow
        
            'check to see if we are within the same ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'set the ticker name
                Ticker = ws.Cells(i, 1).Value
            
                'calculate total volume
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            
                'set close price value
                ClosePrice = ws.Cells(i, 6).Value
            
                'calculate yearly change
                YearlyChange = ClosePrice - OpenPrice

                    'calculate percent change
                    If OpenPrice = 0 Then
                        PercentChange = 0
                    Else
                        PercentChange = YearlyChange / OpenPrice
                    End If
                
                'print in summary table
                ws.Range("I" & SummaryTableRow).Value = Ticker
                ws.Range("L" & SummaryTableRow).Value = TotalVolume
                ws.Range("J" & SummaryTableRow).Value = YearlyChange
                ws.Range("K" & SummaryTableRow).Value = PercentChange
            
                'source: https://www.automateexcel.com/vba/format-numbers/
                'format to percent
                ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
           
                'add one to the summary table to move down rows
                SummaryTableRow = SummaryTableRow + 1
            
                'reset total volume
                TotalVolume = 0
            
                'reset open price
                OpenPrice = ws.Cells(i + 1, 3)
        
            Else
                'add the total volume
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            End If
            
        Next i
    
        'source: https://www.excelcampus.com/vba/find-last-row-column-cell/
        'find last row of the summary table
        lastrowSummary = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
        'change colors of yearly change
        'source formater and grader solution in class activity
        For i = 2 To lastrowSummary
    
            'if values over 0 set to green
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            
            'if values negative set to red
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
        
        Next i
    
    
        'set new categories using string
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
    
        'create another loop for summary table
        For i = 2 To lastrowSummary
    
        'source: https://www.excelanytime.com/excel/index.php?option=com_content&view=article&id=105:find-smallest-and-largest-value-in-range-with-vba-excel&catid=79&Itemid=475
        'find the max
            If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrowSummary)) Then
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
                'source: https://www.automateexcel.com/vba/format-numbers/
                ws.Cells(2, 17).NumberFormat = "0.00%"
        
            'find the min
            ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrowSummary)) Then
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            
            'find the max
            ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrowSummary)) Then
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
                ws.Cells(4, 17).NumberFormat = "0.00E+0"
            End If
            
        Next i
    Next ws
End Sub
