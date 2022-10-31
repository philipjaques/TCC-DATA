Sub StockAnalysis():
    
    For Each ws In Worksheets
    
        ' Stock Vars
        Dim StockTicker As String
        Dim StockDate As Long
        Dim StockOpen As Double
        Dim StockClose As Double
        
        ' Start and Stop range of individual stocks
        Dim StockStart As Long
        Dim StockStop As Long
        
        ' Looping Counter and check value to see if we have reached the end of the values
        Dim StockRow As Long
        StockRow = 2
        Dim StockTickerCheck As String
        
        ' Output Value Ranges
        Dim OutputValueRow As Integer
        OutputValueRow = 2
        
        
        ' Initialize starting values
        StockRow = 2
        StockStart = StockRow
        StockTicker = ws.Cells(StockRow, 1).Value
        StockOpen = ws.Cells(StockRow, 3).Value
        StockTickerCheck = StockTicker
        
        ' Set initial ticker output value
        ws.Cells(OutputValueRow, 9).Value = StockTicker
    
        ' Loop through column A
        Do While StockTickerCheck <> ""
        
        ' Get stock date and ticker value
        StockDate = ws.Cells(StockRow, 2).Value
        StockTickerCheck = ws.Cells(StockRow, 1).Value
        
        ' Check to see if we have changed to a new stock ticker
        If StockTickerCheck = StockTicker Then
            
            ' Get the closing stock price if it's the last day of year
            If StockDate = 20201231 Or StockDate = 20191231 Or StockDate = 20181231 Then
                StockClose = ws.Cells(StockRow, 6).Value
                
                ' Output Yearly Change
                ws.Cells(OutputValueRow, 10).Value = StockClose - StockOpen
            
                ' Output Percent Change
                ws.Cells(OutputValueRow, 11).Value = (1 - (StockClose / StockOpen)) * (-1)
                
                ' Output Total Volume
                ws.Cells(OutputValueRow, 12).Value = Application.WorksheetFunction.Sum(Range(ws.Cells(StockStart, 7), ws.Cells(StockRow, 7)))
                
                ' Incremenet OutputRowValue
                OutputValueRow = OutputValueRow + 1
            End If
            
        Else
            ' Get the opening stock price
            StockOpen = ws.Cells(StockRow, 3).Value
            
            ' Set new ticker value to StockTickerCheck
            StockTicker = StockTickerCheck
            
            ' Output Ticker
            ws.Cells(OutputValueRow, 9).Value = StockTicker
            
            ' Set StockStart to the beginning of the new ticker symbol
            StockStart = StockRow
             
        End If
        
        ' Increment Row
        StockRow = StockRow + 1
        
        Loop
        
        ' -------------------- BONUS! --------------------
        ' Get the values for MaxInc, MinInc, and MaxVol and set them in the sheet
        Dim MaxInc As Double
        MaxInc = Application.WorksheetFunction.Max(ws.Range("K2:K3001"))
        ws.Range("Q2").Value = MaxInc
        Dim MinInc As Double
        MinInc = Application.WorksheetFunction.Min(ws.Range("K2:K3001"))
        ws.Range("Q3").Value = MinInc
        Dim MaxVol As LongLong
        MaxVol = Application.WorksheetFunction.Max(ws.Range("L2:L3001"))
        ws.Range("Q4").Value = MaxVol
           
        ' Create a range to search through
        Dim SearchRange As Range
        Set SearchRange = ws.Range("I2:L3001")
        
        ' Loop through range searching for the three values and return their corresponding ticker
        For Each cell In SearchRange
        
            If cell.Value = MaxInc Then
                ws.Range("P2").Value = ws.Cells(cell.Row, cell.Column - 2).Value
            ElseIf cell.Value = MinInc Then
                ws.Range("P3").Value = ws.Cells(cell.Row, cell.Column - 2).Value
            ElseIf cell.Value = MaxVol Then
                ws.Range("P4").Value = ws.Cells(cell.Row, cell.Column - 3).Value
            End If
        
        Next
    
    Next ws
    
End Sub