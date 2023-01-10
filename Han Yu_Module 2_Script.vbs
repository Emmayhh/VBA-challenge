Attribute VB_Name = "Module1"
Sub StockTable()

For Each ws In Worksheets
    
    Dim WorksheetName As String

    'Determine the last row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
     
    WorksheetName = ws.Name
    
    StockYear = Split(WorksheetName, "_")
    
 'Add some titles of the summary table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % increase"
    ws.Range("O3").Value = "Greatest % decrease"
    ws.Range("O4").Value = "Greatest total volume"
    
    Dim TickerSymbol As String
    
    Dim Yearly As Double
    Yearly = 0
    
    Dim OpenPrice As Double
    OpenPrice = 0
    
    Dim ClosePrice As Double
    ClosePrice = 0
    
    Dim PercentC As Double
    PercentC = 0
    
    Dim StockTotal As LongLong
    StockTotal = 0
    
    Dim TickerCounter As Integer
    TickerCounter = 2
    
    Dim MaxTicker As String
    Dim MaxPercent As Double
    
    Dim MinTicker As String
    Dim MinPercent As Double
    
    Dim MaxTotalTicker As String
    Dim MaxTotal As LongLong
    
    
    For i = 2 To LastRow
        ' Copy the Ticker symbol
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
            TickerSymbol = ws.Cells(i, 1).Value
            ws.Range("I" & TickerCounter).Value = TickerSymbol
        
        'Caculate the Yearly change
            ClosePrice = ws.Cells(i, 6).Value
            Yearly = ClosePrice - OpenPrice
            ws.Range("J" & TickerCounter).Value = Yearly
        
        ' Highlight positive change in green and negative change in red
                If Yearly < 0 Then
                    ws.Range("J" & TickerCounter).Interior.ColorIndex = 3
        
                ElseIf Yearly > 0 Then
                    ws.Range("J" & TickerCounter).Interior.ColorIndex = 4
                End If
        
        'Caculate the Percent change
                'If Yearly = 0 Then
                    'Range("K" & TickerCounter).Value = 0
        
                'Else
                    PercentC = Yearly / OpenPrice
                    ws.Range("K" & TickerCounter).Value = PercentC
                    ws.Range("K" & TickerCounter).NumberFormat = "0.00%"
                'End If
            
        'Caculate the total number
        StockTotal = StockTotal + Cells(i, 7).Value
        ws.Range("L" & TickerCounter).Value = StockTotal

        'Caculate the Greatest % increse
        MaxPercent = WorksheetFunction.Max(ws.Range("K2:K" & TickerCounter))
        MaxTicker = WorksheetFunction.Index(ws.Range("I2:I" & TickerCounter), WorksheetFunction.Match(MaxPercent, ws.Range("K2:K" & TickerCounter), 0))
        ws.Range("Q2").Value = MaxPercent
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("P2").Value = MaxTicker
        
        'Caculate the Greatest % decrese
        MinPercent = WorksheetFunction.Min(ws.Range("K2:K" & TickerCounter))
        MinTicker = WorksheetFunction.Index(ws.Range("I2:I" & TickerCounter), WorksheetFunction.Match(MinPercent, ws.Range("K2:K" & TickerCounter), 0))
        ws.Range("Q3").Value = MinPercent
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("P3").Value = MinTicker
        
        'Caculate the Greatest % decrese
        MaxTotal = WorksheetFunction.Max(ws.Range("L2:L" & TickerCounter))
        MaxTotalTicker = WorksheetFunction.Index(ws.Range("I2:I" & TickerCounter), WorksheetFunction.Match(MaxTotal, ws.Range("L2:L" & TickerCounter), 0))
        ws.Range("Q4").Value = MaxTotal
        ws.Range("P4").Value = MaxTotalTicker
            
    
        TickerCounter = TickerCounter + 1
    
        StockTotal = 0
    
    ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1) Then
        OpenPrice = ws.Cells(i, 3).Value
    
    Else
    StockTotal = StockTotal + Cells(i, 7).Value
    
    End If
    
    Next i
    Next ws
End Sub
