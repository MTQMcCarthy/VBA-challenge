Attribute VB_Name = "Module1"
Sub newMultiYearStockAnalysis()
    'make this happen on every sheet in the book
    For Each ws In Worksheets
        
        'establish the variables
        Dim ticker As String
        Dim tickerCounter As Double
        Dim total As Double
        Dim yearOpen As Double
        Dim yearClose As Double
        Dim lastRow As Double
        Dim lastTicker As Double
        Dim yearlyChange As Double
        Dim percentChange As Double
        
        'set up the results area and add the headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'because I'm getting an overflow error
        On Error Resume Next
        
        'find the last row
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'set the starting values before you start the loop
        ticker = ""
        tickerCounter = 1
        total = 0
        yearOpen = 0
        yearClose = 0
        
        'start the loop
        For Row = 2 To lastRow:
        
            'check if the next ticker differs from current ticker
            If ws.Cells(Row, 1).Value <> ticker Then
                       
                'increase the counter for unique tickers
                tickerCounter = tickerCounter + 1
                
                'set the Ticker
                ticker = ws.Cells(Row, 1).Value
                
                'set the year's open price
                yearOpen = ws.Cells(Row, 3).Value
                    
                'write the Ticker symbol in the results area
                ws.Cells(tickerCounter, 9).Value = ticker
                
                'set the start Total for that first Volume record
                total = ws.Cells(Row, 7).Value
                ws.Cells(tickerCounter, 12) = total
            
              Else
                'when it's the same ticker symbol add the Volume to the Total
                total = total + ws.Cells(Row, 7).Value
                ws.Cells(tickerCounter, 12).Value = total
                    
            End If
            
            'when it's the last entry for this ticker
            If ws.Cells((Row + 1), 1).Value <> ticker Then
            
                'set yearClose value
                yearClose = ws.Cells(Row, 6).Value
            
                'calculate the yearlyChange from yearOpen to yearClose
                yearlyChange = yearClose - yearOpen
                    
                'print to Yearly Change results
                ws.Cells(tickerCounter, 10).Value = yearlyChange
                
                'calculate percentChange from the yearOpen
                If yearOpen = 0 Then
                    percentChange = yearlyChange
                    
                Else
                    percentChange = yearlyChange / yearOpen
                    
                End If
                    
                'print to Percent Change results
                ws.Cells(tickerCounter, 11).Value = percentChange
                
                'apply number format to the cells
                ws.Cells(tickerCounter, 11).NumberFormat = "0.00%"
                
                'apply color format the cells
                If yearlyChange > 0 Then
                    ws.Cells(tickerCounter, 10).Interior.Color = vbGreen
                
                ElseIf yearlyChange < 0 Then
                    ws.Cells(tickerCounter, 10).Interior.Color = vbRed
                
                End If
            
            End If
            
        Next Row
        
        'challenge to check the performers
        
        Dim top_increase As Double
        Dim top_decrease As Double
        Dim top_volume As Double
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        'find the last row of list of Tickers
        lastTicker = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        top_increase = 0
        top_decrease = 0
        top_total = 0
        
        For Row = 2 To lastTicker:
        
            'if the top increase is higher than the current leader
            If ws.Cells(Row, 11).Value > top_increase Then
                               
                'set as new current leader
                top_increase = ws.Cells(Row, 11).Value
                
                'print Ticker and Percent into results
                ws.Range("P2").Value = ws.Cells(Row, 9).Value
                ws.Range("Q2").Value = ws.Cells(Row, 11).Value
            
            End If
            
            'if top decrease is lower than current leader
            If ws.Cells(Row, 11).Value < top_decrease Then
                
                'set as new current leader
                top_decrease = ws.Cells(Row, 11).Value
            
                'print Ticker and Percent into results
                ws.Range("P3").Value = ws.Cells(Row, 9).Value
                ws.Range("Q3").Value = ws.Cells(Row, 11).Value
            
            End If
            
            'if total Stock Volume is higher than current leader
            If ws.Cells(Row, 12).Value > top_total Then
                
                'set as new current leader
                top_total = ws.Cells(Row, 12).Value
                
                'print Ticker and Percent into results
                ws.Range("P4").Value = ws.Cells(Row, 9).Value
                ws.Range("Q4").Value = ws.Cells(Row, 11).Value
            
            End If
        
        Next Row
        
    Next ws
    
End Sub
