Sub stockAnalysis910()
      ' Set an initial variable for holding the ticker name
    Dim tickerName, greatestTickerVolume, greatestTickerPercentIncrease, greatestTickerPercentDecrease As String
    
      ' Set an initial variable for holding the total per ticker
    Dim volumeTotal, greatestVolume, greatestPercentIncrease, greatestPercentDecrease As Double
    Dim firstClosePrice, yearlyChange As Double 'holds the yearly percentage change
    Dim percentageChange As Double
    
    Dim ws As Worksheet
    For Each ws In Worksheets
        
        
        ws.Range("I" & 1).Value = "Ticker"
        ws.Range("J" & 1).Value = "Yearly Change"
        ws.Range("K" & 1).Value = "Percent Change"
        ws.Range("L" & 1).Value = "Total Stock Volume"
        ws.Range("P" & 1).Value = "Ticker"
        ws.Range("Q" & 1).Value = "Value"
        ws.Range("O" & 2).Value = "Greatest % Increase"
        ws.Range("O" & 3).Value = "Greatest % Decrease"
        ws.Range("O" & 4).Value = "Greatest Total Volume"
          
        volumeTotal = 0
        yearlyChange = 0
        percentageChange = 0
    
    
        ' Set lastRow
        Dim lastRow As Double
        lastRow = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row
    
        ' Keep track of the location for each Stocks Ticker in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
      
        firstClosePrice = Cells(2, 6).Value
        
        ' Loop through all rows
        For i = 2 To lastRow
    
            ' Check if it is the first time reading the ticker
            'Get the first date
            firstDate = Cells(i, 2).Value
            
            
            ' Check if we are still within the same ticker, if it is not...
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                ' Set the ticker name
                tickerName = Cells(i, 1).Value
                
                yearlyChange = Cells(i, 6).Value - firstClosePrice
                
                'check first close price
                If firstClosePrice = 0 Then
                    percentageChange = yearlyChange * 100
                Else
                    percentageChange = Round((yearlyChange / firstClosePrice), 4) * 100
                End If
                
                
                'Set the first close price
                firstClosePrice = Cells(i + 1, 6).Value
                  
                'Add to the volumeTotal
                volumeTotal = volumeTotal + Cells(i, 7).Value
        
                'Print the values
                ws.Range("I" & Summary_Table_Row).Value = tickerName
                ws.Range("J" & Summary_Table_Row).Value = yearlyChange
                ws.Range("K" & Summary_Table_Row).Value = percentageChange
                ws.Range("L" & Summary_Table_Row).Value = volumeTotal
                
                'Highlight yearly change if it is positive
                If yearlyChange >= 0 Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                       
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
        
                ' Reset the volumeTotal
                volumeTotal = 0
        
            ' If the cell immediately following a row is the same ticker...
            Else
               ' Add to the ticker volume total
                volumeTotal = volumeTotal + Cells(i, 7).Value
              
            End If
        Next i
    
        
        greatestVolume = ws.Range("L" & 2).Value
        greatestTickerVolume = ws.Range("I" & 2).Value
        
        greatestPercentIncrease = ws.Range("K" & 2).Value
        greatestTickerPercentIncrease = ws.Range("I" & 2).Value
        
        greatestPercentDecrease = ws.Range("K" & 2).Value
        greatestTickerPercentDecrease = ws.Range("I" & 2).Value
        
        
        'Greatest calculations
        For i = 2 To lastRow
            If greatestVolume < ws.Range("L" & i + 1).Value Then
                greatestVolume = ws.Range("L" & i + 1).Value
                greatestTickerVolume = ws.Range("I" & i + 1).Value
            End If
            
            If greatestPercentIncrease < ws.Range("K" & i + 1).Value Then
                greatestPercentIncrease = ws.Range("K" & i + 1).Value
                greatestTickerPercentIncrease = ws.Range("I" & i + 1).Value
            End If
            
            If greatestPercentDecrease > ws.Range("K" & i + 1).Value Then
                greatestPercentDecrease = ws.Range("K" & i + 1).Value
                greatestTickerPercentDecrease = ws.Range("I" & i + 1).Value
            End If
            
        Next i 'greatest calculations
        ' Write the values to the cells
        ws.Range("Q" & 4).Value = greatestVolume
        ws.Range("P" & 4).Value = greatestTickerVolume
        
        ws.Range("Q" & 3).Value = greatestPercentDecrease & "%"
        ws.Range("P" & 3).Value = greatestTickerPercentDecrease
        
        ws.Range("Q" & 2).Value = greatestPercentIncrease & "%"
        ws.Range("P" & 2).Value = greatestTickerPercentIncrease
        
        ws.Activate
    Next ws
End Sub

