'Process
'1. Loop through each worksheet to get summary data
'2. Get a table on the side, row 1 is the ticker name of the stock
'3. Get a yearly change in row 2 of the summary data table
'4. Conditional formatting green for yearly change that is positive, add red if neg
'5. Get a percentage change in row 3 and add % symbols
'6. Sum the volumn to get total stock volumn in row 4
'7. Bonus bit - start with finding max, min of percentage and max total val
'8. Find matching ticker name
'9. Autofit columns so that it shows the data


Sub StockDataSummaryTable()

For Each ws In Worksheets
  
    'Set an initial variable
    
    Dim TickerName As String
    Dim YearlyChange As Double
    Dim PercentageChange As Double
    Dim Summary_Table_Row As Long
    Summary_Table_Row = 2
    Dim YearlyOpen As Double
    Dim YearlyClose As Double
    Dim LastOpen As Long
    LastOpen = 2
    

    
    'Set column headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percentage Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
     
    'Define last row
        
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To LastRow
        
        'Pick the ticker name and calculate the ticker volume first
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            
            TickerName = ws.Cells(i, 1).Value
        
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            
            'Find Yearly change
            
            YearlyOpen = ws.Range("C" & LastOpen).Value
            YearlyClose = ws.Range("F" & i).Value
            YearlyChange = YearlyClose - YearlyOpen
      
            
            'Find Percentage Change
            
            If YearlyOpen <> 0 Then
                PercentageChange = YearlyChange / YearlyOpen
    
            Else
                PercentageChange = 0
                
            End If
    
            
            'Print the output
            ws.Range("I" & Summary_Table_Row).Value = TickerName
            ws.Range("J" & Summary_Table_Row).Value = Format(YearlyChange, "#,##0.00")
            ws.Range("K" & Summary_Table_Row).Value = Format(PercentageChange, "Percent")
            ws.Range("L" & Summary_Table_Row).Value = TotalVolume
                  
            'Conditional Formating YearlyChange
            
            If YearlyChange >= 0 Then
            
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                
            Else
                
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            
            End If
            
            'Move to the next row
            Summary_Table_Row = Summary_Table_Row + 1
            LastOpen = i + 1
            
            TotalVolume = 0
            
        Else
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            
        End If
            
            
    Next i
        
    
     'Find Highest % Increase, Highest% decrease and Max Total Volume'
            
        LastRow2 = ws.Cells(Rows.Count, 11).End(xlUp).Row
            
        Set Rng = ws.Range("K" & 2 & ":" & "K" & LastRow2)
        Set Rng2 = ws.Range("L" & 2 & ":" & "L" & LastRow2)
            
            MaxPercentage = Application.WorksheetFunction.Max(Rng)
            MinPercentage = Application.WorksheetFunction.Min(Rng)
            MaxTotalVolume = Application.WorksheetFunction.Max(Rng2)
            
            ws.Range("Q2").Value = Format(MaxPercentage, "Percent")
            ws.Range("Q3").Value = Format(MinPercentage, "Percent")
            ws.Range("Q4").Value = MaxTotalVolume
            
        'Find matching ticker
        
        For i = 2 To LastRow2
        
            If ws.Range("K" & i) = ws.Range("Q2").Value Then
                ws.Range("P2") = ws.Range("I" & i)
                
            End If
            
            If ws.Range("K" & i) = ws.Range("Q3").Value Then
                ws.Range("P3") = ws.Range("I" & i)
                
            End If
            
            If ws.Range("L" & i) = ws.Range("Q4").Value Then
                ws.Range("P4") = ws.Range("I" & i)
                
            End If
            
        Next i
            
    ws.Columns("I:Q").AutoFit
        
Next ws

End Sub
