Sub tickerCounter()

    ' First just get the ticker and total stock volume to work on a single wkst- CHECK
    ' 2nd, get the yearly change to work and formatted- CHECK
    ' 3rd, get yearly change to the right color format- CHECK
    ' 4th, get percent change to work - CHECK
    ' 5th, get it to work on all the wksts- CHECK
    ' Move onto bonus section if time- CHECK
    
    ' Loop through all of the worksheets in the workbook
    
    For Each ws In Worksheets
    
        ' Set variable for holding the ticker initials
        Dim ticker As String

        ' Set variable for holding the total stock volume
        Dim total_vol As Double
        total_vol = 0
    
        ' Set var for holding opening stock price for the first day of the year
        Dim year_open As Double
        ' Set year_open for first stock opening price- this will be reset for next stock at end of If statement
        year_open = ws.Cells(2, 3).Value
    
        ' Set var for holding closing stock price for the last day of the year
        Dim year_close As Double
    
        ' Set var for holding yearly change
        Dim year_change As Double
    
        ' Set var for holding percent change
        Dim perc_change As Double
    
        ' Keep track of the location for each ticker in the summary table
        Dim summ_TableRow As Integer
        summ_TableRow = 2

        'Set column header labels for I to L
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        ' Set column header labels for P and Q, row labels for O- Bonus Section
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"

        ' Find the last row for the current wkst
        Dim lastRow, lastRow_summTable As Integer
        lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    
    ' Loop through all the stock prices for the year
    For i = 2 To lastRow

        ' If not the same ticker initials in next row then- (current row =ws.Cells(i, 1).Value and next row = ws.Cells(i + 1, 1).Value)
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            ' Set the ticker initials
            ticker = ws.Cells(i, 1).Value

            ' Add to the total stock volume for the last time
            total_vol = total_vol + ws.Cells(i, 7).Value

            ' Set the closing stock price for the last day of the year for current stock
            year_close = ws.Cells(i, 6).Value
            
            ' Calculate yearly change in stock price and save to var year_change
            year_change = year_close - year_open
            
            ' Calculate percent change in stock price and save to var perc_change (format as perc when printing to wksht)
            perc_change = year_change / year_open
            
            ' Print the Ticker Initials in the stock summary table
            ws.Range("I" & summ_TableRow).Value = ticker
            
            ' Print the Yearly Change in the stock summary table
            ws.Range("J" & summ_TableRow).Value = year_change
            
                'If statement to format color of Yearly Change
                If year_change > 0 Then
                    ' Color positive change green
                    ws.Range("J" & summ_TableRow).Interior.ColorIndex = 4
                Else
                    ' Color negative change red
                    ws.Range("J" & summ_TableRow).Interior.ColorIndex = 3
                    
                End If
                
            ' Print the Percent Change in the stock summary table
            ws.Range("K" & summ_TableRow).Value = perc_change
            
            ' Format Percent Change as percent and two decimal places
            ws.Range("K" & summ_TableRow).NumberFormat = "0.00%"
            
            ' Print the total stock volume to the stock summary table
            ws.Range("L" & summ_TableRow).Value = total_vol

            ' Add one to the stock summary table row
            summ_TableRow = summ_TableRow + 1
            
            ' Reset opening stock price for next ticker stock
            year_open = ws.Cells(i + 1, 3).Value
      
            ' Reset the total stock volume
            total_vol = 0

        ' If the cell after is the same ticker initials
        Else

            ' Add to the stock total volume
            total_vol = total_vol + ws.Cells(i, 7).Value

        End If
        
    Next i

        ' Bonus Section- Greatest % Increase, Greatest % Decrease, and Greatest Total Volume
        ' Find the last row for the stock summary table
        lastRow_summTable = ws.Cells(Rows.Count, "L").End(xlUp).Row
        
        ' Greatest % Increase Code:
        ' Max function to pull out Greatest % Increase from stock summary table and print to the bonus stock summary table
        ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & lastRow_summTable))
        ' Format as a percent with two decimal places
        ws.Range("Q2").NumberFormat = "0.00%"
        ' Use Index and Match functions to pull our the ticker initial of Greatest % Increase
        ws.Range("P2").Value = WorksheetFunction.Index(ws.Range("I2:I" & lastRow_summTable), WorksheetFunction.Match(ws.Range("Q2").Value, ws.Range("K2:K" & lastRow_summTable), 0))
        
        ' Greatest % decrease Code:
        ' Max function to pull out Greatest % decrease from stock summary table and print to the bonus stock summary table
        ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & lastRow_summTable))
        ' Format as a percent with two decimal places
        ws.Range("Q3").NumberFormat = "0.00%"
        ' Use Index and Match functions to pull our the ticker initial of Greatest % decrease
        ws.Range("P3").Value = WorksheetFunction.Index(ws.Range("I2:I" & lastRow_summTable), WorksheetFunction.Match(ws.Range("Q3").Value, ws.Range("K2:K" & lastRow_summTable), 0))
        
         ' Greatest Total Volume Code:
        ' Max function to pull out Greatest Total Volume  from stock summary table and print to the bonus stock summary table
        ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & lastRow_summTable))
        ' Use Index and Match functions to pull our the ticker initial of Greastest Total Volume
        ws.Range("P4").Value = WorksheetFunction.Index(ws.Range("I2:I" & lastRow_summTable), WorksheetFunction.Match(ws.Range("Q4").Value, ws.Range("L2:L" & lastRow_summTable), 0))
        
        
    Next ws
        
End Sub




