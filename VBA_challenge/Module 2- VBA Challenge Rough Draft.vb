Sub tickerCounter()

    ' First just get the ticker and total stock volume to work on a single wkst- CHECK
    ' 2nd, get the yearly change to work and formatted- CHECK
    ' 3rd, get yearly change to the right color format
    ' 4th, get percent change to work - CHECK
    ' 5th, get it to work on all the wksts
    ' Move onto bonus section if time
    

    ' Set variable for holding the ticker initials
    Dim ticker As String

    ' Set variable for holding the total stock volume
    Dim total_vol As Double
    total_vol = 0
    
    ' Set var for holding opening stock price for the first day of the year
    Dim year_open As Double
    ' Set year_open for first stock opening price- this will be reset for next stock at end of If statement
    year_open = Cells(2, 3).Value
    
    ' Set var for holding closing stock price for the last day of the year
    Dim year_close As Double
    
    ' Set var for holding yearly change
    Dim year_change As Double
    
    ' Set var for holding percent change
    Dim perc_change As Double
    
    ' Keep track of the location for each ticker in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    'Set column header labels for I to L
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

    ' Find the last row for the current wkst
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
 
    
  ' Loop through all the stock prices for the year
    For i = 2 To lastRow

        ' Check if we are still within the same credit card brand, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            ' Set the ticker initials
            ticker = Cells(i, 1).Value

            ' Add to the total stock volume
            total_vol = total_vol + Cells(i, 7).Value

            ' Set the closing stock price for the last day of the year
            year_close = Cells(i, 6).Value
            
            ' Calculate yearly change in stock price and save to var year_change
            year_change = year_close - year_open
            
            ' Calculate percent change in stock price and save to var perc_change then format as perc
            perc_change = year_change / year_open
            
            ' Print the Ticker Initials in the Summary Table
            Range("I" & Summary_Table_Row).Value = ticker
            
            ' Print the Yearly Change in the Summary Table
            Range("J" & Summary_Table_Row).Value = year_change
            
            ' Print the Percent Change in the Summary Table
            Range("K" & Summary_Table_Row).Value = perc_change

            ' Print the total stock volume to the Summary Table
            Range("L" & Summary_Table_Row).Value = total_vol

            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            
            ' Reset opening stock price for next ticker stock
            year_open = Cells(i + 1, 3).Value
      
            ' Reset the total stock volume
            total_vol = 0

        ' If the cell immediately following a row is the same brand...
        Else

            ' Add to the Brand Total
            total_vol = total_vol + Cells(i, 7).Value

        End If

    Next i

    ' Format Percent Change as two decimals places and %
    Range("K2:K" & lastRow).NumberFormat = "0.00%"
    
    
End Sub

