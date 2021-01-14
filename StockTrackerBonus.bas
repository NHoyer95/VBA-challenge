Attribute VB_Name = "Module1"
Sub StockTrackerBonus()

    ' Loop through all sheets
    For Each ws In Worksheets

        ' Set the Summary Table Headings
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Declare variables to hold values
        Dim Ticker As String
        Dim NextTicker As String
        Dim OpeningPrice As Double
        Dim ClosingPrice As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim StockVolume As Double
        
        ' Set Opening Price for the first stock
        OpeningPrice = ws.Cells(2, 3).Value
        
        ' Create Summary Table and set initial value to 2
        Dim SummaryTableRow As Integer
        SummaryTableRow = 2
        
        ' Determine the last row and declare it as a variable
        Dim LastRow As Long
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Begin the loop
        For i = 2 To LastRow
        
            ' Find the ticker value
            Ticker = ws.Cells(i, 1).Value
            NextTicker = ws.Cells(i + 1, 1).Value
            
            ' Check for ticker changes
            If NextTicker <> Ticker Then
            
                ' Get the closing price from the row
                ClosingPrice = ws.Cells(i, 6).Value
                
                ' Calculate the Yearly Change
                YearlyChange = ClosingPrice - OpeningPrice
                
                ' Calculate the Percent Change
                If (OpeningPrice = 0) Then
                    PercentChange = 0
                Else
                    PercentChange = (YearlyChange / OpeningPrice)
                End If
                
                ' Calculate the Stock Volume
                StockVolume = StockVolume + ws.Cells(i, 7).Value
                
                ' Write the table values to the appropriate cells
                ws.Cells(SummaryTableRow, 9).Value = Ticker
                ws.Cells(SummaryTableRow, 10).Value = YearlyChange
                ws.Cells(SummaryTableRow, 11).Value = PercentChange
                ws.Cells(SummaryTableRow, 12).Value = StockVolume
                
                ' Set values back to zero
                StockVolume = 0
                ClosingPrice = 0
                OpeningPrice = 0
                
                ' Set the next OpeningPrice
                OpeningPrice = ws.Cells(i + 1, 3).Value
               
                ' Conditional formatting for YearlyChange
                If YearlyChange > 0 Then
                    ws.Cells(SummaryTableRow, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(SummaryTableRow, 10).Interior.ColorIndex = 3
                End If
                
                ' Move to the next row in the summary table before posting the next value
                SummaryTableRow = SummaryTableRow + 1
                
            Else
                 
                ' Add the StockVolume from each row and add to to the SummaryTable
                StockVolume = StockVolume + ws.Cells(i, 7).Value
    
            End If
            
        Next i
    
        ' Find the last row of the summary table
        Dim LastRowBonus As Long
        LastRowBonus = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
        ' Declare Headings for the Bonus Table
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        
        ' Declare Variables
        Dim BonusTicker1 As String
        Dim BonusTicker2 As String
        Dim BonusTicker3 As String
        Dim GreatestIncrease As Double
        Dim GreatestDecrease As Double
        Dim GreatestTotalVolume As Double
        
        ' Set the initial value of the Greatest's?
        GreatestIncrease = ws.Cells(2, 11).Value
        GreatestDecrease = ws.Cells(2, 11).Value
        GreatestTotalVolume = ws.Cells(2, 12).Value
        
        ' Begin Next Loop
        For i = 2 To LastRowBonus
        
            ' Check for conditions and write to bonus table
            If ws.Cells(i, 11).Value > GreatestIncrease Then
                GreatestIncrease = ws.Cells(i, 11).Value
                BonusTicker1 = ws.Cells(i, 9).Value
            End If
            
            If ws.Cells(i, 11).Value < GreatestDecrease Then
                GreatestDecrease = ws.Cells(i, 11).Value
                BonusTicker2 = ws.Cells(i, 9).Value
            End If
            
            If ws.Cells(i, 12).Value > GreatestTotalVolume Then
                GreatestTotalVolume = ws.Cells(i, 12).Value
                BonusTicker3 = ws.Cells(i, 9).Value
            End If
            
            ' Write the values to their cells and format
            ws.Cells(2, 15).Value = BonusTicker1
            ws.Cells(3, 15).Value = BonusTicker2
            ws.Cells(4, 15).Value = BonusTicker3
            ws.Cells(2, 16).Value = GreatestIncrease
            ws.Cells(2, 16).NumberFormat = "0.00%"
            ws.Cells(3, 16).Value = GreatestDecrease
            ws.Cells(3, 16).NumberFormat = "0.00%"
            ws.Cells(4, 16).Value = GreatestTotalVolume
            
        Next i
        
        ' Format decimals for the yearly change, percent change columns
        For i = 2 To LastRow
            ws.Cells(i, 11).NumberFormat = "0.00%"
        Next i
        
        For i = 2 To LastRow
            ws.Cells(i, 10).NumberFormat = "0.00"
        Next i
        
        ' Autofit cells
        ws.Columns("A:P").AutoFit
        
    Next ws
    
End Sub

