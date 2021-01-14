Attribute VB_Name = "Module2"
Sub StockTracker()

    
        ' Set the Summary Table Headings
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        ' Declare variables to hold values
        Dim Ticker As String
        Dim NextTicker As String
        Dim OpeningPrice As Double
        Dim ClosingPrice As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim StockVolume As Double
        
        ' Set Opening Price for the first stock
        OpeningPrice = Cells(2, 3).Value
        
        ' Create Summary Table and set initial value to 2
        Dim SummaryTableRow As Integer
        SummaryTableRow = 2
        
        ' Determine the last row and declare it as a variable
        Dim LastRow As Long
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Begin the loop
        For i = 2 To LastRow
        
            ' Find the ticker value
            Ticker = Cells(i, 1).Value
            NextTicker = Cells(i + 1, 1).Value
            
            ' Check for ticker changes
            If NextTicker <> Ticker Then
            
                ' Get the closing price from the row
                ClosingPrice = Cells(i, 6).Value
                
                ' Calculate the Yearly Change
                YearlyChange = ClosingPrice - OpeningPrice
                
                ' Calculate the Percent Change
                If (OpeningPrice = 0) Then
                    PercentChange = 0
                Else
                    PercentChange = (YearlyChange / OpeningPrice)
                End If
                
                ' Calculate the Stock Volume
                StockVolume = StockVolume + Cells(i, 7).Value
                
                ' Write the table values to the appropriate cells
                Cells(SummaryTableRow, 9).Value = Ticker
                Cells(SummaryTableRow, 10).Value = YearlyChange
                Cells(SummaryTableRow, 11).Value = PercentChange
                Cells(SummaryTableRow, 12).Value = StockVolume
                
                ' Set values back to zero
                StockVolume = 0
                ClosingPrice = 0
                OpeningPrice = 0
                
                ' Set the next OpeningPrice
                OpeningPrice = Cells(i + 1, 3).Value
               
                ' Conditional formatting for YearlyChange
                If YearlyChange > 0 Then
                    Cells(SummaryTableRow, 10).Interior.ColorIndex = 4
                Else
                    Cells(SummaryTableRow, 10).Interior.ColorIndex = 3
                End If
                
                ' Move to the next row in the summary table before posting the next value
                SummaryTableRow = SummaryTableRow + 1
                
            Else
                 
                ' Add the StockVolume from each row and add to to the SummaryTable
                StockVolume = StockVolume + Cells(i, 7).Value
    
            End If
            
        Next i
    
        
        ' Format decimals for the yearly change, percent change columns
        For i = 2 To LastRow
            Cells(i, 11).NumberFormat = "0.00%"
        Next i
        
        For i = 2 To LastRow
            Cells(i, 10).NumberFormat = "0.00"
        Next i
        
        ' Autofit cells
        Columns("A:P").AutoFit
        
    
End Sub


