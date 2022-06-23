Attribute VB_Name = "Module1"
' This program analyzes sample stock data stored on multiple sheets by looping and outputting the following into a summary table:
    ' Ticker symbols
    ' Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    ' The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    ' The total stock volume of the stock.

Sub stockChallenge():
    
    ' set variable to define to define objects:
    
    ' Ticker
    Dim ticker As String
    ' Open
    Dim firstOpen As Double
    ' Close
    Dim lastClose As Double
    ' Total Stock Volume
    Dim totalVolume As Double
    ' Percent Change
    Dim percentChange As Double
    ' Yearly Change
    Dim yearlyChange As Double
    ' Summary Table Row Value
    Dim summaryTableRow As Integer
         
    
    ' Begin worksheet and row loop function
    
        ' to loop worksheets
        For Each ws In Worksheets
        
        ' Set summaryTable headers and autofit formatting
        ws.Range("I1") = "Ticker"
        ws.Range("I1").Columns.AutoFit
        ws.Range("J1") = "Yearly" + " " + "Change"
        ws.Range("J1").Columns.AutoFit
        ws.Range("K1") = "Percent" + " " + "Change"
        ws.Range("K1").Columns.AutoFit
        ws.Range("L1") = "Total" + " " + "Stock" + " " + "Volume"
        ws.Range("L1").Columns.AutoFit
        
        ' Define Summary Data Table start
        summaryTableRow = 2
    
        ' total volume Beginning value
        totalVolume = 0
    
        ' percentChange = 0
    
        ' Define first open value
        firstOpen = ws.Cells(2, 3)
    
        ' function that finds last row of stock data
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
            ' loop from row 2 in column A out to the last row
            For Row = 2 To lastRow

        
            ' check to see if ticker changes from row to row
            If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
            
         
            'set the ticker and opening value in summary table
            ticker = ws.Cells(Row, 1).Value
        
            ' add ticker to the I row of the summary table
            ws.Cells(summaryTableRow, 9).Value = ticker
    
            ' add volume from the row
            totalVolume = totalVolume + ws.Cells(Row, 7).Value
        
            ' add totalVolume to the L row of the summary table
            ws.Cells(summaryTableRow, 12).Value = totalVolume
            
                    
            ' Define last close value
            lastClose = ws.Cells(Row, 6)
            
            ' Begin calculations for yearly change
            yearlyChange = lastClose - firstOpen
                
            ' add yearly change to the J row of the summary table
            ws.Cells(summaryTableRow, 10).Value = yearlyChange
                
            If firstOpen = 0 Then
                percentChange = 0
                    
                Else
                    percentChange = (yearlyChange / firstOpen) * 100
                
            End If
                
            ' add Percent Change to the K row of the summary table
            ws.Cells(summaryTableRow, 11).Value = FormatPercent(percentChange / 100, 2)
                
                
            ' Impose condtional formatting:
                ' Yearly change greater than 1 = Green
                ' Yearly Change less than 1 = Red
            If ws.Cells(summaryTableRow, 10) < 0 Then
                    ws.Cells(summaryTableRow, 10).Interior.ColorIndex = 3
                    
            Else
                ws.Cells(summaryTableRow, 10).Interior.ColorIndex = 4
            
            End If
            
             
            
            ' Go to next summary table row
            summaryTableRow = summaryTableRow + 1
            
            'reset Total Volume to zero
            totalVolume = 0
    
            firstOpen = ws.Cells(Row + 1, 3)

    
        Else
            ' if ticker is same, add <vol>
            totalVolume = totalVolume + ws.Cells(Row, 7).Value
        

        End If
    
    Next Row
    
  Next ws
    
End Sub
