Attribute VB_Name = "Module1"
Sub StockTest3():

    For Each ws In Worksheets
      ' create script to loop for ticker, yearly change in prices, percentage of change and total stock volume for all worksheets
        ' insert ticker, yearly change, percent change, and total stock volumn Column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
       ' Dim WorksheetName As String
      
       ' WorksheetName = ws.Name
       
       ' variable to hold the ticker, open price, close price and d
       Dim ticker As String
       Dim openPrice As Double
       open_Row = 2
       
       Dim closePrice As Double
       Dim yearlyChange As Double
       Dim percentChange As Double
       
       
         
       ' summary table row variable
       Dim summaryTableRow As Long
       summaryTableRow = 2
       
        ' count number of rows\
         ' variable to hold last row
       Dim lastRow As String
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
           
   
        
         ' loop through all the ticker rows
    For Row = 2 To lastRow
    
             ' check if still same ticker, if not do the following
             If ws.Cells(Row, 1).Value <> Cells(Row + 1, 1).Value Then
         
                ' Set (reset) ticker
                ticker = ws.Range("A" & Row).Value
         
                ' calculate stock value
                 stockVolTotal = stockVolTotal + ws.Range("G" & Row).Value
         
                 ' calculate yearly change
                 openPrice = ws.Range("C" & open_Row).Value
                 closePrice = ws.Cells(Row, 6).Value
                 yearlyChange = closePrice - openPrice
                  
                      If openPrice = 0 Then
                         percentChange = 0
                        Else
                           percentChange = yearlyChange / openPrice
                     End If
                  
                  ' add values to summaryTable
                 ws.Range("I" & summaryTableRow).Value = ticker
                 ws.Range("L" & summaryTableRow).Value = stockVolTotal
                 ws.Range("J" & summaryTableRow).Value = yearlyChange
                 ws.Range("J" & summaryTableRow).NumberFormat = "$#,##0.00"
                 ' calculate percentage
                 ws.Range("K" & summaryTableRow).Value = percentChange
                 ws.Range("K" & summaryTableRow).NumberFormat = "0.00%"
                        ' highlight negative cells in red and others in green for yearly change
                        If ws.Range("J" & summaryTableRow).Value < 0# Then
                            ws.Range("J" & summaryTableRow).Interior.ColorIndex = 3
                        Else
                            ws.Range("J" & summaryTableRow).Interior.ColorIndex = 4
                        End If
                 ' add one to summary row count once summary table populated
                  summaryTableRow = summaryTableRow + 1
                  open_Row = Row + 1
                ' reset stock value
                  stockVolTotal = 0
                  ticker = 0
           Else
                 ' if using the same ticker, calculate total stock value
                 stockVolTotal = stockVolTotal + ws.Range("G" & Row).Value
           
        End If
     Next Row
          
 
  
 Next ws
End Sub

            

