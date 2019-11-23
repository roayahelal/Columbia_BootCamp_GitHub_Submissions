Attribute VB_Name = "Module1"
Sub RHSTOCKS()

  Dim Ticker_Name As String

  ' Set an initial variable for holding the total per ticker
  Dim Volume_Total As Double
  Volume_Total = 0

  ' Keep track of the location for ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  Dim start_price As Double
  
  Dim end_price As Double
  
  

  ' Loop through all sheets and all tickers
For s = 1 To Worksheets.Count
  
        Sheets(s).Activate
  
            Cells(1, 9).Value = "Ticker"
            Cells(1, 10).Value = "Yearly Change"
            Cells(1, 11).Value = "Percent Change"
            Cells(1, 12).Value = "Total Stock Volume"
            
            Range("K:K").NumberFormat = "0.00%"
            
            row_number_in_sheet = Cells(Rows.Count, 1).End(xlUp).Row
            
            start_price = Cells(2, 3).Value
        
          
            For i = 2 To row_number_in_sheet
        
                ' Check if we are still within the same ticker and if it is not...
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
        
              ' Set the static values in the summary
                Ticker_Name = Cells(i, 1).Value
                
                end_price = Cells(i, 3).Value
                
                Range("J" & Summary_Table_Row).Value = end_price - start_price
                
                Range("K" & Summary_Table_Row).Value = (end_price - start_price) / start_price
                
                    If Range("J" & Summary_Table_Row).Value < 0 Then Range("J" & Summary_Table_Row).Interior.ColorIndex = 3 Else Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
              
              
              ' Add to the  ticker Total (last row)
              Volume_Total = Volume_Total + Cells(i, 7).Value
        
              ' Print the ticker name in the Summary Table
                Range("I" & Summary_Table_Row).Value = Ticker_Name
        
              ' Print the total volume to the Summary Table
                Range("L" & Summary_Table_Row).Value = Volume_Total
                
              ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
              
              ' Reset the ticker Total
                Volume_Total = 0
                
                start_price = Cells(i + 1, 3).Value
        
            ' If the cell immediately following a row is the same brand...
                Else
        
              ' Add to the ticker Total
                Volume_Total = Volume_Total + Cells(i, 7).Value
                
                
                End If
        
            Next i
            
            
            Summary_Table_Row = 2
            

    Next s
        
            


End Sub
