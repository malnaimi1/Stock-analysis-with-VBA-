# Stock-analysis-with-VBA-
THis is the code used to automate the analysis 






I have used VBA scripting to analyze real stock market data. The tasks is to create a script that will loop through all the stocks for one year and output the following information.

  1- The ticker symbol.

  2- Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  3- The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  4- The total stock volume of the stock.






    Sub Homework()
    
    For Each ws In Worksheets
 
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total volum"
    
    
    starting_row = 2
   
    summary_table_row = 2
    
    tvol = 0
    
   
    'loop
    Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    For Row = starting_row To Lastrow
        
       'THINGS THAT SHOULD HAPPEN AT THE STARTING OF A NEW TICKER
        Bprice = (ws.Cells(starting_row, 3).Value)
       
        tvol = tvol + ws.Cells(Row, 7).Value
        currentBrand = ws.Cells(Row, 1)
        
        
       'THINGS THAT SHOULD HAPPEN WHEN THE TICKER CHANGE
        If ws.Cells(Row + 1, 1) <> currentBrand Then
         ending_row = Row
         Eprice = (ws.Cells(ending_row, 6).Value)
         yearchange = Eprice - Bprice
         
    
         If Bprice = 0 Then
         
         Percent_Change = 0
         Else
         Percent_Change = (yearchange / Bprice)
         End If
           
         
         ' Update Summary Table
         ws.Cells(summary_table_row, 10).Value = yearchange
         ws.Cells(summary_table_row, 9).Value = currentBrand
         ws.Cells(summary_table_row, 11).Value = Percent_Change
        ws.Cells(summary_table_row, 12).Value = tvol

        'color
            If yearchange <= 0 Then
                'color red
                ws.Cells(summary_table_row, 10).Interior.Color = RGB(255, 0, 0)
            Else
            ' green
             ws.Cells(summary_table_row, 10).Interior.Color = RGB(0, 255, 0)
            
            End If
            
            'formating
            ws.Cells(summary_table_row, 11).NumberFormat = "0.00%"
            
            
        'reset for the new ticker
        starting_row = Row + 1
        tvol = 0
        summary_table_row = summary_table_row + 1
           
        End If
        
    Next Row
    Next ws
    End Sub
