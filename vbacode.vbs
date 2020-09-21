Sub StockDataTest()

'Create script that will loop through stocks for each year and return essential information'

'Loop this script through ALL worksheets'
    For Each ws In Worksheets

        'Name all column headers for output'
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'Name all spaces for CHALLENGE output'
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        'Format Table to auto fit for expected large results'
        ws.Columns("I:Q").AutoFit
        

        'Declare all variable for ticker name'
        
        Dim stock_ticker As String
       
       'Declare variable to hold total stock volume for each ticker name'
       
        Dim total_vol As Double
        total_vol = 0
        
        'Set variable and value of summary table'
        
        Dim SumTblRow As Long
        SumTblRow = 2
        
         
       'Declare variables for year open, year close, and year change prices'
       
        Dim year_open As Double
        Dim year_close As Double
        Dim year_change As Double
        
        Dim firstamount As Long
        firstamount = 2
        
        'Declare percent change and last row'
        
        Dim PercentChange As Double
        
        Dim lastrow As Long
        Dim lastrowValue As Long
       

        'Find the last row through all worksheets'
        
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Loop through last row in all worksheets'
        
        For i = 2 To lastrow

'Find and return total volume and corresponding ticker name'

            'Calculate total volume for each ticker'
            total_vol = total_vol + ws.Cells(i, 7).Value
            
            'Conditional to find if we are within the same ticker'
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            stock_ticker = ws.Cells(i, 1).Value
            
                'Print output in corresponding spaces in summary table'
                
                ws.Range("I" & SumTblRow).Value = stock_ticker
                ws.Range("L" & SumTblRow).Value = total_vol
                
                'Reset total volume to zero so it can keep looping through'
                total_vol = 0


    'Find yearly change in each ticker'
    
                'Set values for yearly open, yearly close and yearly change name'
                year_open = ws.Range("C" & firstamount)
                year_close = ws.Range("F" & i)
                year_change = year_close - year_open
                
                'Print yearly change in summary table'
                ws.Range("J" & SumTblRow).Value = year_change

                'Find percent change for each ticker'
                If year_open = 0 Then
                    PercentChange = 0
                Else
                    year_open = ws.Range("C" & firstamount)
                    PercentChange = year_change / year_open
                End If
                
'Formatting changes'
                
                'Format percent change to appear in percent format'
                
                ws.Range("K" & SumTblRow).NumberFormat = "0.00%"
                ws.Range("K" & SumTblRow).Value = PercentChange

               'Use conditional formatting to indicate positive and negative results green for positive and red for negative'
               
                If ws.Range("J" & SumTblRow).Value >= 0 Then
                    ws.Range("J" & SumTblRow).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & SumTblRow).Interior.ColorIndex = 3
                End If
            
                'Reset summary table row so looping can occur'
                SumTblRow = SumTblRow + 1
                firstamount = i + 1
                End If
            Next i

'CHALLENGE'

'Create a solution that will be able to output the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume'

            'Redefine lastrow for CHALLENGE'
            lastrow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
            'Conditional for greatest % increase loop'
            For i = 2 To lastrow
                If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                    ws.Range("Q2").Value = ws.Range("K" & i).Value
                    ws.Range("P2").Value = ws.Range("I" & i).Value
                End If
            'Conditional for greatest % decrease loop'
                If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                    ws.Range("Q3").Value = ws.Range("K" & i).Value
                    ws.Range("P3").Value = ws.Range("I" & i).Value
                End If
            'Conditional for greatest total volume'
                If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                    ws.Range("Q4").Value = ws.Range("L" & i).Value
                    ws.Range("P4").Value = ws.Range("I" & i).Value
                End If

            Next i
            
        'Format results of greatest increase and decrease to have %'
        
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
            
       

    Next ws

End Sub
