Sub StockAnalysis()
                'Create a script that loops through all the stocks for one year and outputs the following information:
                  ' The ticker symbol.
                  'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
                  'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
                  'The total stock volume of the stock.

           ' Loop  through all worksheets
                  Dim ws As Worksheet
                For Each ws In Worksheets

                        ' Column Headers
                        ws.Range("I1").Value = "Ticker"
                        ws.Range("J1").Value = "Yearly Change"
                        ws.Range("K1").Value = "Percent Change"
                        ws.Range("L1").Value = "Total Stock Volume"
                       
                        ' Declare initial variables
                        Dim ticker As String
                        Dim next_ticker As String
                        Dim LR As Long
                        Dim total_ticker_volume As Double
                        total_ticker_volume = 0
                        Dim summary_table_row As Long
                        summary_table_row = 2
                        Dim yearly_open As Double
                        Dim yearly_close As Double
                        Dim yearly_change As Double
                        Dim previous_amount As Long
                        previous_amount = 2
                        Dim percent_change As Double
        

                        ' Determine the Last Row
                        LR = ws.Cells(Rows.Count, 1).End(xlUp).Row
                         
        'loop  for tickers
      For i = 2 To LR

                                ticker = ws.Cells(i, 1).Value
                                next_ticker = ws.Cells(i + 1, 1).Value
                total_ticker_volume = total_ticker_volume + ws.Cells(i, 7).Value
            
                    ' Check If We Are Still Within The Same Ticker Name If It Is Not...
        If next_ticker <> ticker Then
                                
                                ' Print the ticker name in the summary table
                                ws.Range("I" & summary_table_row).Value = ticker
                                ' Print the ticker total volume to the summary table
                                ws.Range("L" & summary_table_row).Value = total_ticker_volume
                                ' Reset ticker total
                                total_ticker_volume = 0
            
                                    ' Set yearly open, yearly close and yearly change name
                                yearly_open = ws.Range("C" & previous_amount)
                                yearly_close = ws.Range("F" & i)
                                yearly_change = yearly_close - yearly_open
                                ws.Range("J" & summary_table_row).Value = yearly_change
            
                                    ' Determine percent change
                    If yearly_open = 0 Then
                                            percent_change = 0
                         Else
                                    yearly_open = ws.Range("C" & previous_amount)
                                    percent_change = yearly_change / yearly_open
                    End If
                                    ' Format double to include % symbol and two decimal places
                                    ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
                                    ws.Range("K" & summary_table_row).Value = percent_change
                    
                                    ' Conditional formatting highlight positive (green) / negative (red)
                    If ws.Range("J" & summary_table_row).Value >= 0 Then
                                        ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
                        Else
                                    ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
                    End If
                            
                                ' Add one to the summary table row
                                summary_table_row = summary_table_row + 1
                                previous_amount = i + 1
        End If
                        Next i

' ### Bonus ###
         'Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
                      
                        ' Column headers
                               ws.Range("O2").Value = "Greatest % Increase"
                               ws.Range("O3").Value = "Greatest % Decrease"
                              ws.Range("O4").Value = "Greatest Total Volume"
                              ws.Range("P1").Value = "Ticker"
                              ws.Range("Q1").Value = "Value"
                              
                    
                      ' Declare initial variablesorur table
                            Dim greatest_increase As Double
                            greatest_increase = 0
                            Dim greatest_decrease As Double
                            greatest_decrease = 0
                            Dim LRValue As Long
                            Dim greatest_total_volume As Double
                            greatest_total_volume = 0
                
                    '  Start loop for final results
        For i = 2 To LR
        
            If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                        ws.Range("Q2").Value = ws.Range("K" & i).Value
                        ws.Range("P2").Value = ws.Range("I" & i).Value
            End If
            

            If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                        ws.Range("Q3").Value = ws.Range("K" & i).Value
                        ws.Range("P3").Value = ws.Range("I" & i).Value
            End If

            If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                    ws.Range("Q4").Value = ws.Range("L" & i).Value
                    ws.Range("P4").Value = ws.Range("I" & i).Value
            End If

       Next i
                    ' Format double to include % symbol and two decimal places
                         ws.Range("Q2").NumberFormat = "0.00%"
                         ws.Range("Q3").NumberFormat = "0.00%"
                    'Bold table  header
                         ws.Range("A1:Q1").Font.Bold = True
                      'Format table columns to auto fit
                          ws.Columns("A:Q").AutoFit

    Next ws

End Sub

