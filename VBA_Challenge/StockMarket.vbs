Sub StockMarket()
'Set ws as worksheet variable
    Dim ws As Worksheet
    
           'Loop through all the worksheets
            For Each ws In Worksheets
            
                   'Set the column headers for the summary table on each worksheet
                    ws.Range("I1").Value = "Ticker"
                    ws.Range("J1").Value = "Yearly Change"
                    ws.Range("K1").Value = "Percent Change"
                    ws.Range("L1").Value = "Total Stock Volume"
                    ws.Range("O2").Value = "Greatest % Increase"
                    ws.Range("O3").Value = "Greatest % Decrease"
                    ws.Range("O4").Value = "Greatest Total Value"
                    ws.Range("P1").Value = "Ticker"
                    ws.Range("Q1").Value = "Value"
                  

                   'Define initial variables for holding the values of Ticker letter and the Open price, Final price, Price Percent change and Total volume per Ticket letter.
                   'Set the Summary table (counter) to keep track of the location of the total values per Ticket letter.
                   'and the First day of the year (counter) for the Open Price to keep track for that date.
                    Dim LastRow As Long
                    Dim Ticker As String
                    Dim Open_price As Double
                    Dim Close_price As Double
                    Dim Yearly_change As Double
                    Dim Percent_change As Double
                    Dim Total_volume As Double
                    Total_volume = 0
                    Dim Summary_Table_Row As Long
                    Summary_Table_Row = 2
                    Dim First_yearday_amount As Long
                    First_yearday_amount = 2
                    Dim Greatest_increase As Double
                    Dim Greatest_decrease As Double
                    Dim Greatest_Total_Volume As Double
                    Dim LastRowValue As Long
                                     
 
                    'Set a variable to determine the last row on each worksheet
                    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                    
                    'Loop through rows
                    For i = 2 To LastRow
                    
                        'Loop through columns to check if we are still within the same Ticker letter, if it is not...
                        If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then                          '
                                         Total_volume = Total_volume + ws.Cells(i, 7).Value
                                         
                        Else
                        
                                         'Grab the Ticker letter
                                         Ticker = ws.Cells(i, 1).Value
                                         'Add the corresponding volume to the Total
                                         Total_volume = Total_volume + ws.Cells(i, 7).Value
                                         'Print the Ticker letter in the Summary Table
                                         ws.Range("I" & Summary_Table_Row).Value = Ticker
                                         'Print the Total Volume per Ticker in the Summary Table
                                         ws.Range("L" & Summary_Table_Row).Value = Total_volume
                                         'Reset the Total Volume
                                         Total_volume = 0
                                                             
                                         'Set the variables Open_price, Close_price, Yearly_change and Percent_change.
                                         Open_price = ws.Range("C" & First_yearday_amount).Value
                                         Close_price = ws.Range("F" & i).Value
                                         Yearly_change = Close_price - Open_price
                                         ws.Range("J" & Summary_Table_Row).Value = Yearly_change
                                         
                                         
                                        ' Define Percent Change
                                        If Open_price = 0 Then
                                            Percent_change = 0
                                        Else
                                            Open_price = ws.Range("C" & First_yearday_amount)
                                            Percent_change = Yearly_change / Open_price
                                        End If
                                         
                                         
                                         ws.Range("K" & Summary_Table_Row).Value = Percent_change
                                         'Format Percent change to include two decimals and the symbol %
                                         ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                                        
                                             'Give format to column Yearly Change
                                             If ws.Range("J" & Summary_Table_Row).Value >= 0 Then
                                                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                                             Else
                                                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                                             End If
                                    
                                    'Add one to add a row to the Summary table.
                                    Summary_Table_Row = Summary_Table_Row + 1
                                    'Add one to find the next row that contains the value with the first day of the year for the next letter
                                    First_yearday_amount = i + 1
  
                                 
                        End If
             
                    
                    Next i
                                    'Set a variable to determine the last row on column Percent Change
                                    LastRowValue = ws.Cells(Rows.Count, 11).End(xlUp).Row
                                    ws.Range("Q2").Value = Greatest_increase
                                    
                                    'Loop through rows looking fo results
                                    For i = 2 To LastRow
                                    
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
                                    'Format final values to show two decimals and % Symbol
                                        ws.Range("Q2").NumberFormat = "0.00%"
                                        ws.Range("Q3").NumberFormat = "0.00%"
                                        
                                    'Format Table Columns To Auto Fit
                                    ws.Columns("I:Q").AutoFit
                 
        Next ws
        
        MsgBox ("Fix complete")
End Sub
