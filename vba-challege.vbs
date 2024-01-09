Sub stock_analysis()

        Dim ws As Worksheet
        
        For Each ws In ThisWorkbook.Worksheets
        
                'create a long variable called LastRow
                Dim LastRow As Long
                
                'setting LastRow variable as the last row of data with column 1 as index
                LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
                
                'creating a variable to store the tickers for the summary table
                Dim Ticker_Name As String
                ' creating a variable for the summary table first row and intializing it to be 2
                Dim Summary_Table_Row As Integer
                Summary_Table_Row = 2
                
                ' creating variables for the yearly opening and closing prices for the tickers
                Dim Yearly_OpeningPrice As Double
                Dim Yearly_ClosingPrice As Double
                
                'creating a variable that will store the total stock volume of a ticker
                Dim Total_Stock_Volume As Double
                
                ' creating a variable to keep track of the row index for the next ticker
                Dim col As Long
                
                'creating headers or titles for each of the columns and rows
                ws.Cells(1, 9).Value = "Ticker"
                ws.Cells(1, 10).Value = "Yearly Change"
                ws.Cells(1, 11).Value = "Percent Change"
                ws.Cells(1, 12).Value = "Total Stock Volume"
                ws.Cells(1, 18).Value = "Ticker"
                ws.Cells(1, 19).Value = "Value"
                ws.Cells(2, 17).Value = "Greatest % increase"
                ws.Cells(3, 17).Value = "Greatest % decrease"
                ws.Cells(4, 17).Value = "Greatest total volume"
                
                'creating a starting point for the first row index for opening prices
                Yearly_OpeningPrice = ws.Cells(2, 3).Value
                
                'initializing the row index for tracking the stock volume
                col = 2
                
                'for loop that will loop through each ticker value in column a
                For i = 2 To LastRow
                
                'if statement to that will loop through each ticker to evaluate the whether tickers are same or different
                    If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
                       
                 'storing the ticker string value
                        Ticker_Name = ws.Cells(i, 1).Value
                 
                 ' putting the ticker string value in summary table, which is column I
                        ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
                
                ' yearly closing price will equal the last ticker value before the change in ticker name
                        Yearly_ClosingPrice = ws.Cells(i, 6).Value
                
                ' formula to find the yearly change in ticker value and storing it in summary table, column J
                        ws.Cells(Summary_Table_Row, 10) = Yearly_ClosingPrice - Yearly_OpeningPrice
                
                ' if statement that formats positive changes in green and negative changes in red
                    If ((ws.Cells(Summary_Table_Row, 10) >= 0)) Then
                        ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
                    Else
                        ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
                    End If
                        
                'calculating the percent change from the begining to ending of the year and storing the value in column K
                        ws.Cells(Summary_Table_Row, 11) = FormatPercent(((Yearly_ClosingPrice - Yearly_OpeningPrice) / Yearly_OpeningPrice))
                 
                 ' formula to find the sum of the stock volume of one particular ticker
                       Total_Stock_Volume = WorksheetFunction.Sum(ws.Range(ws.Cells(col, 7), ws.Cells(i, 7)))
                 
                 'storing stock volume in summary table, Column L
                       ws.Cells(Summary_Table_Row, 12) = Total_Stock_Volume
                        
                'grabbing the price for the next ticker
                        Yearly_OpeningPrice = ws.Cells(i + 1, 3).Value
                
                'grabbing the next row for the next ticker
                        col = i + 1
                
                'moving to the next row in the summary table for columns I-L
                        Summary_Table_Row = Summary_Table_Row + 1
                    
                    End If
                Next i
        
                ' creating a variable for the last row in the summary table, Columns - L
                Dim LastRow_Summary_Table As Long
                
                'last row of the summary table is the last row in column I
                LastRow_Summary_Table = ws.Cells(ws.Rows.Count, 9).End(xlUp).row
                
                'creating variable to track row index of the summary table
                Dim col_summary As Long
                
                'set the variable equal to 2 or the first row index
                col_summary = 2
                
                'creating variables to track and store the value and tickers for  greatest % increase, greatest % decrease, and greatest total volume
                Dim max_percent As Double
                Dim min_percent As Double
                Dim max_volume As Variant
                Dim maxp_cell As Long
                Dim result_maxp As Variant
                Dim minp_cell As Long
                Dim result_minp As Variant
                Dim maxv_cell As Long
                Dim result_maxv As Variant
                
                'for loop that will loop through the rows of the summary table
                For i = 2 To LastRow_Summary_Table
                
                        'formula to find the maximum percent change value
                        max_percent = WorksheetFunction.Max(ws.Range(ws.Cells(col_summary, 11), ws.Cells(i, 11)))
                        
                        'storing value in cell s2
                        ws.Cells(2, 19) = FormatPercent(max_percent)
                        
                        'finding the row index of the max volume
                        maxp_cell = WorksheetFunction.Match(max_percent, ws.Range(ws.Cells(col_summary, 11), ws.Cells(i, 11)), 0)
                        
                        'using row index to find the corresponding ticker value
                        result_maxp = ws.Cells(maxp_cell + 1, 9)
                        
                        'storing the ticker value in cell r2
                        ws.Cells(2, 18) = result_maxp
                        
                        'formula to find the minmum percent change value
                        min_percent = WorksheetFunction.Min(ws.Range(ws.Cells(col_summary, 11), ws.Cells(i, 11)))
                        
                        'storing the value in cell s3
                        ws.Cells(3, 19) = FormatPercent(min_percent)
                        
                        'finding the row index of the min volume
                        minp_cell = WorksheetFunction.Match(min_percent, ws.Range(ws.Cells(col_summary, 11), ws.Cells(i, 11)), 0)
                        
                        'using the rowndex to find the corresponding ticker value
                        result_minp = ws.Cells(minp_cell + 1, 9)
                        
                        'storing ticker value in cell r3
                        ws.Cells(3, 18) = result_minp
                        
                        'formula to find the max total volume in summary table
                        max_volume = (WorksheetFunction.Max(ws.Range(ws.Cells(col_summary, 12), ws.Cells(i, 12))))
                        
                        'storing the total volume in cell s4
                        ws.Cells(4, 19) = max_volume
                        
                        'finding the row index of the max volume in summary table
                        maxv_cell = WorksheetFunction.Match(max_volume, ws.Range(ws.Cells(col_summary, 12), ws.Cells(i, 12)), 0)
                        
                        'using the row index to find the corresponding ticker value
                        result_maxv = ws.Cells(maxv_cell + 1, 9)
                        
                        'storing the ticker value in cell r4
                        ws.Cells(4, 18) = result_maxv
                
                Next i
        
        Next ws

End Sub