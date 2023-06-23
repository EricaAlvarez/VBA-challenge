Attribute VB_Name = "Multiple_Year_Stock"
Sub Multiple_Year_Stock()
    
    For Each ws In Worksheets
    
        ' Column Labeling
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Chance"
        ws.Range("L1") = "Total Stock Volume"
        
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
    
        ' Keep track of the location for each credit card brand in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        ' Set an initial variable for holding the ticker symbol
        Dim Ticker_Symbol As String
         
        ' Set an initial variable for holding general variables
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        
        Dim Stock_Volume As Double
        Stock_Volume = 0
        
        ' Set an initial variable for holding the opening price
        Dim Open_Price As Double
        Open_Price = ws.Cells(2, 3).Value
         
        ' Set a variable for holding the closing price
        Dim Close_Price As Double
        
        ' Indicates how to find data Lastrow
        Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        
        '------------------------
    
        ' LOOP TROUGH ALL THE ROWS BY THICKERS (alpha-numeric sorted name needed)
        For i = 2 To Lastrow
         
            ' Check if we are still within the same ticker. If it is not, then...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            ' Set the ticker symbol
            Ticker_Symbol = ws.Cells(i, 1).Value
         
            ' Set the close price
            Close_Price = ws.Cells(i, 6).Value
            
            'Set and calculate the yearly change
            Yearly_Change = Close_Price - Open_Price
            
            ' Print the ticker symbol in the summary Table
            ws.Range("I" & Summary_Table_Row).Value = Ticker_Symbol
            
            'Print the yearly change between the open and the close
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
            
            'Calculate the percent change
            Percent_Change = Yearly_Change / Open_Price
               
            'Print the percent change between the open and the close
            ws.Range("K" & Summary_Table_Row).Value = Percent_Change
            ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            
            'Calculate the stock volume
            Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
            
            'Print the stock volume
            ws.Range("L" & Summary_Table_Row).Value = Stock_Volume
            
            '------------------------
            
            'Change Row. Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
              
            ' Reset the opening price
            Open_Price = ws.Cells(i + 1, 3)
            
            'Reset the stock volume
            Stock_Volume = 0
            
            Else
            
            'Add the stock volume
            Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
          
            
            End If

            
        Next i
         
        '------------------------
         
        ' Indicates how to find summary table Lastrow
        Lastrow_Summary_Table = Cells(Rows.Count, 9).End(xlUp).Row
         
        ' CONDITIONAL FORMATTING APPLIED TO THE YEARLY CHANCE COLUMN
        For i = 2 To Lastrow_Summary_Table
         
                         
                If ws.Cells(i, 10).Value > 0 Then
                    ws.Cells(i, 10).Interior.ColorIndex = 50
                    
                Else
                    ws.Cells(i, 10).Interior.ColorIndex = 3
                    
                End If
         
        Next i
         
          
        ' CONDITIONAL FORMATTING APPLIED TO THE PERCENT CHANCE COLUMN
        For i = 2 To Lastrow_Summary_Table
        
                If ws.Cells(i, 11).Value > 0 Then
                    ws.Cells(i, 11).Interior.ColorIndex = 50
                    
                Else
                    ws.Cells(i, 11).Interior.ColorIndex = 3
                    
                End If
         
        Next i
         
                  
          
        ' LOOP TROUGH SUMMARY TABLE ROWS TO FIND THE GREATER VALUES
         
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"
        
        Dim Greatest_Increase As Double
        Greatest_Increase = 0
        
        Dim Greatest_Decrease As Double
        Greatest_Decrease = 0
        
        Dim Greatest_Volume As Double
        Greatest_Volume = 0
                 
        For i = 2 To Lastrow_Summary_Table
            
                If ws.Cells(i, 11).Value > Greatest_Increase Then
                
                'Set value for the greatest increase
                Greatest_Increase = ws.Cells(i, 11).Value
                     
                'Print the greates increase
                ws.Range("P" & 2).Value = ws.Cells(i, 9)
                ws.Range("Q" & 2).Value = Greatest_Increase
                ws.Range("Q" & 2).NumberFormat = "0.00%"
                    
                Else
                           
                    If ws.Cells(i, 11).Value < Greatest_Decrease Then
                
                    'Set value for the greatest decrease
                    Greatest_Decrease = ws.Cells(i, 11).Value
                     
                    'Print the greates decrease
                    ws.Range("P" & 3).Value = ws.Cells(i, 9)
                    ws.Range("Q" & 3).Value = Greatest_Decrease
                    ws.Range("Q" & 3).NumberFormat = "0.00%"
                    
                    Else
                    
                    End If
                    
                End If
                
                
                
                If ws.Cells(i, 12).Value > Greatest_Volume Then
                
                'Set value for the greatest increase
                Greatest_Volume = ws.Cells(i, 12).Value
                     
                'Print the greates increase
                ws.Range("P" & 4).Value = ws.Cells(i, 9)
                ws.Range("Q" & 4).Value = Greatest_Volume
                    
                Else
                
                End If
         
        Next i
        
        
        ' Adjusts the column width
        ws.Columns("A:Q").AutoFit
                                        
            
    Next ws
    
    MsgBox ("Analysis complete")


End Sub
