Attribute VB_Name = "Module1"
Sub StockDataSummary():
' variable for storing the current workbook data
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
ws.Activate


        ' declare variables
        ' variable for storing the ticker name
        Dim ticker As String
        
        ' variable for storing quarterly change
        Dim quarterly_change As Double
        
        ' variable for storing percentage change
        Dim percentage_change As Double
        
        ' variable for storing stock open date
        'Dim open_date As Integer
        
        ' variable for storing open price
        Dim open_price As Double
        
        ' variable for storing closing price
        Dim close_price As Double
        
        ' variable for storing low price
        Dim low_price As Double
        
        ' variable for storing high price
        Dim high_price As Double
        
        ' variable for storing total volume
        Dim total_volume As Long
        total_volume = 0
        
        ' variables for tracking summary table position
        Dim summary_table_row As Integer
        summary_table_row = 2
          
        
        ' add headers for the summary table
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Quarterly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        Range("O2").Value = "Greatest % increase"
        Range("O3").Value = "Greatest % decrease"
        Range("O4").Value = "Greatest total volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
        ' set last row
        Dim lastRow As Long
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
        ' MsgBox (lastRow)
        
        ' For Index = 2 To 300
        For Index = 2 To lastRow
                ' check if the current ticker and the previous ticker match, if not set opening price of the new ticker
                If Cells(Index - 1, 1) <> Cells(Index, 1) Then
                    open_price = Cells(Index, 3).Value
                    ' MsgBox (open_price)
                
                ' check if the next ticker and current ticker matc, if not set closing price of the new ticker
                ElseIf Cells(Index + 1, 1).Value <> Cells(Index, 1).Value Then
                    ' set new ticker name
                    ticker = Cells(Index, 1).Value
                    ' set the closing price
                    close_price = Cells(Index, 6).Value

                    ' update the total volume
                    total_volume = total_volume + Cells(Index, 7).Value
                    'MsgBox (total_volume)
                    
                    ' calculate the difference between open price and close price
                    quarterly_change = open_price - close_price
                    'MsgBox (quarterly_change)
                    
                    ' calculate the percentage change of the stock
                    'If open_price = 0 Then
                        'percentage_change = 0
                    'Else
                        percentage_change = (open_price - close_price) / open_price
                        On Error Resume Next
                    
                    ' add the ticker in the summary table
                    'MsgBox ("Setting Summary Table Ticker")
                    Cells(summary_table_row, 9).Value = ticker
      
                    ' add quarterly change to the summary table
                    'MsgBox ("Setting Summary Table quarterly_change")
                    Cells(summary_table_row, 10).Value = quarterly_change
      
                    ' add the percentage change to the summary table
                    'MsgBox ("Setting Summary Table percent_change")
                    Cells(summary_table_row, 11).Value = percentage_change
                    Columns("K:K").NumberFormat = "0.00%"

                    ' add the ticker volume to the summary table
                    'MsgBox ("Setting Summary Table total_volume")
                    Cells(summary_table_row, 12).Value = total_volume
      
                    ' update the summary table row
                    summary_table_row = summary_table_row + 1
      
                    ' reset the total volume
                    total_volume = 0
                    
            ' if ticker is the same
            Else
                ' update total volume of the stock
                total_volume = total_volume + Cells(Index, 7).Value

            End If
            

       
        Next Index
        
        ' loop for finding greatest values
        Dim greatest_increase As Double
        Dim greatest_decrease As Double
        Dim greatest_volume As Double
        
        greatest_increase = Cells(2, 11).Value
        greatest_decrease = Cells(2, 11).Value
        greatest_volume = Cells(2, 12).Value
        final_row = Cells(Rows.Count, 10).End(xlUp).Row
  
       For Row = 2 To final_row
        ' update color of the cell based on value
        If Cells(Row, 10) < 0 Then
            Cells(Row, 10).Interior.ColorIndex = 3
            
        ElseIf Cells(Row, 10) >= 0 Then
            Cells(Row, 10).Interior.ColorIndex = 4
   
        End If
        
         'check for the greatest increase value
        If Cells(Row, 11) > greatest_increase Then
            greatest_increase = Cells(Row, 11)
            Cells(2, 16) = Cells(Row, 9)
            Cells(2, 17) = greatest_increase
            Cells(2, 17).NumberFormat = "0.00%"
            
   
        End If
   
        ' check for greatest decrease value
        If Cells(Row, 11) < greatest_decrease Then
            greatest_decrease = Cells(Row, 11)
            Cells(3, 16) = Cells(Row, 9)
            Cells(3, 17) = greatest_decrease
            Cells(3, 17).NumberFormat = "0.00%"
            
  
        End If
        
        ' check for the greatest total volume
        If Cells(j, 12) > greatest_volume Then
            greatest_volume = Cells(j, 12)
            Cells(4, 16) = Cells(j, 9)
            Cells(4, 17) = greatest_volume
            
   
        End If
   
       Next Row

Next
MsgBox ("done")
End Sub

