Sub Stock()

'Loop through all the worksheets
    For Each ws In ThisWorkbook.Worksheets

'set all the variables

    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Percentage_Change As Double
    Dim Open_price As Double
    Dim Close_price As Double
    Dim Total As Double
    Dim lastrow As String
    
       
'set the total, open_price, close_price & percentage_change to start from 0

    Total = 0
    Open_price = 0
    Close_price = 0
    Yearly_Change = 0
    Percentage_Change = 0
    
'define the summary table

    Dim summary_table_row As Integer
    summary_table_row = 2
    
'allocate cell names
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
        
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
'create a loop
    
    For i = 2 To lastrow
    
'Grabbing the Open_price
    
    If ws.Cells(i, 1) <> ws.Cells(i - 1, 1).Value Then
    Open_price = ws.Cells(i, 3).Value
         
    End If
       
                                 
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    Ticker = ws.Cells(i, 1).Value
    Total = Total + ws.Cells(i, 7).Value
    Close_price = ws.Cells(i, 6).Value
    Yearly_Change = Close_price - Open_price
    Percentage_Change = (Close_price - Open_price) / Close_price
    
    ws.Range("I" & summary_table_row).Value = Ticker
    ws.Range("J" & summary_table_row).Value = Yearly_Change
    
    'create the conditional formatting for yearly_change
    If ws.Range("J" & summary_table_row).Value >= 0# Then
    ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
        
    Else
    ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
        
    End If
    
    ws.Range("K" & summary_table_row).Value = Percentage_Change
    ws.Range("L" & summary_table_row).Value = Total
    summary_table_row = summary_table_row + 1
    Total = 0
    
    Else
    Total = Total + ws.Cells(i, 7).Value
    
    End If
    
              
    Next i
    
    'create second summary table
    
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
          
           
    ws.Cells(2, 17).Value = Application.WorksheetFunction.Max(ws.Range("K:K").Value)
    Increased_ticker = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
    ws.Cells(2, 16).Value = ws.Cells(Increased_ticker + 1, 9)
    
        
    ws.Cells(3, 17).Value = Application.WorksheetFunction.Min(ws.Range("K:K").Value)
    Decreased_ticker = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
    ws.Cells(3, 16).Value = ws.Cells(Decreased_ticker + 1, 9)
    
    
    ws.Cells(4, 17).Value = Application.WorksheetFunction.Max(ws.Range("L:L").Value)
    Max_Volume = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & lastrow)), ws.Range("L2:L" & lastrow), 0)
    ws.Cells(4, 16).Value = ws.Cells(Max_Volume + 1, 9)
    
      
    'Format the numbers and autofit cells
    ws.Range("K:K").NumberFormat = "0.00%"
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Cells(4, 17).NumberFormat = "0"
    
    ws.Columns("I:Q").EntireColumn.AutoFit
                    
    
    
    Next ws
    
    
End Sub






