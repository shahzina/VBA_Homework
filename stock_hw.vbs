Sub stocktotals()

Dim ws As Worksheet
Dim cur_stockname, next_stockname, hold_stockname, wsname As String
Dim cur_stockval, next_stockval, hold_stockval, LastRow, result_row As Long
Dim count As Long
Dim stock_total As Double

count = 0


For Each ws In Worksheets
    
    stock_total = 0
    wsname = ws.Name
    MsgBox ("worksheet " + wsname)
    
    'Last row number
    LastRow = ws.Cells(Rows.count, 1).End(xlUp).Row
    
    'Populate titles for I and J.
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Total Volume"
    
    result_row = 2
    For i = 2 To LastRow
    
  
    
        count = count + 1
        
        If i = 2 Then
           hold_stockname = ws.Range("A2").Value
       End If
        
        'assign values of ticker and total stock trades from cell values to variables.
        cur_stockname = ws.Cells(i, 1).Value
        cur_stockval = ws.Cells(i, 7).Value
        next_stockname = ws.Cells(i + 1, 1).Value
        next_stockval = ws.Cells(i + 1, 7).Value
        
        If cur_stockname = next_stockname Then
            
            stock_total = stock_total + cur_stockval
        
        Else
            stock_total = stock_total + cur_stockval
            ' Print the ticker name  in the Result Table.
            ws.Range("I" & result_row).Value = hold_stockname
            
            ' Print the Total stock value into the Result Table.
            ws.Range("J" & result_row).Value = stock_total
            
            stock_total = 0
            hold_stockname = next_stockname
            result_row = result_row + 1
            
        
        End If
        
    
    Next i
            
    

Next ws









End Sub

