Sub quarter_change()

    Dim ws As Worksheet
    For Each ws In Worksheets
    
        Dim row As Long
        Dim resultrow As Long
            resultrow = 2
        Dim total As Double
            total = 0
        Dim ticker As String
        
    'Add column labels for results table
    
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        RowCount = Cells(Rows.Count, "A").End(xlUp).row
        
    'Create results using for loop
        For row = 2 To RowCount
        
    'Assign variables for calculations
            Dim openval As Double
            Dim change As Double
            
    'Add to total volume
            total = total + ws.Cells(row, 7).Value
            
     'Store open value for first iteration of each ticker
                If ws.Cells(row - 1, 1).Value <> ws.Cells(row, 1).Value Then
                    openval = ws.Cells(row, 3).Value
                End If
                
     'Printing Ticker and Quarterly Change in results table
                If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
                    change = ws.Cells(row, 6).Value - openval
                    ws.Cells(resultrow, 9).Value = ws.Cells(row, 1).Value
                    ws.Cells(resultrow, 10).Value = change
                    ws.Cells(resultrow, 10).NumberFormat = "0.00"
                    
      'Conditional formatting of Quarterly Change
                    If change > 0 Then
                        ws.Cells(resultrow, 10).Interior.ColorIndex = 4
                    ElseIf change < 0 Then
                        ws.Cells(resultrow, 10).Interior.ColorIndex = 3
                    Else: ws.Cells(resultrow, 10).Interior.ColorIndex = 0
                    End If
                                                                  
      'Calculating and printing Percent Change
                    ws.Cells(resultrow, 11).Value = change / openval
                    ws.Cells(resultrow, 11).NumberFormat = "0.00%"
        'Conditional formatting of Percent Change
                    If ws.Cells(resultrow, 11).Value > 0 Then
                        ws.Cells(resultrow, 11).Font.ColorIndex = 4
                    ElseIf ws.Cells(resultrow, 11).Value < 0 Then
                        ws.Cells(resultrow, 11).Font.ColorIndex = 3
                    Else: ws.Cells(resultrow, 11).Font.ColorIndex = 1
                    End If
       'Printing Total Stock volume
                    ws.Cells(resultrow, 12).Value = total
                    
        'Reseting Total Stock Volume and advancing resultrow
                    total = 0
                    resultrow = resultrow + 1
                End If
        Next row
        
        'Add Column and Row Labels for final table
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
            ws.Cells(1, 16).HorizontalAlignment = xlCenter
            
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        
        'Assigning variables to store results for final table
        Dim bigticker As String
        Dim bigtickernum As Double
            bigtickernum = 0
        Dim smallticker As String
        Dim smalltickernum As Double
            smallticekernum = 0
        Dim fatticker As String
        Dim fattickernum As Double
            fattickernum = 0
            
        ResultRowCount = Cells(Rows.Count, "I").End(xlUp).row
        
        'Create final table using For loop
        For row = 2 To ResultRowCount
                    If ws.Cells(row, 11).Value > bigtickernum Then
                        bigticker = ws.Cells(row, 9).Value
                        bigtickernum = ws.Cells(row, 11).Value
                    ElseIf ws.Cells(row, 11).Value < smalltickernum Then
                        smallticker = ws.Cells(row, 9).Value
                        smalltickernum = ws.Cells(row, 11).Value
                    ElseIf ws.Cells(row, 12).Value > fattickernum Then
                        fatticker = ws.Cells(row, 9).Value
                        fattickernum = ws.Cells(row, 12).Value
                    End If
            Next row
        'Print and format final table
            ws.Cells(2, 15).Value = bigticker
            ws.Cells(2, 16).Value = bigtickernum
                ws.Cells(2, 16).NumberFormat = "0.00%"
            ws.Cells(3, 15).Value = smallticker
            ws.Cells(3, 16).Value = smalltickernum
                ws.Cells(3, 16).NumberFormat = "0.00%"
            ws.Cells(4, 15).Value = fatticker
            ws.Cells(4, 16).Value = fattickernum
   'Reset stored values
   bigtickernum = 0
   smalltickernum = 0
   fattickernum = 0
   Next ws
   
End Sub
