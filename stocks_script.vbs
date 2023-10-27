Attribute VB_Name = "Module1"
Sub stocks()

'loop through all sheets
For Each ws In Worksheets
    
    'name columns and rows where we will add the information
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percentage Change"
    ws.Range("L1") = "Total Stock Volume"

    ws.Range("O2") = "Greatest % increase"
    ws.Range("O3") = "Greatest % decrease"
    ws.Range("O4") = "Greatest total volume"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    
    'set row to start adding information
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    'find last row
    Dim LastRow As Long
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'define variables
    Dim ticker As String
    Dim vol As Double
    Dim price_open As Double
    Dim price_close As Double
    
    vol = 0
    
    'attribute the first value to open price, since the loop below won't get that value
    price_open = ws.Cells(2, 3).Value

    'loop through all rows to find the values we need
    For i = 2 To LastRow
        'if the next cell is different to the current one, then ticker has changed, so we need to get most of the key values here
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'create a list of tickers
            ticker = ticker + ws.Cells(i, 1).Value
            
            'total volume for each stock
            vol = vol + ws.Cells(i, 7).Value
            
            'get closing price
            price_close = ws.Cells(i, 6).Value
                
            'add ticker to summary table
            ws.Range("I" & Summary_Table_Row).Value = ticker
        
            'add yearly change to summary table
            ws.Range("J" & Summary_Table_Row).Value = price_close - price_open
        
            'add percent change to summary table
            ws.Range("K" & Summary_Table_Row).Value = (price_close - price_open) / price_open
        
            'add total stock volume to summary table
            ws.Range("L" & Summary_Table_Row).Value = vol

            'add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
        
            'reset ticker and total vol
            ticker = ""
            vol = 0
        
            'get price_open for new ticker
            price_open = ws.Cells(i + 1, 3).Value

        'if the next cell is different to the current one then we only want to add to the total volume
        Else
            'add to the stock total
            vol = vol + ws.Cells(i, 7).Value
        End If
    
    Next i

    'calculate greatest % increase, % decrease and volume
    max_inc = Application.WorksheetFunction.Max(ws.Range("K:K"))
    max_dec = Application.WorksheetFunction.Min(ws.Range("K:K"))
    max_vol = Application.WorksheetFunction.Max(ws.Range("L:L"))
   
    'find lastTicker row
    Dim LastTicker As Long
    LastTicker = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'loop through values for each ticker to find greatest % increase, % decrease and volume
    For i = 2 To LastTicker
        'find greatest % increase and copy to correct cell
        If ws.Cells(i, 11) = max_inc Then
            ws.Range("P2") = ws.Cells(i, 9)
            ws.Range("Q2") = ws.Cells(i, 11)
        
        'find greatest % decrease and copy to correct cell
        ElseIf ws.Cells(i, 11) = max_dec Then
            ws.Range("P3") = ws.Cells(i, 9)
            ws.Range("Q3") = ws.Cells(i, 11)
        End If
                
        'find greatest total volume and copy to correct cell
        If ws.Cells(i, 12) = max_vol Then
             ws.Range("P4") = ws.Cells(i, 9)
             ws.Range("Q4") = ws.Cells(i, 12)
        End If
               
        'conditional formating
        If ws.Cells(i, 10) < 0 Then
        'negative values should be highlighted in red
            ws.Cells(i, 10).Interior.ColorIndex = 3
            ws.Cells(i, 11).Interior.ColorIndex = 3
        
        Else
        'positive values should be highlighted in green
            ws.Cells(i, 10).Interior.ColorIndex = 43
            ws.Cells(i, 11).Interior.ColorIndex = 43
        End If
    
    Next i

Next ws

End Sub
