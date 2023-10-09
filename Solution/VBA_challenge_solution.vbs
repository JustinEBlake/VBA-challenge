Sub Stocks():

'Loop through each worksheet
Dim ws As Worksheet

For Each ws In ThisWorkbook.Sheets:

    'Find the amount of rows in each sheet
    Dim row_count As Long
    row_count = ws.Cells(Rows.Count, "A").End(xlUp).Row - 1
    
    'Create headers & Autofit to display data
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest  % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Columns("I:Q").AutoFit
    
    'Change column K to Percent style
    ws.Range("K:K").NumberFormat = "0.00%"
        
    'Keep track of row index after adding each new ticker
    Dim ticker_row_index As Integer
    ticker_row_index = 2
    
    'Declare volume total variable
    Dim volume_total As Variant
    volume_total = 0
    
    'Declare open and closing prices as variables
    Dim open_price As Double
    Dim close_price As Double
    open_price = 0
    close_price = 0
       
    'Loop through each row in the worksheet
    For i = 2 To row_count:
    
        'Check to see if Ticker does not match the next ticker in the row.
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
            'Apply the ticker name in the "i" col.
            ws.Cells(ticker_row_index, 9).Value = ws.Cells(i, 1).Value
    
            'apply volume total to the "L" col.
            volume_total = volume_total + ws.Cells(i, 7).Value
            ws.Cells(ticker_row_index, 12) = volume_total
            
            'Get closing price data from the sixth col.
            close_price = ws.Cells(i, 6).Value
            
            'Calculate yearly change and apply to "J" col.
            Dim yearly_change As Double
            yearly_change = close_price - open_price
            ws.Cells(ticker_row_index, 10) = yearly_change
            
            'Check to see if yearly change is negative
            If ws.Cells(ticker_row_index, 10) < 0 Then
                
                'Make the interior color red
                ws.Cells(ticker_row_index, 10).Interior.ColorIndex = 3
                
            Else
            
                'Make the interior color green
                ws.Cells(ticker_row_index, 10).Interior.ColorIndex = 4
                
            End If
    
            'Calculate Percentage change and apply to "K" col.
            Dim percent_change As Double
            percent_change = (close_price / open_price) - 1
            ws.Cells(ticker_row_index, 11) = percent_change
             
            'Increment row_ticker_index by 1 to add new ticker on cell below.
            'Reset volume_total
            ticker_row_index = ticker_row_index + 1
            volume_total = 0
        
        Else
        
            'If the ticker matches the next one then we add the volume total together
            volume_total = volume_total + ws.Cells(i, 7).Value
            
            'Check if the the ticker does not match the previous ticker's value
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            
                'If it does not, we will grab the open price data from the third col.
                open_price = ws.Cells(i, 3).Value
                
            End If
                   
        End If
               
    Next i
    
    'Find the amount of rows in Col J
    Dim j_row_count As Long
    j_row_count = ws.Cells(Rows.Count, "J").End(xlUp).Row - 1
    
    'Create variables to store data need from columns
    Dim greatest_increase As Variant
    Dim greatest_decrease As Variant
    Dim greatest_volume As Variant
    greatest_increase = 0
    greatest_decrease = 0
    greatest_volume = 0
    
    
    
    'Loop through the newly created data to assign values in col O:Q
    For j = 2 To j_row_count:
    
        'Find greatest volume
        If ws.Cells(j, 12).Value > greatest_volume Then
        
            'Store new value in variable
            greatest_volume = ws.Cells(j, 12).Value
            
            'Write ticker and value to sheet
            ws.Range("Q4").Value = greatest_volume
            ws.Range("P4").Value = Cells(j, 9)
            ws.Columns("Q").AutoFit
               
        End If
        
        'Find greatest % decrease
        If ws.Cells(j, 11).Value < greatest_decrease Then
        
            'Store new value in variable
            greatest_decrease = ws.Cells(j, 11).Value
            
            'Write ticker and value to sheet
            ws.Range("Q3").Value = greatest_decrease
            ws.Range("P3").Value = Cells(j, 9)
            
            'Format value to percentage
            ws.Range("Q3").NumberFormat = "0.00%"
            
        End If
        
        'Find greatest % increase
        If ws.Cells(j, 11).Value > greatest_increase Then
        
            'Store new value in variable
            greatest_increase = ws.Cells(j, 11).Value
            
            'Write ticker and value to sheet
            ws.Range("Q2").Value = greatest_increase
            ws.Range("P2").Value = Cells(j, 9)
            ws.Columns("Q").AutoFit
            
            'Format value to percentage
            ws.Range("Q2").NumberFormat = "0.00%"

        End If
            
    Next j
            
Next ws

End Sub

