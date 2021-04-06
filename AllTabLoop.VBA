Sub Wkst_Loop_test()

'Set as intitial variable for holiding Ticker Symbol
Dim ticker_symbol As String

'set as initial variable for holding open price
Dim open_price As Double

Dim close_price As Double

'Set as initial variable for holding close price

Dim date_val As Long
date_val = Application.WorksheetFunction.Min(Range("B:B"))


'set as initial variable for holding Yearly Change
Dim year_change As Double
year_change = 0

'set as initial variable for holding Percent Change
Dim percent_change As Double
percent_change = 0

'set as initial variable for holding Total Stock Volume
Dim stock_vol As LongLong
stock_vol = 0

' Keep track of the location for each credit card brand in the summary table
Dim Summary_Table_Row As Double
Summary_Table_Row = 2

WS_Count = ActiveWorkbook.Worksheets.Count

last_row = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row - 1


'For Each ws In Worksheets

    For I = 2 To last_row
    
        'open_price = ws.Cells(I, 3).Value
        'close_price = ws.Cells(I, 6).Value
        'date_val = ws.Cells(I, 2).Value
        
            If ws.Cells(I, 2).Value = date_val Then
                         
                    open_price = ws.Cells(I, 3).Value
               
            End If
         
       
            If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
                                
                    'set ticker symbol
                    ticker_symbol = ws.Cells(I, 1).Value
                    
                    'year_change = close_price - open_price
                    'Range("J" & Summary_Table_Row).Value = year_change
                    
                    '
                    close_price = ws.Cells(I, 6).Value
                    'Range("J" & Summary_Table_Row).Value = close_price
                    
                    year_change = close_price - open_price
                    Range("J" & Summary_Table_Row).Value = year_change
                    
                    If open_price <> 0 Then
                    
                        percent_change = (year_change / open_price) * 100
                        Range("K" & Summary_Table_Row).Value = percent_change
                    
                    Else
                        percent_change = 0
                        Range("K" & Summary_Table_Row).Value = percent_change
                        
    
                    End If
                    'Range("K" & Summary_Table_Row).Value = open_price
                        
                    
                    'Add to the Stock Vol Total
                    stock_vol = stock_vol + ws.Cells(I, 7).Value
                    'year_change = year_change + close_price - open_price
        
                    'Range("J" & Summary_Table_Row).Value = year_change
                    Range("I" & Summary_Table_Row).Value = ticker_symbol
                    Range("L" & Summary_Table_Row).Value = stock_vol
                    
                    Summary_Table_Row = Summary_Table_Row + 1
                                       
                    
                Else
                
                    'add to the stock_vol
                    stock_vol = stock_vol + ws.Cells(I, 7).Value
                    
                End If
                

            
            Next I
    
      'Next ws
    
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percentage Change"
    Range("L1") = "Stock Volume"
    
End Sub
