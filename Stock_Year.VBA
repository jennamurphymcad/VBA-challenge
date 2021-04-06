Sub Wkst_Loop_test()

'Set as intitial variable for holiding Ticker Symbol
Dim ticker_symbol As String
Dim open_price As Double
Dim close_price As Double
Dim date_val As Long
date_val = Application.WorksheetFunction.Min(Range("B:B"))
Dim year_change As Double
Dim percent_change As Double
Dim stock_vol As LongLong
Dim Summary_Table_Row As Double
Summary_Table_Row = 2

'find last row of sheet
last_row = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row - 1

'loops through all tabs of current sheet
'For Each ws In Worksheets
    
    'Loop through active sheet
    For I = 2 To last_row
        
            'find open price
            If Cells(I, 2).Value = date_val Then
                         
                    open_price = Cells(I, 3).Value
               
            End If
         
            'find where ticker symbol changes
            If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
                                
                    'set ticker symbol
                    ticker_symbol = Cells(I, 1).Value
                    
                    'set close price
                    close_price = Cells(I, 6).Value
                   
                    'calculate year change
                    year_change = close_price - open_price
                    
                    'add year change to summary table
                    Range("J" & Summary_Table_Row).Value = year_change
                    
                    'calculate precentage change
                    If open_price <> 0 Then
                    
                        percent_change = (year_change / open_price)
                        Range("K" & Summary_Table_Row).Value = percent_change
                        Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                    
                    Else
                    
                        percent_change = nbNullstring
                        Range("K" & Summary_Table_Row).Value = percent_change
                        
                    End If
                    
                        
                    'Conditional Formatting for Year Change Price
                    If Range("J" & Summary_Table_Row).Value > 0 Then
                        Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                    ElseIf Range("J" & Summary_Table_Row).Value <= 0 Then
                        Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                    End If
                    
                    'Conditional Formatting for Percent Change Price
                    If Range("K" & Summary_Table_Row).Value > 0 Then
                        Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
                    ElseIf Range("K" & Summary_Table_Row).Value <= 0 Then
                        Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
                    End If
                    
                    'Add to the Stock Vol Total
                    stock_vol = stock_vol + Cells(I, 7).Value
                   
                   'Add Ticker Symbol to Summary Table
                    Range("I" & Summary_Table_Row).Value = ticker_symbol
                    
                    'Add Stock Vol to Summary Table
                    Range("L" & Summary_Table_Row).Value = stock_vol
                    
                    'Resize Stock Vol Column so numbers show entirely
                    Range("L" & Summary_Table_Row).Columns.AutoFit
                    
                    'Add additional row
                    Summary_Table_Row = Summary_Table_Row + 1
                                       
                    
                Else
                
                    'add to the stock_vol
                    stock_vol = stock_vol + Cells(I, 7).Value
                    
                End If
                       
            Next I
    
      'Next ws
    
    'Set Headers
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percentage Change"
    Range("L1") = "Stock Volume"
    
End Sub