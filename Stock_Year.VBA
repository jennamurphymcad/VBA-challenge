Sub Wkst_Loop_test()

'Set as intitial variable for holiding Ticker Symbol
Dim ticker_symbol As String

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

lastRowState = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row - 1

For Each ws In Worksheets
    
    

    For I = 2 To lastRowState
    
        If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
                            
                'set ticker symbol
                ticker_symbol = ws.Cells(I, 1).Value
                
                'Add to the Stock Vol Total
                stock_vol = stock_vol + ws.Cells(I, 7).Value
                
                Range("I" & Summary_Table_Row).Value = ticker_symbol
                Range("L" & Summary_Table_Row).Value = stock_vol
                
                Summary_Table_Row = Summary_Table_Row + 1
                
            Else
                'add to the stock_vol
                stock_vol = stock_vol + ws.Cells(I, 7).Value
                
            End If
            
        Next I
    
    Next ws
    
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percentage Change"
    Range("L1") = "Stock Volume"
    
    
End Sub