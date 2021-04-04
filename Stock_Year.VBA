Sub Wkst_Loop_test()

'Set as intitial variable for holiding Ticker Symbol
Dim ticker_symbol As String

'set as initial variable for holding Yearly Change
Dim year_change as Double
year_change = 0

'set as initial variable for holding Percent Change
Dim percent_change as Double
percent_change = 0

'set as initial variable for holding Total Stock Volume
Dim stock_vol as LongLong
stock_vol = 0

' Keep track of the location for each credit card brand in the summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2


'testing just first tab summary
For i = 2 To 70926
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        ticker_symbol = Cells(i, 1).Value

        Range("I" & Summary_Table_row).Value = ticker_symbol

        Summary_Table_Row = Summary_Table_Row + 1

    End If
Next I
End Sub








Dim WS_Count As Integer
Dim I As Integer

       ' Set WS_Count equal to the number of worksheets in the active
       ' workbook.
       WS_Count = ActiveWorkbook.Worksheets.Count

       ' Begin the loop.
       For I = 1 To WS_Count

          ' Insert your code here.
          ' The following line shows how to reference a sheet within
          ' the loop by displaying the worksheet name in a dialog box.
          MsgBox ActiveWorkbook.Worksheets(I).Name

       Next I

End Sub