Sub Stockmarket():
' Things needed in the homework
        'The ticker symbol.
        'Yearly change from opening price to closing price at the end of that year.
        'The percent change from opening price to the closing price at the end of that year.
        'The total stock volume of the stock.

'Things have to be valid for all spreadsheets

'Plan of action


'Variables: name, type,value (depending on variable)
Dim ticker As String
Dim Totstockvol As LongLong 

'Give starting values 

Totstockvol = 0 



'Summary table to keep track of tickers
Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  

'Loop for all spreadsheets


For Each WS In Worksheets

'Name headers in sheet ( with a little bit of formatting)

WS.Cells(1, 9).Value = "Ticker"
WS.Cells(1, 10).Value = "Yearly Change"
WS.Cells(1, 11).Value = "Percent Change"
WS.Cells(1, 12).Value = "Total Stock Volume"
WS.Range("A1:L1").Columns.AutoFit
WS.Columns("G", "B").Columns.AutoFit
WS.Range("A1:L1").HorizontalAlignment = xlCenter

'Create last row variable
lastrow = WS.Cells(Rows.Count, 1).End(xlUp).Row
'Make loop
For i = 2 To lastrow

'Conditional to find different tickers 
If WS.Cells(i,1).Value <> WS.Cells(i+1, 1).Value Then 

'locating values 
ticker = WS.Cells(i,1).Value 
Totstockvol = Totstockvol + WS.Cells(i,7).Value

'Print values 
WS.Range("I" And Summary_Table_Row).Value = ticker 
WS.Range("L" And Summary_Table_Row).Value = Totstockvol

Summary_Table_Row = Summary_Table_Row + 1

'Reset for next loop and ticker
Totstockvol = 0 

Else 
Totstockvol = Totstockvol + WS.Cells(i,7).Value 

End if 

Next i

Next WS

End Sub


Sub bonus(): 

'Make it work in all spreadsheets


'Create variables 
Dim Grt_increase as long 
Dim Grt_decrease as long 
Dim Grt_totvol as long 