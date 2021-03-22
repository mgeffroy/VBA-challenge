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

Dim Yearlychange As Double
Yearlychange = 0

Dim Opening_val As Double

Dim Closing_val As Double

Dim Percentchange As Double

Dim Totstockvol As Long
Totstockvol = 0


'Summary table to keep track of tickers
Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  

'Loop for all spreadsheets

For Each ws In Worksheets

'Name headers in spreadsheet ( with a little bit of formatting)

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Range("A1:L1").Columns.AutoFit
ws.Range("A1:L1").HorizontalAlignment = xlCenter

Next ws

End Sub

End Sub
Sub bonus(): 

'Make it work in all spreadsheets


'Create variables 
Dim Grt_increase as long 
Dim Grt_decrease as long 
Dim Grt_totvol as long 