Sub Stockmarket():
' Things needed in the homework
        'The ticker symbol.
        'Yearly change from opening price to closing price at the end of that year.
        'The percent change from opening price to the closing price at the end of that year.
        'The total stock volume of the stock.

'Things have to be valid for all spreadsheets

'Plan of action

'Variables: name, type,value (depending on variable)
Dim Ticker As String

Dim Yearlychange As Double
Yearlychange = 0

Dim Opening_val As Double

Dim Closing_val As Double

Dim Percentchange As Double

Dim ToTvol As LongLong
Totvol = 0


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
ws.Columns("G").Columns.AutoFit
ws.Columns("B").Columns.AutoFit
ws.Range("A1:L1").HorizontalAlignment = xlCenter

'Create lastrow variable
 Dim lastrow as Long  
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  
  ' Loop through all the tickers 
  For i = 2 To lastrow

    ' If same ticker ..... if not 
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set ticker name
      Ticker = Cells(i, 1).Value

      ' Add to the Total Volume
      TotVol = TotVol + Cells(i, 7).Value

      ' Print ticker in summary table
      Range("I" & Summary_Table_Row).Value = Ticker

      ' Print the total volume in table 
      Range("L" & Summary_Table_Row).Value = TotVol

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the total volume to prepare it for next ticker 
      TotVol = 0

    ' If immediate cell is same ticker
    Else

      ' Add to the total volume
      TotVol= TotVol + Cells(i, 7 ).Value

    End If

  Next i

Next ws

End Sub
