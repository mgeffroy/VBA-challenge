Sub Stockmarket():
' Things needed in the homework
        'The ticker symbol.
        'Yearly change from opening price to closing price at the end of that year.
        'The percent change from opening price to the closing price at the end of that year.
        'The total stock volume of the stock.

'Things have to be valid for all spreadsheets

'Plan of action
'Loop for all spreadsheets

For Each ws In Worksheets

'Variables: name, type,value (depending on variable)
Dim Ticker As String

Dim Yearlychange As Double
Yearlychange = 0
'Yearly change = Opening_Val - Closing_Val'

Dim Opening As Double
Opening = ws.Cells(2, 3).Value

Dim Closing As Double
'Could be Closing = ws.Cells(2,6).Value but it has to be the last value of the ticker

Dim Percentchange As Double
'Yearlychange in percent

Dim ToTvol As LongLong
ToTvol = 0


'Summary table to keep track of tickers
Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  



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
 Dim lastrow As Long
  lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  ' Loop through all the tickers
  For i = 2 To lastrow

    ' If same ticker .....
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set ticker name
      Ticker = ws.Cells(i, 1).Value

      ' Add to the Total Volume
      ToTvol = ToTvol + ws.Cells(i, 7).Value

      'Add closing price value
      Closing = ws.Cells(i, 6).Value
      'Calculate Yearly Change
      Yearlychange = Closing - Opening
      'Calculate Change percentage
      'Probem with PLANT :z
      Percentchange = Opening / Yearlychange

      ' Print ticker in summary table
     ws.Range("I" & Summary_Table_Row).Value = Ticker

      ' Print the total volume in table
      ws.Range("L" & Summary_Table_Row).Value = ToTvol

      'Print the Yearly change in table
      ws.Range("J" & Summary_Table_Row).Value = Yearlychange

      'Print the percent change in table
      ws.Range("K" & Summary_Table_Row).Value = Percentchange
      ws.Range("K" & Summary_Table_Row).NumberFormat = "00.00%"
    
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the total volume to prepare it for next ticker
      ToTvol = 0

    ' If immediate cell is same ticker
    Else

      ' Add to the total volume
      ToTvol = ToTvol + ws.Cells(i, 7).Value

    End If
    'Give formatting to Yearly Change
'Red if it is negative numbers and Green if it is positive number
If ws.Cells(i, 10).Value > 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 4
ElseIf ws.Cells(i, 10).Value < 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 3
Else
ws.Cells(i, 10).Interior.ColorIndex = 2
End If

  Next i

Next ws

End Sub
  