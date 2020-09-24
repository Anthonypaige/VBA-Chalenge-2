Sub StockMarket()
    'Loop through all sheets
'Dim ws As Worksheet
Dim Ticker_name As String
Dim Open_Price As Double
Dim Close_Price As Double
Dim Percent_Change As Double
Dim Yearly_Change As Double
Dim Volume As Double
Dim Row As Long, mainrow As Long
​
Dim Lastrow As Long
Dim i As Long
​
For Each ws In Worksheets
​
    'Set header info
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
    'Set variables
​
Volume = 0
Open_Price = 0
Close_Price = 0
Row = 2
mainrow = 2
​
  'Loop through all sheets to find last cell
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
​
    'Loop from Row 2 to last row
For i = 2 To Lastrow
    'check if we are still on the same ticker name
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    'Set variables
        Ticker_name = ws.Cells(i, 1).Value
        Open_Price = ws.Cells(mainrow, 3).Value
        Close_Price = ws.Cells(i, 6).Value
        Yearly_Change = Close_Price - Open_Price
        If Open_Price = 0 Then
            Percent_Change = 0
        Else
            Percent_Change = Yearly_Change / Open_Price
        End If
        
        Volume = Volume + ws.Cells(i, 7).Value
    'Print data in Summary Table
        ws.Range("I" & Row).Value = Ticker_name
        ws.Range("J" & Row).Value = Yearly_Change
        ws.Cells(Row, 11).Value = Percent_Change
        ws.Cells(Row, 12).Value = Volume
        ws.Cells(Row, 11).NumberFormat = "0.00%"
    
    
    
    'Add one to the summary table row
    Row = Row + 1
    mainrow = i + 1
    Volume = 0
    'Total_Stock_Volume = 0
    Else
        Volume = Volume + ws.Cells(i, 7).Value
    
 'Color fill Percent Change, Red for negative, Green for positive
   If Yearly_Change > 0 Then
   ws.Range("J" & Row).Interior.ColorIndex = 4
   ElseIf Yearly_Change <= 0 Then
   ws.Range("J" & Row).Interior.ColorIndex = 3
    End If
    End If
    
Next i
Next ws
End Sub




