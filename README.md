# VBA-challenge
Sub Yearstockdata()
For Each Ws In Worksheets

Dim i As Long
Dim ticker_acronyms As Double



Dim ticker As String
Dim lastrow As Long
Dim open_price As Double
Dim percent_change As Double
Dim total_stock As Double
Dim close_price As Double

ticker_acronyms = 1


Ws.Cells(2, 9).Value = Ws.Cells(2, 1).Value
open_price = Cells(2, 3).Value

total_stock = 0


lastrow = Ws.Cells(Rows.Count, 1).End(xlUp).Row

Ws.Range("i1").Value = "Ticker"
Ws.Range("j1").Value = "Yearly Change"
Ws.Range("k1").Value = "Percentage change"
Ws.Range("l1").Value = "Total Stock Volume"

For i = 2 To lastrow
If Ws.Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
ticker = Ws.Cells(i + 1, 1).Value
Ws.Cells(ticker_acronyms + 2, 9).Value = ticker
close_price = Ws.Cells(i, 6).Value
Ws.Cells(ticker_acronyms + 1, 10).Value = close_price - open_price
percent_change = ((close_price - open_price) / open_price)
Ws.Cells(ticker_acronyms + 1, 11).Value = percent_change
Ws.Cells(ticker_acronyms + 1, 12).Value = total_stock
open_price = Ws.Cells(i + 1, 3).Value
ticker_acronyms = ticker_acronyms + 1

Else: total_stock = total_stock + Ws.Cells(i, 7).Value



End If
Next i
For i = 2 To lastrow
If Ws.Cells(i, 10).Value > 0 Then
Ws.Cells(i, 10).Interior.ColorIndex = 4
ElseIf Ws.Cells(i, 10).Value < 0 Then
Ws.Cells(i, 10).Interior.ColorIndex = 3
ElseIf Ws.Cells(i, 10).Value = 0 Then



End If
Next i
Next Ws

End Sub

