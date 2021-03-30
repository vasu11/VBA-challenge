Attribute VB_Name = "Module2"
Sub Tickers()

  For Each ws In Worksheets
        ' Determine the Last Row
        Dim i, countTicker, printRow As Integer
        Dim ticker, prevTicker As String
        Dim volume, openingPrice, closingPrice, priceChange, percentChange As Double
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        MsgBox LastRow
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        printRow = 2
        prevTicker = " "
        For i = 2 To LastRow
           ticker = ws.Cells(i, 1).Value
           If ticker = prevTicker Or i = 2 Then
              countTicker = countTicker + 1
              If countTicker = 1 Then
                openingPrice = ws.Cells(i, 3).Value
              End If
              volume = volume + ws.Cells(i, 7).Value
           Else
              closingPrice = ws.Cells(i - 1, 6).Value
              priceChange = openingPrice - closingPrice
              percentChange = Round((priceChange / openingPrice) * 100, 2)
              ws.Cells(printRow, 9).Value = prevTicker
              ws.Cells(printRow, 10).Value = priceChange
              ws.Cells(printRow, 11).Value = percentChange
              ws.Cells(printRow, 12).Value = volume
              printRow = printRow + 1
              countTicker = 0
              volume = 0
              priceChange = 0
              closingPrice = 0
              openingPrice = 0
              percentChange = 0
           End If
           
           prevTicker = ticker
           
        Next i
        
        
  Next ws

End Sub
