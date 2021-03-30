Attribute VB_Name = "Module2"
Sub Tickers()

  For Each ws In Worksheets
        ' Determine the Last Row
        Dim i, countTicker, printRow As Integer
        Dim ticker, prevTicker As String
        Dim volume, openingPrice, closingPrice, priceChange, percentChange, tickerValueChange As Double
        Dim WorksheetName As String
        
        WorksheetName = ws.Name
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ' MsgBox LastRow
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        printRow = 2
        countTicker = 0
        prevTicker = " "
        For i = 2 To LastRow
           ticker = ws.Cells(i, 1).Value
           If ticker = prevTicker Or i = 2 Then
              countTicker = countTicker + 1
              If i = 2 Then
                 openingPrice = ws.Cells(i, 3).Value
              End If
              If countTicker = 1 And i <> 2 Then
                 openingPrice = ws.Cells(i - 1, 3).Value
              End If
              volume = volume + ws.Cells(i, 7).Value
           Else
              closingPrice = ws.Cells(i - 1, 6).Value
              ' If ticker = "AA" Then
              '   MsgBox i
              '   MsgBox closingPrice
              '   MsgBox openingPrice
              ' End If
              priceChange = closingPrice - openingPrice
              If openingPrice <> 0 Then
                 tickerValueChange = priceChange / openingPrice
                 percentChange = FormatPercent(tickerValueChange, 2)
              End If
              ws.Cells(printRow, 9).Value = prevTicker
              ws.Cells(printRow, 10).Value = priceChange
              ws.Cells(printRow, 11).Value = percentChange
              ws.Cells(printRow, 12).Value = volume
              If priceChange > 0 Then
                 ws.Cells(printRow, 10).Interior.ColorIndex = 4
              Else
                 ws.Cells(printRow, 10).Interior.ColorIndex = 3
              End If
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
