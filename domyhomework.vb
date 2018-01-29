'unit 2 homework
'vba verily and voraciously alliterated
'do some fun things to an excel spreadsheet, three different difficulties

''---'
'EASY'
''---'
'DONE1. Loop through data each year and grab total amount of volume each stock had over the year.
'DONE2. Display the ticker name along side the total volume
'DONE3. I column ticker J column total starts at i2 cells(1, 9)

''-----'
'MEDIUM'
''-----'
'DONE1. Loop through the stocks and take out...
'   DONEa. yearly change from opening price to closing price
'   DONEb. percent change from open to close over the year
'   DONEc. total volume
'   DONEd. ticker name (c and d are easy)
'conditional formatting highlights positive in green and negative in red (this is not good in case of color blindness)
'I ticker name J yearly change K percent change L total stock volume

''---'
'HARD'
''---'
'MEDIUM +
'DONE1. locate the stock with 'greatest % increase' 'greatest % decrease' and 'greatest total volume'
'O greatest... labels P Ticker Q Value

'CHALLENGE
'allow script to run on all worksheets at once




Sub doMyHomework():

For Each ws In Worksheets

    'variables: iterers, a few that were giving me overflow errors so I've removed their type until I take the time to look up a better type
    Dim i, j As Integer, stockName As String, stockVolume, columnBottom, startPrice, closePrice As Double, yearlyChange As Double, percentageChange As Double
    Dim greatestPIncrease As Double, greatestPDecrease As Double, greatestTotalVolume, gpiName As String, gpdName As String, gtvName As String


    'set up some variables, stock volume starts at 0, our iterererer for each stock name can start at 2 for location purposes
    stockVolume = 0
    j = 2
    columnBottom = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ws.Cells(1, 9).Value = "Ticker Name"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    startPrice = ws.Cells(2, 3).Value

    'going to need to copy this loop for i = 32767 to columnBottom
    For i = 2 To columnBottom
        stockName = ws.Cells(i, 1).Value
        stockVolume = stockVolume + ws.Cells(i, 7).Value
        If stockName <> ws.Cells(i + 1, 1).Value Then
            closePrice = ws.Cells(i, 6).Value
            yearlyChange = closePrice - startPrice
            If closePrice = 0 Then
                percentageChange = 0
            Else
                percentageChange = yearlyChange / closePrice
            End If

            'conditional formatting done in vba not in excel
            If yearlyChange > 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If

            'put data nicely in the sheet and get ready for the next ticker
            ws.Cells(j, 9).Value = stockName
            ws.Cells(j, 10).Value = yearlyChange
            ws.Cells(j, 11).Value = percentageChange
            ws.Cells(j, 11).Style = "Percent"
            ws.Cells(j, 12).Value = stockVolume
            j = j + 1
            stockVolume = 0
            startPrice = ws.Cells(i + 1, 3).Value
        End If
    Next i

    greatestPIncrease = ws.Cells(2, 11).Value
    gpiName = ws.Cells(2, 9).Value
    greatestPDecrease = ws.Cells(2, 11).Value
    gpdName = ws.Cells(2, 9).Value
    greatestTotalVolume = ws.Cells(2, 12).Value
    gtvName = ws.Cells(2, 9).Value
    For j = 2 To ws.Cells(Rows.Count, 9).End(xlUp).Row
        If greatestPIncrease < ws.Cells(j, 11).Value Then
            greatestPIncrease = ws.Cells(j, 11).Value
            gpiName = ws.Cells(j, 9).Value
        End If
        If greatestPDecrease > ws.Cells(j, 11).Value Then
            greatestPDecrease = ws.Cells(j, 11).Value
            gpdName = ws.Cells(j, 9).Value
        End If
        If greatestTotalVolume < ws.Cells(j, 12).Value Then
            greatestTotalVolume = ws.Cells(j, 12).Value
            gtvName = ws.Cells(j, 9).Value
        End If
    Next j
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(2, 16).Value = gpiName
    ws.Cells(2, 17).Value = greatestPIncrease
    ws.Cells(2, 17).Style = "Percent"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(3, 16).Value = gpdName
    ws.Cells(3, 17).Value = greatestPDecrease
    ws.Cells(3, 17).Style = "Percent"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(4, 16).Value = gtvName
    ws.Cells(4, 17).Value = greatestTotalVolume
    
Next ws

End Sub


