Attribute VB_Name = "Module1"
'## Instructions

'Create a script that loops through all the stocks for one year and outputs the following information:

  '* The ticker symbol.

  '* Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  '* The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  '* The total stock volume of the stock.

'**Note:** Make sure to use conditional formatting that will highlight positive change in green and negative change in red.

Sub AnalyzeStocksData()
   
    For Each ws In Worksheets

        Dim WorksheetName As String
        'Get the WorksheetName
        WorksheetName = ws.Name

        'Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Variable Declaration

        Dim totalStockVolume As LongLong
        Dim ticker As Long
        Dim openPrice As Double
        Dim closingPrice As Double
        Dim yearlyChange As Double
        Dim percentChange As Double
        Dim greatestIncrease As Double
        Dim greatestDecrease As Double
        Dim greatesTotVol As LongLong



        'Label column header (Ticker, Yearly Change, Percent Change, Total Stock Volume)
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"


        'Label column header Ticker, Value
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

        'Label row  Greatest % Increase, Greatest % Decrease,Greatest Total Volume
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"


        'Initialize variables
        ticker = 2
        greatestIncrease = 0
        greatestDecrease = 0
        greatesTotVol = 0
        openPrice = ws.Range("C2").Value



        For i = 2 To LastRow

            totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                closingPrice = ws.Cells(i, 6).Value
                ws.Cells(ticker, 9).Value = ws.Cells(i, 1).Value
                yearlyChange = closingPrice - openPrice


                'Use conditional formatting that will highlight positive change in green and negative change in red.
                If yearlyChange < 0 Then
                    ws.Cells(ticker, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(ticker, 10).Interior.ColorIndex = 4
                End If

                ws.Cells(ticker, 10).Value = yearlyChange

                'Compute percentChange
                percentChange = (closingPrice - openPrice) / openPrice
                ws.Cells(ticker, 11).Value = percentChange
                ws.Cells(ticker, 12).Value = totalStockVolume
    
                
                'Return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    tickerPercentIncrease = ws.Cells(i, 1).Value

                ElseIf percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    tickerPercentDecrease = ws.Cells(i, 1).Value
                End If

                If totalStockVolume > greatesTotVol Then
                    greatesTotVol = totalStockVolume
                    greatTotVol = ws.Cells(i, 1).Value
                End If


                totalStockVolume = 0
                openPrice = ws.Cells(i + 1, 3).Value
                ticker = ticker + 1

            End If

        Next i
        
        'Write ticker, value in cells (Greatest % increase, Greatest % decrease, Greatest Total Volume)
        '----------------------------------------------------------
        ws.Cells(2, 16).Value = tickerPercentIncrease
        ws.Cells(2, 17).Value = greatestIncrease
        ws.Cells(3, 16).Value = tickerPercentDecrease
        ws.Cells(3, 17).Value = greatestDecrease
        ws.Cells(4, 16).Value = greatTotVol
        ws.Cells(4, 17).Value = greatesTotVol
       
        '----------------------------------------------------------
        'Format cells values to number format
        ws.Range("K:K").NumberFormat = "0.00%"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
        'Change the column width to automatically fit the contents
        ws.Columns("A:Q").AutoFit
        

    Next ws

End Sub

