# VBA-Homework-
sub StockHw()

'looping through all worksheets
dim ws as Worksheet


For Each ws in Worksheets

    'ticker symbol and total volume variables
    dim tickerSym as string
    dim totalVol as double

    'year open and close variable
    dim yearOpen as double
    dim yearClose as double

    'summary table row variable
    dim summaryRow as integer
    summaryRow = 2

    'print summary table headers
    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = "Yearly Change"
    ws.Cells(1, 12).Value = "Percent Change"
    ws.Cells(1, 13).Value = "Total Volume"

    'find last row
    dim lastRow as double
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 to lastRow
        'if no opening cost, skip row   
        If (ws.Cells(i, 3).Value = 0) Then
            'if last cell for a ticker symbol with absolutely no data, set ticker symbol
            If (ws.Cells(i + 1).Value <> ws.Cells(i, 1).Value) Then
                'set ticker symbol variable
                tickerSym = ws.Cells(i, 1).Value
            End if
        'if next cell = this cell, add to total volume
        Elseif (ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value) Then
            totalVol = totalVol + ws.Cells(i, 7).Value
            'if last cell <> this cell, set yearOpen
            If (ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value) Then
                yearOpen = ws.Cells(i, 3).Value
            End if
        Else
            'set ticker symbol variable
            tickerSym = ws.Cells(i, 1).Value
            'add to total volume
            totalVol = totalVol + ws.Cells(i, 7).Value
            'set yearClose
            yearClose = ws.Cells(i, 6).Value
            'print ticker symbol and total volume in summary table
            ws.Cells(summaryRow, 10).Value = tickerSym
            ws.Cells(summaryRow, 13).Value = totalVol
            'to avoid dividing by zero
            If (totalVol > 0) Then
                'print yearly change in summary table
                ws.Cells(summaryRow, 11).Value = yearClose - yearOpen
                    'change color to green if > 0, else red
                    If (ws.Cells(summaryRow, 11).Value > 0) Then
                        ws.Cells(summaryRow, 11).Interior.ColorIndex = 4
                    Else
                        ws.Cells(summaryRow, 11).Interior.ColorIndex = 3
                    End if
                'print % change in summary table
                ws.Cells(summaryRow, 12).Value =  ws.Cells(summaryRow, 11).Value / yearOpen
            Else
                'set yearly and % change to zero if no stock data
                ws.Cells(summaryRow, 11).Value = 0
                ws.Cells(summaryRow, 12).Value = 0
            End if
            'set cell format to percent
            ws.Cells(summaryRow, 12).Style = "percent"
            'reset total volume
            totalVol = 0
            'next summary row
            summaryRow = summaryRow + 1
        End If
    Next i

    '''finding greatest total volume
    'greatest total volume variable
    dim greatTotVol as double

    'print table labels
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"

    'set baseline for greatest total volume
    greatTotVol = 0

    ''MsgBox(summaryRow)
    'offset summaryRow to equal number of ticker symbols
    summaryRow = summaryRow - 2
'Find % Change 
    'greatest % increase and decrease variables
    dim increasePerc as double
    dim decreasePerc as double

    'set baseline for greatest & increase and decrease
    increasePerc = 0
    decreasePerc = 0

    For i = 2 to summaryRow
        'if cell > greatest % increase, set cell as greatest % increase 
        If (ws.Cells(i, 12).Value > increasePerc) Then 
            increasePerc = ws.Cells(i, 12).Value

            'print ticker symbol in table
            ws.Cells(2, 16) = ws.Cells(i, 10).Value
        'if cell < greatest % decrease, set cell as greatest % decrease
        Elseif (ws.Cells(i, 12).Value < decreasePerc) Then
            decreasePerc = ws.Cells(i, 12).Value

            'print ticker symbol in table
            ws.Cells(3, 16).Value = ws.Cells(i, 10).Value
        End If
    Next i

    'print greatest % increase and decrease in table
    ws.Cells(2, 17).Value = increasePerc
    ws.Cells(3, 17).Value = decreasePerc

    'set cell format to percent
    ws.Cells(2, 17).Style = "percent"
    ws.Cells(3, 17).Style = "percent"
'Find Total Volume 
    'if cell > greatest total volume, set cell as greatest total volume
    For i = 2 to summaryRow
        If (ws.Cells(i, 13).Value > greatTotVol) Then
            greatTotVol = ws.Cells(i, 13).Value

            'print ticker symbol in table
            ws.Cells(4, 16).Value = ws.Cells(i, 10).Value
        End if
    Next i

    'print greatest total volume in table
    ws.Cells(4, 17).Value = greatTotVol

    'auto fit table columns
    ws.Columns("J:Q").AutoFit

Next ws

End sub

