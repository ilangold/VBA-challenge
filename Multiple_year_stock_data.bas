Attribute VB_Name = "Module1"
'Ilan Goldstein
'December 9, 2024
'Module 2 Challenge
'Part I: Return ticker symbols with quarterly price change, percent change, and total volume
'Part II: Returnticker symbols for stocks with greatest %increase, %decrease, and total volume with respective values

Sub stockTickers()
    'Initialize variables
    Dim i As Long
    Dim ws As Worksheet
    Dim outputRow As Integer
    Dim LastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim totalVol As LongLong
    Dim priceChange As Double
    Dim percentChange As Double
    
    For Each ws In Worksheets
        'Create output section
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Volume"
        totalVol = 0
        openPrice = 0
        closePrice = 0
        outputRow = 2
        
        'Find table length
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Loop through list....
        For i = 2 To LastRow
            'Get open price if first row for ticker
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                openPrice = ws.Cells(i, 3).Value
            End If
            'If the next row is a new ticker....
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'Get close price
                closePrice = ws.Cells(i, 6).Value
                
                'Get ticker name
                ticker = ws.Cells(i, 1).Value
                
                'Add final Daily value
                totalVol = totalVol + ws.Cells(i, 7).Value
                
                'Calculate price change and percent change
                priceChange = closePrice - openPrice
                percentChange = priceChange / openPrice
                
                'Output ticker name, price change, percent change, and total final
                ws.Cells(outputRow, 9).Value = ticker
                ws.Cells(outputRow, 10).Value = priceChange
                ws.Cells(outputRow, 11).Value = percentChange
                ws.Cells(outputRow, 12).Value = totalVol
                
                'Set color codes for price change
                If priceChange > 0 Then
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 4
                ElseIf priceChange < 0 Then
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 3
                End If
                
                'Set next output row and reset total volume
                outputRow = outputRow + 1
                totalVol = 0
                
            'If the next row is the same ticker...
            Else
                'Add daily volume to total volume
                totalVol = totalVol + ws.Cells(i, 7)
            End If
        Next i
    Next ws
End Sub

