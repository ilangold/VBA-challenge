Attribute VB_Name = "Module1"
' Ilan Goldstein
' December 9, 2024
' Module 2 Challenge
' Part I: Return ticker symbols with quarterly price change, percent change, and total volume
' Part II: Returnticker symbols for stocks with greatest %increase, %decrease, and total volume with respective values

Sub stockTickers()
    ' Create output section
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Quarterly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Volume"
    
    ' Initialize looping variables
    Dim i As Long
    Dim ws As Worksheet
    Dim outputRow As Integer
    outputRow = 2
    
    ' Initialize ticker calculation variables
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim totalVol As LongLong
    totalVol = 0
    ' Define quarter start/end dates
    
    ' Loop through list
    For i = 2 To 22771
        ' If the next row is a new ticker....
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticker = Cells(i, 1).Value
            totalVol = totalVol + Cells(i, 7).Value
            Cells(outputRow, 9).Value = ticker
            Cells(outputRow, 12).Value = totalVol
            outputRow = outputRow + 1
            totalVol = 0
        ' If the next row is the same ticker...
        Else
            totalVol = totalVol + Cells(i, 7)
        End If
        Cells(i, 8).Value = i
    Next i
    
End Sub
