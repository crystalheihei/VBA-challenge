
# VBA Challenge

Overview of Project:

Create a script that loops through all the stocks and outputs the following:
 * The ticker symbol.
 * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
 * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
 * The total stock volume of the stock.

This script should run on every worksheet at once.

Use conditional formatting that will highlight positive change in green and negative change in red.

Add functionality to the script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
 * This is to see which tickers you shouldn't buy and which tickers you can buy by reviewing the recent years' data.

Code and Analysis:
'Set was as a worksheet object variable.  
Dim ws As Worksheet

'Loop through all of the worksheets in the workbook
For Each ws In Worksheets

'Set intitial variables for calculations
Dim TickerSymbol As String
Dim BeginYear, EndYear, YearlyChange, YearlyPercentageChange As Double
Dim LastRow As Long
Dim TotalStockVolume As Long

'Initially set the variables to be 0 for each row
BeginYear = 0
EndYear = 0
YearlyPercentageChange = 0
TotalStockVolume = 0

' Set titles for Column I, J, K, L
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volumn"

'Set location for variables
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

'Counts the number of rows
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Loop through each row
For i = 2 To LastRow
'Check if we are still on the same ticker, set the ticker name startting point and ending point, calculate the BeginYear, TickerSymbol, TotalStockVolumn, Endyear and yearlyChange
If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
BeginYear = ws.Cells(i, 3).Value

End If

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
TickerSymbol = ws.Cells(i, 1).Value
TotalStockVolumn = TotalStockVolumn + ws.Cells(i, 7).Value
EndYear = ws.Cells(i, 6).Value
YearlyChange = EndYear - BeginYear

End If

'Set condition for a zero value and calculate the YearlyPercentageChange
If BeginYear <> 0 Then
    YearlyPercentageChange = (YearlyChange / BeginYear) * 100

End If

'Print the TickerSymbol in the summary table Column I, print the YearlyChange in the summary table Column J, print the YearlyPercentageChange in the summary table column K, print the TotalStockVolumn in the summary talbe Colmn L
ws.Range("I" & Summary_Table_Row).Value = TickerSymbol
ws.Range("J" & Summary_Table_Row).Value = YearlyChange
ws.Range("K" & Summary_Table_Row).Value = YearlyPercentageChange & "%"
ws.Range("L" & Summary_Table_Row).Value = TotalStockVolumn

'Add 1 to the summary table row count
Summary_Table_Row = Summary_Table_Row + 1

'Reset values
YearlyChange = 0
YearlyPercentageChange = 0
TotalStockVolumn = 0

'Else if in next ticker name, enter new ticker stock volumn
TotalStockVolumn = TotalStockVolumn + ws.Cells(i, 7).Value
YearlyChange = EndYear - BeginYear

End If

'Change YearlyChange column: Red for negative and green for positive
If ws.Cells(i, 10).Value > 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 4

ElseIf ws.Cells(i, 10).Value < 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 3
ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3

End If

Next i

'Add functionality to the script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume"

Dim PercentageRange As Range
Dim VolumnRange As Range

Set PercentageRange = ws.Range("K2:K" & LastRow)
Set VolumnRange = ws.Range("L2:L" & LastRow)

ws.Cells(2, 17).Value = Application.WorksheetFunction.Max(PercentageRange)
ws.Cells(3, 17).Value = Application.WorksheetFunction.Min(PercentageRange)
ws.Cells(4, 17).Value = Application.WorksheetFunction.Max(VolumnRange)
ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(3, 17).NumberFormat = "0.00%"

For j = 2 To LastRow

If ws.Cells(2, 17).Value = ws.Cells(j, 11).Value Then

ws.Cells(2, 16).Value = ws.Cells(j, 9).Value

End If

If ws.Cells(3, 17).Value = ws.Cells(j, 11).Value Then

ws.Cells(3, 16).Value = ws.Cells(j, 9).Value

End If

If ws.Cells(4, 17).Value = ws.Cells(j, 12).Value Then

ws.Cells(4, 16).Value = ws.Cells(j, 9).Value

End If


Next j

ws.Activate
Debug.Print ws.Name

Next

End Sub