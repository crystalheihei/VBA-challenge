Sub Stocks()

Dim ws As Worksheet
For Each ws In Worksheets

Dim TickerSymbol As String
Dim BeginYear, EndYear, YearlyChange, YearlyPercentageChange As Double
Dim LastRow As Long
Dim TotalStockVolume As Long

BeginYear = 0
EndYear = 0
YearlyPercentageChange = 0
TotalStockVolume = 0

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volumn"

Dim Summary_Table_Row As Integer

Summary_Table_Row = 2

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LastRow


If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
BeginYear = ws.Cells(i, 3).Value

End If


If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
TickerSymbol = ws.Cells(i, 1).Value
TotalStockVolumn = TotalStockVolumn + ws.Cells(i, 7).Value
EndYear = ws.Cells(i, 6).Value
YearlyChange = EndYear - BeginYear

End If

If BeginYear <> 0 Then
    YearlyPercentageChange = (YearlyChange / BeginYear) * 100
    
End If

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

ws.Range("I" & Summary_Table_Row).Value = TickerSymbol
ws.Range("J" & Summary_Table_Row).Value = YearlyChange
ws.Range("K" & Summary_Table_Row).Value = YearlyPercentageChange & "%"
ws.Range("L" & Summary_Table_Row).Value = TotalStockVolumn
Summary_Table_Row = Summary_Table_Row + 1

YearlyChange = 0
YearlyPercentageChange = 0
TotalStockVolumn = 0

Else
TotalStockVolumn = TotalStockVolumn + ws.Cells(i, 7).Value
YearlyChange = EndYear - BeginYear

End If

If ws.Cells(i, 10).Value > 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 4

ElseIf ws.Cells(i, 10).Value < 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 3

End If

Next i

Dim PercentageRange As Range
Dim VolumnRange As Range


Set PercentageRange = ws.Range("K2:K" & LastRow)
Set VolumnRange = ws.Range("L2:L" & LastRow)


ws.Cells(2, 17).Value = Application.WorksheetFunction.Max(PercentageRange)
ws.Cells(3, 17).Value = Application.WorksheetFunction.Min(PercentageRange)
ws.Cells(4, 17).Value = Application.WorksheetFunction.Max(VolumnRange)
ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(3, 17).NumberFormat = "0.00%"

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volumn"

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

