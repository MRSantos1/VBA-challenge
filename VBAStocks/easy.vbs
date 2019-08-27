Sub HomeworkStockCleaner()



Dim WS As Worksheet

For Each WS In ActiveWorkbook.Worksheets

WS.Activate


LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row



Cells(1, "I").Value = "Ticker"

Cells(1, "J").Value = "Yearly Change"

Cells(1, "K").Value = "Percent Change"

Cells(1, "L").Value = "Total Stock Volume"


Dim Open_Price As Double

Dim Close_Price As Double

Dim Yearly_Change As Double

Dim Ticker_Name As String

Dim Percent_Change As Double

Dim Volume As Double

Volume = 0

Dim Row As Double

Row = 2

Dim Column As Integer

Column = 1

Dim i As Long





Open_Price = Cells(2, Column + 2).Value





For i = 2 To LastRow



If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then



Ticker_Name = Cells(i, Column).Value

Cells(Row, Column + 8).Value = Ticker_Name



Close_Price = Cells(i, Column + 5).Value



Yearly_Change = Close_Price - Open_Price

Cells(Row, Column + 9).Value = Yearly_Change



If (Open_Price = 0 And Close_Price = 0) Then

Percent_Change = 0

ElseIf (Open_Price = 0 And Close_Price <> 0) Then

Percent_Change = 1

Else

Percent_Change = Yearly_Change / Open_Price

Cells(Row, Column + 10).Value = Percent_Change

Cells(Row, Column + 10).NumberFormat = "0.00%"

End If



Volume = Volume + Cells(i, Column + 6).Value

Cells(Row, Column + 11).Value = Volume



Row = Row + 1



Open_Price = Cells(i + 1, Column + 2)



Volume = 0



Else

Volume = Volume + Cells(i, Column + 6).Value

End If

Next i