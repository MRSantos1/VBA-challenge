
Sub StockData()
    ' looks through the sheets

Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
        ' Determine the Last Row
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

        ' Add Heading for summary
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"
        'Create Variable to hold Value
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

        ' Use for opening price
        Open_Price = Cells(2, Column + 2).Value
         ' Loop through all ticker symbol

        For i = 2 To LastRow
         ' Check if we are still within the same ticker symbol, if it is not...
            If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
                ' Set Ticker name
                Ticker_Name = Cells(i, Column).Value
                Cells(Row, Column + 8).Value = Ticker_Name
                ' Set Close Price
                Close_Price = Cells(i, Column + 5).Value
                ' Add Yearly Change
                Yearly_Change = Close_Price - Open_Price
                Cells(Row, Column + 9).Value = Yearly_Change
                ' Add Percent Change
                If (Open_Price = 0 And Close_Price = 0) Then
                    Percent_Change = 0
                ElseIf (Open_Price = 0 And Close_Price <> 0) Then
                    Percent_Change = 1
                Else
                    Percent_Change = Yearly_Change / Open_Price
                    Cells(Row, Column + 10).Value = Percent_Change
                    Cells(Row, Column + 10).NumberFormat = "0.00%"
                End If
                ' Add Total Volumn
                Volume = Volume + Cells(i, Column + 6).Value
                Cells(Row, Column + 11).Value = Volume
                ' Add one to the summary table row
                Row = Row + 1
                ' reset the Open Price
                Open_Price = Cells(i + 1, Column + 2)
                ' reset the Volumn Total
                Volume = 0
            'if cells are the same ticker
            Else
                Volume = Volume + Cells(i, Column + 6).Value
            End If

        Next i


       Next WS
End Sub
