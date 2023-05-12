VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock_Market()
Dim ws As Worksheet
For Each ws In Worksheets

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"


Dim Ticker_volume As LongLong
Ticker_volume = 0
Dim Tickerrows As Integer
Tickerrows = 2
Dim Lastrow As Long
Dim i As Long
Dim cpcount As Long
cpcount = 2

Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Dim open_price As Double
open_price = 0
Dim close_price As Double
close_price = 0
Dim price_change As Double
price_change = 0
Dim price_change_percent As Double
price_change_percent = 0

Dim PrChg, GrIcr, GrDcr, GrVol As Double

For i = 2 To Lastrow
    Ticker_volume = Ticker_volume + ws.Cells(i, 7)
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    ws.Cells(Tickerrows, 9) = ws.Cells(i, 1).Value
    
    open_price = ws.Cells(cpcount, 3).Value
    close_price = ws.Cells(i, 6).Value
    price_change = close_price - open_price
    price_change_percent = (price_change / open_price)
    ws.Cells(Tickerrows, 10) = price_change
    ws.Cells(Tickerrows, 11).Value = Format(price_change_percent, "0.00%")
    ws.Cells(Tickerrows, 12).Value = Ticker_volume
    If price_change < 0 Then
        ws.Cells(Tickerrows, 10).Interior.ColorIndex = 3
    Else
        ws.Cells(Tickerrows, 10).Interior.ColorIndex = 4
    End If
    cpcount = 1 + i
    Tickerrows = Tickerrows + 1
    Ticker_volume = 0

End If

Next i
Lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
GrVol = 0
GrIcr = -100
GrDcr = 100
    For i = 2 To Lastrow
        If ws.Cells(i, 12).Value > GrVol Then
            GrVol = ws.Cells(i, 12).Value
            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(4, 17).Value = GrVol
        End If
        If ws.Cells(i, 11).Value > GrIcr Then
            GrIcr = ws.Cells(i, 11)
            ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(2, 17).Value = Format(GrIcr, "0.00%")
        End If
        If ws.Cells(i, 11).Value < GrDcr Then
            GrDcr = ws.Cells(i, 11)
            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(3, 17).Value = Format(GrDcr, "0.00%")
        End If
    Next i
Next ws

End Sub


