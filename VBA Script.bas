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
    'set title to column I1, J1, K1, L1, P1, Q1, O2, O3, O4
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"

    'define variables
    Dim Ticker_volume As LongLong
    Ticker_volume = 0
    Dim Tickerrows As Integer
    Tickerrows = 2
    Dim Lastrow As Long
    Dim i As Long
    Dim cpcount As Long
    cpcount = 2

    'define variables
    Dim open_price, close_price, price_change, price_change_percent As Double
    open_price = 0
    close_price = 0
    price_change = 0
    price_change_percent = 0

    'define variables
    Dim GrIcr, GrDcr, GrVol As Double
    GrVol = 0
    GrIcr = -100
    GrDcr = 100
    
    'look for the last row in worksheet
    Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    'loop from row 2 to the last row
    For i = 2 To Lastrow
        'adding ticket volume for each row
        Ticker_volume = Ticker_volume + ws.Cells(i, 7)
        'finding different Ticker
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'update Tickername to the Ticker column
            ws.Cells(Tickerrows, 9) = ws.Cells(i, 1).Value
            'fiding the open price of the Ticker
            open_price = ws.Cells(cpcount, 3).Value
            'finding the closing price of the Ticker
            close_price = ws.Cells(i, 6).Value
            'finding the different in price
            price_change = close_price - open_price
            'fiding the difference in percentage
            price_change_percent = (price_change / open_price)
            
            'update the change in price
            ws.Cells(Tickerrows, 10) = price_change
            'update the change in price in percentage
            ws.Cells(Tickerrows, 11).Value = Format(price_change_percent, "0.00%")
            'update Ticker Volume
            ws.Cells(Tickerrows, 12).Value = Ticker_volume
            
            'Conditional Formatting
            If price_change < 0 Then
                'red when less than 0
                ws.Cells(Tickerrows, 10).Interior.ColorIndex = 3
            Else
                'green otherwise
                ws.Cells(Tickerrows, 10).Interior.ColorIndex = 4
            End If
            
            'update dummy variables
            cpcount = 1 + i
            Tickerrows = Tickerrows + 1
            'reset ticker volume to 0
            Ticker_volume = 0
        End If

Next i
    'set new last row of column I
    Lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    For i = 2 To Lastrow
            'finding greatest % increase
        If ws.Cells(i, 11).Value > GrIcr Then
            GrIcr = ws.Cells(i, 11)
            ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(2, 17).Value = Format(GrIcr, "0.00%")
        End If
        If ws.Cells(i, 11).Value < GrDcr Then
            'finding greatest % decrease
            GrDcr = ws.Cells(i, 11)
            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(3, 17).Value = Format(GrDcr, "0.00%")
        End If
           'finding greatest volume
        If ws.Cells(i, 12).Value > GrVol Then
            'update it when finding the value greater than previous assinged GrVol
            GrVol = ws.Cells(i, 12).Value
            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(4, 17).Value = GrVol
        End If
    Next i
Next ws
End Sub


