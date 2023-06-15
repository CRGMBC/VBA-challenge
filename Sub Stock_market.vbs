Sub Stock_market()

'Declare and set worksheet
Dim ws As Worksheet

'Loop through all stocks for one year
For Each ws In Worksheets

'Create the column headings
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

'Define Ticker variable
Dim Ticker As String
Ticker = " "
Dim Ticker_volume As Double
Ticker_volume = 0

'Create variable to hold stock volume
Dim stock_volume As Double
stock_volume = 0

'Set initial and last row for worksheet
Dim Lastrow As Long
Dim i As Long
Dim j As Integer

'Define Lastrow of worksheet
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Set new variables for prices and percent changes
Dim open_price As Double
open_price = ws.Cells(2, 3).Value
Dim close_price As Double

Dim price_change As Double

Dim price_change_percent As Double

j = 0



    'Do loop of current worksheet to Lastrow
    For i = 2 To Lastrow

        'Ticker symbol output
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker = ws.Cells(i, 1).Value

        'Calculate change in Price
        close_price = ws.Cells(i, 6).Value
        price_change = close_price - open_price

            'if Open price = 0
            'If open_price <> 0 Then
            price_change_percent = (price_change / open_price) * 100

        'End If

                ' Output summary results to new columns
                ' Output Ticker to column I
                ' ws.Cells(i, 9).Value = Ticker
                ws.Range("I" & 2 + j).Value = Ticker
                ' Output Yearly Change to column J
                ws.Range("J" & 2 + j).Value = price_change
                
                ' Output Percent Change to column K
                ws.Range("K" & 2 + j).Value = price_change_percent
                
                ' Output Total Stock Volume to column L
                ws.Range("L" & 2 + j).Value = stock_volume
                
                'Formatting
                'formatnumber
                ws.Range("K" & 2 + j).NumberFormat = "0.00%"
                
                'select Case for color
                If price_change >= 0 Then
                ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                
                Else
                ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                
                End If
                
                ' Reset variables for the next ticker
                
                open_price = ws.Cells(i + 1, 3).Value
                
                stock_volume = 0
                j = j + 1
                
            Else
                ' Accumulate stock volume
                stock_volume = stock_volume + ws.Cells(i, 7).Value
                
                ' Capture the opening price of the ticker
               If open_price = 0 Then
                    open_price = ws.Cells(i, 3).Value
           End If
        End If

    Next i

    'use wsfunction for min and max
    'find the index of where the numbrs came from - wsfunction
    'from index, look up stock ticker





Next ws

End Sub

