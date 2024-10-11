Attribute VB_Name = "Module1"
Sub StockAnalysisFillSheet()

    Dim ws As Worksheet
    Dim lastrow As Long
    Dim i As Long
    Dim ticker As String
    Dim openprice As Double
    Dim closeprice As Double
    Dim quarterlychange As Double
    Dim percentchange As Double
    Dim totalvolume As Double
    Dim summaryRow As Long
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String

    ' Initialize tracking variables for greatest values
    greatestIncrease = 0
    greatestDecrease = 100 ' Start with a high value for decrease
    greatestVolume = 0

    ' Loop through each sheet (A, B, C, D, E, F)
    For Each ws In ThisWorkbook.Sheets(Array("A", "B", "C", "D", "E", "F"))

        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ' Add headers to columns H, I, J, and K
        ws.Cells(1, 8).Value = "Ticker"
        ws.Cells(1, 9).Value = "Quarterly Change"
        ws.Cells(1, 10).Value = "Percent Change"
        ws.Cells(1, 11).Value = "Total Stock Volume"

        ' Initialize variables for tracking quarterly data
        openprice = ws.Cells(2, 3).Value ' Opening price of the first stock in the sheet
        totalvolume = 0

        ' Loop through each stock entry in the worksheet
        For i = 2 To lastrow

            ' Get the current ticker and accumulate volume
            ticker = ws.Cells(i, 1).Value
            totalvolume = totalvolume + ws.Cells(i, 7).Value

            ' If this is the last row for this stock or we're at the last row of the sheet
            If ws.Cells(i + 1, 1).Value <> ticker Or i = lastrow Then

                ' Get the closing price for this stock
                closeprice = ws.Cells(i, 6).Value

                ' Calculate quarterly change and percentage change
                quarterlychange = closeprice - openprice
                If openprice <> 0 Then
                    percentchange = (quarterlychange / openprice) * 100
                Else
                    percentchange = 0
                End If

                ' Output the same values for every row that belongs to this stock
                For j = 2 To i
                    ws.Cells(j, 8).Value = ticker ' Ticker in column H
                    ws.Cells(j, 9).Value = quarterlychange ' Quarterly change in column I
                    ws.Cells(j, 10).Value = percentchange ' Percentage change in column J
                    ws.Cells(j, 11).Value = totalvolume ' Total volume in column K

                    ' Apply conditional formatting for Percent Change (column J)
                    If percentchange > 0 Then
                        ws.Cells(j, 10).Interior.Color = RGB(0, 255, 0) ' Green for positive
                    ElseIf percentchange < 0 Then
                        ws.Cells(j, 10).Interior.Color = RGB(255, 0, 0) ' Red for negative
                    End If
                Next j

                ' Check for greatest % increase
                If percentchange > greatestIncrease Then
                    greatestIncrease = percentchange
                    greatestIncreaseTicker = ticker
                End If

                ' Check for greatest % decrease
                If percentchange < greatestDecrease Then
                    greatestDecrease = percentchange
                    greatestDecreaseTicker = ticker
                End If

                ' Check for greatest total volume
                If totalvolume > greatestVolume Then
                    greatestVolume = totalvolume
                    greatestVolumeTicker = ticker
                End If

                ' Reset variables for next stock
                If i + 1 <= lastrow Then
                    openprice = ws.Cells(i + 1, 3).Value ' Set the opening price for the next stock
                End If
                totalvolume = 0 ' Reset total volume for the next stock
            End If

        Next i

        ' Output the greatest values in columns L, M, N
        ws.Cells(2, 13).Value = "Greatest % Increase"
        ws.Cells(3, 13).Value = "Greatest % Decrease"
        ws.Cells(4, 13).Value = "Greatest Total Volume"
        
        ws.Cells(2, 14).Value = greatestIncreaseTicker
        ws.Cells(3, 14).Value = greatestDecreaseTicker
        ws.Cells(4, 14).Value = greatestVolumeTicker
        
        ws.Cells(2, 15).Value = Format(greatestIncrease, "0.00") & "%"
        ws.Cells(3, 15).Value = Format(greatestDecrease, "0.00") & "%"
        ws.Cells(4, 15).Value = Format(greatestVolume, "0.00E+00")
    
    Next ws

End Sub
