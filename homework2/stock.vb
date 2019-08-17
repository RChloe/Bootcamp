Sub stock():

For Each ws In Worksheets

Dim lastrow, startrow, startdate, volume, value_row, lastrow2, value_row2, value_row3 As Long
Dim startprice, endprice, year_change, percent_change, volumestart, volumeend, most_increase, least_increase As Double
Dim ticker As String
Dim headers, headers2, headers3

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
startrow = 1
startdate = ws.Range("B2").Value

' Set Headers
headers = Array("Ticker", "Yearly Change", "Percent Change", "Total Volume")
ws.Range("I1:L1") = headers
headers2 = Array("Ticker", "Value")
ws.Range("P1:Q1") = headers2
headers3 = Array("Greatest % increase", "Greatest % decrease", "Greatest total volume")
ws.Range("O2:O4") = Application.Transpose(headers3)

For i = 2 To lastrow

    If ws.Cells(i, 2).Value = startdate Then
        startrow = startrow + 1
        ticker = ws.Cells(i, 1).Value
        ws.Cells(startrow, 9).Value = ticker
        startprice = ws.Cells(i, 3).Value
        volumestart = ws.Cells(i, 7).Row
    ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        endprice = ws.Cells(i, 6).Value
        year_change = endprice - startprice
        percent_change = year_change / startprice
        ws.Cells(startrow, 10).Value = year_change
        ws.Cells(startrow, 11).Value = FormatPercent(percent_change, 2)
        volumeend = ws.Cells(i, 7).Row
        ws.Cells(startrow, 12).Value = "=SUM(" & ws.Range("G" & volumestart, "G" & volumeend).Address(False, False) & ")"
        volume = 0
        If year_change < 0 Then
            ws.Cells(startrow, 10).Interior.ColorIndex = 3
        Else
            ws.Cells(startrow, 10).Interior.ColorIndex = 4
        End If
        
    Else
        
    End If
    
Next i

lastrow2 = ws.Cells(Rows.Count, 11).End(xlUp).Row
most_increase = ws.Cells(2, 11).Value
least_increase = ws.Cells(2, 11).Value

ws.Range("Q4").Value = "=MAX(" & ws.Range("L2", "L" & lastrow2).Address(False, False) & ")"

For i = 2 To lastrow2
    
    If ws.Cells(i, 11).Value > most_increase Then
        most_increase = ws.Cells(i, 11).Value
        ws.Range("P2").Value = ws.Cells(i, 9).Value
        ws.Range("Q2").Value = FormatPercent(most_increase, 2)
    End If
    If ws.Cells(i, 11).Value < least_increase Then
        least_increase = ws.Cells(i, 11).Value
        ws.Range("P3").Value = ws.Cells(i, 9).Value
        ws.Range("Q3").Value = FormatPercent(least_increase, 2)
    End If
    If Str(ws.Cells(i, 12).Value) = Str(ws.Range("Q4").Value) Then
        ws.Range("P4").Value = ws.Cells(i, 9).Value
    End If
    
Next i

Next ws

End Sub
