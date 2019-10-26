Sub modSol()
Dim current As Worksheet
Dim maxVolume As Double
Dim greatestInc, greatestDec As Double
Dim total_row As Double
Dim cellNumber As Long
Dim perc1, perc2 As String


For Each current In Worksheets
    total_row = current.Cells(Rows.Count, 12).End(xlUp).Row
    current.Range("P1").Value = "Ticker"
    current.Range("Q1").Value = "Value"
    current.Range("O2").Value = "Greatest % increase"
    current.Range("O3").Value = "Greatest % Decrease"
    current.Range("O4").Value = "Greatest Total Volume"
    maxVolume = WorksheetFunction.Max(current.Range("L2:L" & total_row))
    greatestInc = WorksheetFunction.Max(current.Range("K2:K" & total_row))
    greatestDec = WorksheetFunction.Min(current.Range("K2:K" & total_row))
    
    For i = 2 To total_row
        If (current.Cells(i, 11).Value = greatestInc) Then
            current.Range("P2").Value = current.Cells(i, 9).Value
            perc1 = FormatPercent(greatestInc, 2)
            current.Range("Q2").Value = perc1
        End If
        If (current.Cells(i, 11).Value = greatestDec) Then
            current.Range("P3").Value = current.Cells(i, 9).Value
            perc1 = FormatPercent(greatestDec, 2)
            current.Range("Q3").Value = perc1
            
        End If
        If (current.Cells(i, 12).Value = maxVolume) Then
            current.Range("P4").Value = current.Cells(i, 9).Value
            current.Range("Q4").Value = maxVolume
        End If
    Next i
  Next current
End Sub


