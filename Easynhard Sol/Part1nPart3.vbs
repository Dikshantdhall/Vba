Sub easyNhardsol()
Dim ticker_Name As String
    Dim total_vol As Double
    Dim part_tickervolume As Double
    Dim opn_Yearlyprice, close_Yearlyprice, total_Yearlychange As Single
    Dim percent_Change As Double
    Dim ind As Long
    Dim pc As String
    
    
    For Each current In Worksheets
        current.Range("I1").Value = "Ticker"
        current.Range("J1").Value = "Yearly Change"
        current.Range("k1").Value = "Percent Change"
        current.Range("L1").Value = "Total Stock Volume"
        total_row = current.Cells(Rows.Count, 1).End(xlUp).Row
        tick = 2
        total_vol = 0
        ind = 2
       
        
        
        For i = 2 To total_row
           
            If (current.Cells(i, 1).Value <> current.Cells(i + 1, 1).Value) Then
               
                ticker_Name = current.Cells(i, 1).Value
                current.Range("I" & tick).Value = ticker_Name
                
                opn_Yearlyprice = current.Range("C" & ind).Value
                close_Yearlyprice = current.Cells(i, 6).Value
                
                total_Yearlychange = close_Yearlyprice - opn_Yearlyprice
                current.Range("J" & tick).Value = total_Yearlychange
                If (current.Range("J" & tick).Value) > 0 Then
                    current.Range("J" & tick).Interior.ColorIndex = 4
                Else
                    current.Range("J" & tick).Interior.ColorIndex = 3
                End If
                If (opn_Yearlyprice = 0) Then
                    percent_Change = 0
                Else
                    percent_Change = total_Yearlychange / opn_Yearlyprice
                End If
                pc = FormatPercent(percent_Change, 2)
                current.Range("K" & tick).Value = pc
                part_tickervolume = total_vol + current.Cells(i, 7).Value
                current.Range("L" & tick).Value = part_tickervolume
                tick = tick + 1
                ind = i + 1
                total_vol = 0
                
            Else
                total_vol = total_vol + current.Cells(i, 7).Value
            End If
        Next i
     Next current
End Sub



