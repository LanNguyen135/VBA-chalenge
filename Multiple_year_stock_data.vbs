Sub YearlyInfo()
    'Set up the 1st table
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    'Add ticker symbol to I column
    Dim ticker As String
    Dim i As Long
    i = 2
    Dim j As Long
    j = 2
    While Cells(i, "A").Value <> ""
        ticker = Cells(i - 1, "A").Value
        If Cells(i, "A").Value <> ticker Then
            Cells(j, "I").Value = Cells(i, "A").Value
            j = j + 1
        End If
        i = i + 1
    Wend
        
        
    'Calculate yearly change & percentage change
    Dim d As Long
    d = 3
    Dim openDate As Long
    Dim endDate As Long
    openDate = Range("B2").Value
    endDate = 0
    While Cells(d, "A").Value <> ""
        If Cells(d, "B").Value > openDate And Cells(d, "B").Value > endDate Then
            endDate = Cells(d, "B").Value
        End If
        d = d + 1
    Wend
    
    Dim a As Long
    a = 2
    Dim openValue As Single
    Dim endValue As Single
    Dim yearlyChange As Single
    i = 2
    
    While Cells(i, "A").Value <> ""
        If Cells(i, "A").Value = Cells(a, "I").Value And Cells(i, "B").Value = openDate Then
            openValue = Cells(i, "C").Value
        ElseIf Cells(i, "A").Value = Cells(a, "I").Value And Cells(i, "B").Value = endDate Then
            endValue = Cells(i, "F").Value
            yearlyChange = endValue - openValue
            Cells(a, "J").Value = Format(yearlyChange, "0.00")
            
            'Format yearly change
            If yearlyChange < 0 Then
                Cells(a, "J").Interior.ColorIndex = 3
            Else
                Cells(a, "J").Interior.ColorIndex = 4
            End If
            
            Cells(a, "K").Value = endValue / openValue - 1
            Cells(a, "K").NumberFormat = "0.00%"
            a = a + 1
        End If
        i = i + 1
    Wend
    
    'Calculate total stock volume
    Dim stockVolume As LongLong
    stockVolume = 0
    Dim b As Long
    b = 2
    i = 2
    
    While Cells(i, "A").Value <> ""
        If Cells(i + 1, "A").Value <> Cells(i, "A").Value Then
            stockVolume = stockVolume + Cells(i, "G").Value
            Cells(b, "L").Value = stockVolume
            b = b + 1
            stockVolume = 0
        Else
            stockVolume = stockVolume + Cells(i, "G").Value
        End If
        i = i + 1
    Wend
    
    'Set up the 2nd table
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    Dim greatestInc As Single
    Dim greatestDec As Single
    Dim volume As LongLong
    Dim c As Long
    c = 2
    
    While Cells(c, "K") <> ""
        'Find the greastest % increase
        If Cells(c, "K").Value > greatestInc Then
            greatestInc = Cells(c, "K").Value
            Range("P2").Value = Cells(c, "I").Value
            Range("Q2").Value = Format(Cells(c, "K").Value, "0.00%")
        End If
        
        'Find the greatest % decrease
        If Cells(c, "K").Value < greatestDec Then
            greatestDec = Cells(c, "K").Value
            Range("P3").Value = Cells(c, "I").Value
            Range("Q3").Value = Format(Cells(c, "K").Value, "0.00%")
        End If
        
        'Find the greatest total volume
        If Cells(c, "L").Value > volume Then
            volume = Cells(c, "L").Value
            Range("P4").Value = Cells(c, "I").Value
            Range("Q4").Value = Cells(c, "L").Value
        End If
        
        c = c + 1
    Wend
    
    
    
End Sub


