Attribute VB_Name = "Module1"
'==================================================================================
'Author: - ray -                                                Written: 06.26.2019
'Narrative:
'1. Create a script that will loop through one year of stock data for each run and
'   return the total volume each stock had over that year
'2. Display the ticker symbol to coincide with the total stock volume
'3. Add yearly change & percent change from opening price at the beginning of a given year to the
'   closing price at the end of that year
'4. Add conditional formatting that will highlight positive change in green and negative change in red
'5. Return the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume".
'----------------------------------------------------------------------------------
'Versions:
'06.26.2019 - ray - initial version, the "easy" one
'06.27.2019 - ray - moderate version
'06.28.2019 - ray - hard version
'==================================================================================

Sub MyStockAnalyzer():
'declarations   -------------------------------------------------------------
    Dim tickerHdr As String
    Dim tickerIn As String
    Dim tickerBefore As String
    Dim tickerOut(0 To 5000) As String
    
    Dim volumeHdr As String
    Dim volumeIn As Double
    Dim volumeSofar As Double
    Dim volumeOut(0 To 5000) As Double

    Dim yearIn As String
    Dim yearToAnalyze As String
    Dim endRow As Double
    Dim tick As Double
    
    Dim priceChgHdr As String
    Dim pricePctHdr As String
    Dim priceOpnIn As Double
    Dim priceClsIn As Double
    Dim priceBeg As Double
    Dim priceEnd As Double
    Dim priceChg(0 To 5000) As Double
    Dim pricePct(0 To 5000) As Double
    Dim colormeGreen As Integer
    Dim colormeRed As Integer
    
    Dim bestIncHdr As String
    Dim bestDecHdr As String
    Dim bestVolHdr As String
    Dim bestTickHdr As String
    Dim bestValueHdr As String
    Dim bestTick(0 To 2) As String
    Dim bestValue(0 To 2) As Double
        
'config         -------------------------------------------------------------
    tickerHdr = "< Ticker >"
    volumeHdr = "< Total Stock Volume >"
    priceChgHdr = "< Yearly Change >"
    pricePctHdr = "< Percent Change >"
    bestIncHdr = "Greatest % Increase"
    bestDecHdr = "Greatest % Decrease"
    bestVolHdr = "Greatest Test Volume"
    bestTickHdr = "< Ticker >"
    bestValueHdr = "< Value >"
    tickerBefore = "firstpass"
    yearToAnalyze = "2015"
    colormeGreen = 4
    colormeRed = 3
    endRow = Cells(Rows.Count, "A").End(xlUp).Row
    tick = 0
'main loop      -------------------------------------------------------------
    For i = 2 To endRow
'get input
        tickerIn = Cells(i, 1)
        volumeIn = Cells(i, 7)
        yearIn = Left((Cells(i, 2).Value), 4)
        priceOpnIn = Cells(i, 3)
        priceClsIn = Cells(i, 6)
'if year out of scope
        If (yearIn <> yearToAnalyze) Then '-> Keep this line. It validates year in scope, else do next i
        'ElseIf (Left(tickerIn, 1) <> "A") Then '-> Keep this line for debugging
        ElseIf (tickerBefore = "firstpass") Then
'if first pass
            yearBefore = yearIn
            tickerBefore = tickerIn
            volumeSofar = volumeSofar + volumeIn
            volumeOut(tick) = volumeSofar
            tickerOut(tick) = tickerBefore
            
            priceBeg = priceOpnIn
            priceEnd = priceClsIn
            
            For h = 0 To 2
                bestTick(h) = tickerBefore
                bestValue(h) = volumeSofar
            Next h
        ElseIf (tickerIn = tickerBefore) Then
'if same ticker, aggregate into collection
            volumeSofar = volumeSofar + volumeIn
            volumeOut(tick) = volumeSofar
            tickerOut(tick) = tickerBefore
'calculate volume changes and percentages
            priceEnd = priceClsIn
            priceChg(tick) = (priceEnd - priceBeg)
            If priceBeg > 0 Then
                    pricePct(tick) = ((priceEnd - priceBeg) / priceBeg)
            ElseIf pricePct(tick) = 0 Then
            End If
        ElseIf (tickerIn <> tickerBefore) Then
'calculate volume changes and percentages
            priceChg(tick) = (priceEnd - priceBeg)
            If priceBeg > 0 Then
                    pricePct(tick) = ((priceEnd - priceBeg) / priceBeg)
            ElseIf pricePct(tick) = 0 Then
            End If
'if new ticker, push out what has been aggregated
            volumeOut(tick) = volumeSofar
            tickerOut(tick) = tickerBefore
'compare and collect greatest increase, decrease or Volume
            If pricePct(tick) > bestValue(0) Then
                                                        bestTick(0) = tickerOut(tick)
                                                        bestValue(0) = pricePct(tick)
            ElseIf pricePct(tick) < bestValue(1) Then
                                                        bestTick(1) = tickerOut(tick)
                                                        bestValue(1) = pricePct(tick)
            End If
            If volumeOut(tick) > bestValue(2) Then
                                                        bestTick(2) = tickerOut(tick)
                                                        bestValue(2) = volumeOut(tick)
            End If
'reset tokens for new ticker
            tickerBefore = tickerIn
            volumeSofar = volumeIn
            tick = tick + 1
            volumeOut(tick) = volumeSofar
            tickerOut(tick) = tickerBefore
            
            priceBeg = priceOpnIn
            priceEnd = priceClsIn
        End If
    Next i
'output loop    -------------------------------------------------------------
    Range("I1") = tickerHdr
    Range("J1") = priceChgHdr
    Range("K1") = pricePctHdr
    Range("L1") = volumeHdr
    Range("N2") = bestIncHdr
    Range("N3") = bestDecHdr
    Range("N4") = bestVolHdr
    Range("O1") = bestTickHdr
    Range("P1") = bestValueHdr

    For j = 2 To (tick + 2)
        Cells(j, 9) = tickerOut(j - 2)
        Cells(j, 10) = priceChg(j - 2)
        Cells(j, 11) = pricePct(j - 2)
        Cells(j, 12) = volumeOut(j - 2)
      
        Cells(j, 10).NumberFormat = "0.000000000"
        Cells(j, 11) = Format(Cells(j, 11), "percent")
        
        If Cells(j, 10) = 0 Then
        ElseIf Cells(j, 10) > 0 Then
                Cells(j, 10).Interior.ColorIndex = colormeGreen
        Else:   Cells(j, 10).Interior.ColorIndex = colormeRed
        End If
    Next j
    
    Cells(2, 16) = Format(Cells(2, 16), "percent")
    Cells(3, 16) = Format(Cells(3, 16), "percent")
    Cells(4, 16).NumberFormat = "0"
    
    For k = 2 To 4
        Cells(k, 15) = bestTick(k - 2)
        Cells(k, 16) = bestValue(k - 2)
    Next k
End Sub
'end of script  -------------------------------------------------------------

