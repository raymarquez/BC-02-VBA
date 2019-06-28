Attribute VB_Name = "Module1"
'==================================================================================
'Author: - ray -                                                Written: 06.26.2019
'Narrative:
'1. Create a script that will loop through one year of stock data for each run and
'   return the total volume each stock had over that year
'2. Display the ticker symbol to coincide with the total stock volume
'----------------------------------------------------------------------------------
'Versions:
'06.26.2019 - ray - initial version, the "easy" one
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
'config         -------------------------------------------------------------
    tickerHdr = "< Ticker >"
    volumeHdr = "< Total Stock Volume >"
    tickerBefore = "firstpass"
    yearToAnalyze = "2015"
    'endRow = 760192
    'lRow = Cells(.Rows.Count).End(xlUp).Row
    'lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row   '<== uncomment me please
    endRow = Cells(Rows.Count, "A").End(xlUp).Row
    tick = 0
'main loop      -------------------------------------------------------------
    For i = 2 To endRow
'get input
        tickerIn = Cells(i, 1)
        volumeIn = Cells(i, 7)
        yearIn = Left((Cells(i, 2).Value), 4)
'if year out of scope
        If (yearIn <> yearToAnalyze) Then '-> Keep this line. It validates year in scope, else do next i
        'ElseIf (Left(tickerIn, 1) <> "A") Then
        ElseIf (tickerBefore = "firstpass") Then
'if first pass
            yearBefore = yearIn
            tickerBefore = tickerIn
            volumeSofar = volumeSofar + volumeIn
            volumeOut(tick) = volumeSofar
            tickerOut(tick) = tickerBefore
        ElseIf (tickerIn = tickerBefore) Then
'if same ticker, aggregate into collection
            volumeSofar = volumeSofar + volumeIn
            volumeOut(tick) = volumeSofar
            tickerOut(tick) = tickerBefore
        ElseIf (tickerIn <> tickerBefore) Then
'if new ticker, push out
            volumeOut(tick) = volumeSofar
            tickerOut(tick) = tickerBefore
'reset tokens for new ticker
            tickerBefore = tickerIn
            volumeSofar = volumeIn
            tick = tick + 1
            volumeOut(tick) = volumeSofar
            tickerOut(tick) = tickerBefore
        End If
    Next i
'output loop    -------------------------------------------------------------
    Range("I1") = tickerHdr
    Range("J1") = volumeHdr

    For j = 2 To (tick + 2)
        Cells(j, 9) = tickerOut(j - 2)
        Cells(j, 10) = volumeOut(j - 2)
    Next j
End Sub
'end of script  -------------------------------------------------------------

