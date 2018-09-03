Attribute VB_Name = "Module1"
Sub hw2()

Dim lastrow, i, tickernumber, tickernumber2 As Integer
Dim maxtotalvolume As Double
Dim totvol, voladd As Double
Dim firstdate, maxperc_inc, maxperc_dec As Double
Dim ticker_inc, ticker_dec, ticker_maxvol As String



For Each ws In Worksheets

    ws.Range("i1") = "Ticker"
    ws.Range("j1") = "Yearly Change"
    ws.Range("k1") = "Percent Change"
    ws.Range("l1") = "Total Stock Volume"
    ws.Range("p1") = "Ticker"
    ws.Range("q1") = "Value"
    ws.Range("o2") = "Greatest % Increase"
    ws.Range("o3") = "Greatest % Decrease"
    ws.Range("o4") = "Greatest Total Volume"


    lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    tickernumber = 3
    firstdate = ws.Cells(2, 3).Value
    tickernumber2 = 2
    ws.Cells(2, 9).Value = ws.Cells(2, 1).Value
    ws.Cells(2, 12).Value = ws.Cells(2, 7).Value
    maxperc_inc = 0
    maxperc_dec = 0
    totvol = 0

       ' For i = 3 To 15
        For i = 3 To (lastrow + 1)
        
             voladd = CDbl(ws.Cells(i, 7).Value)
             
            If (ws.Cells(i, 1).Value = ws.Cells(i - 1, 1).Value) Then
                
                totvol = totvol + voladd
                ws.Cells(tickernumber - 1, 12).Value = totvol
            
            Else
                
                ws.Cells(tickernumber, 9).Value = ws.Cells(i, 1).Value
                If (i <> (lastrow + 1)) Then
                   totvol = voladd
                ws.Cells(tickernumber, 12).Value = voladd
                End If
                'thisvol = ws.Cells(tickernumber - 1, 12).Value
                
                If (totvol > maxtotalvolume) Then
                
                    maxtotalvolume = CDbl(totvol)
    
                    ticker_maxvol = ws.Cells(tickernumber - 1, 9).Value
                End If
                
                
                tickernumber = tickernumber + 1
                
                        
                        ws.Cells(tickernumber2, 10) = ws.Cells(i - 1, 6).Value - firstdate
                        If firstdate <> 0 Then
                        
                        ws.Cells(tickernumber2, 11) = (ws.Cells(i - 1, 6).Value - firstdate) / firstdate
                        Else
                        ws.Cells(tickernumber2, 11) = 0
                        End If
                        
                        ws.Cells(tickernumber2, 11).NumberFormat = "0.00%"
                        
                        
                        
                        If ws.Cells(tickernumber2, 11) > maxperc_inc Then
                            maxperc_inc = ws.Cells(tickernumber2, 11)
                            ticker_inc = ws.Cells(tickernumber2, 9)
                        ElseIf ws.Cells(tickernumber2, 11) < maxperc_dec Then
                            maxperc_dec = ws.Cells(tickernumber2, 11)
                            ticker_dec = ws.Cells(tickernumber2, 9)
                        End If
                        
                        
                        
                        tickernumber2 = tickernumber2 + 1
                        firstdate = ws.Cells(i, 3).Value
            
            End If
            
            
    
         
            
    
        Next i


ws.Range("p2") = ticker_inc
ws.Range("p3") = ticker_dec
ws.Range("q2") = maxperc_inc
ws.Range("q3") = maxperc_dec
ws.Range("q2:q3").NumberFormat = "0.00%"

ws.Range("q4") = maxtotalvolume
ws.Range("p4") = ticker_maxvol


Next ws



End Sub

