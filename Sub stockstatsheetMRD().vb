Sub stockstatsheet()
           
            Dim yearopnfound As Boolean
            yearopnfound = False
            tot_Stock_vol = 0
            Dim year_opn As Double, year_cls As Double, perc_cha As Double
            Dim myRange As Range, myrange2 As Range, maxperc As Double, minperc As Double, maxtot As Double, progress As Double
            
    For Each ws In Worksheets
        Sum_Tab_Row = 2
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
       
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(4, 15).Value = "Greatest Total Volume"
             
       For i = 2 To lastrow
                        
            If yearopnfound = False Then
                year_opn = ws.Cells(i, 3)
                yearopnfound = True
                
            End If
                        
                            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    ticker = ws.Cells(i, 1).Value
                    tot_Stock_vol = tot_Stock_vol + ws.Cells(i, 7).Value
                    ws.Range("I" & Sum_Tab_Row).Value = ticker
                    ws.Range("L" & Sum_Tab_Row).Value = tot_Stock_vol
                    tot_Stock_vol = 0
                    
                    year_cls = ws.Cells(i - 1, 6)
                    
                    yearly_cha = year_cls - year_opn
                    ws.Range("J" & Sum_Tab_Row).Value = yearly_cha
                   
                   If year_opn = 0 Then
                        ws.Range("K" & Sum_Tab_Row).Value = "NA"
                    Else
                    
                    perc_cha = (year_cls - year_opn) / (year_opn)
                        ws.Range("K" & Sum_Tab_Row).Value = perc_cha
                        ws.Range("K" & Sum_Tab_Row).NumberFormat = "0.00%"
                    End If
                    
                    
                    Sum_Tab_Row = Sum_Tab_Row + 1
                Else
                    tot_Stock_vol = tot_Stock_vol + ws.Cells(i, 7).Value
            End If
                
            If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
                yearopnfound = False
            End If
            
           
            progress = i / lastrow
             Application.StatusBar = "Progress: " & i & " of " & lastrow & ": " & Format(progress, "0%")
       Next i
            
    
     lastrow2 = ws.Cells(Rows.Count, 10).End(xlUp).Row
    For i = 2 To lastrow2
           
            Set myRange = Worksheets(ws.Name).Range("k2:k" & lastrow2)
            Set myrange2 = Worksheets(ws.Name).Range("L2:L" & lastrow2)
               
                If (ws.Cells(i, 10).Value < 0) Then
                        ws.Cells(i, 10).Interior.ColorIndex = 3
                ElseIf (ws.Cells(i, 10).Value > 0) Then
                        ws.Cells(i, 10).Interior.ColorIndex = 4
                ElseIf (ws.Cells(i, 10).Value = 0) Then
                       ws.Cells(i, 10).Interior.ColorIndex = 6
                End If
                    
               If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(myRange) Then
                   maxperc = ws.Cells(i, 11).Value
                   maxptick = ws.Cells(i, 9).Value
                   ws.Cells(2, 17).Value = maxperc
                        ws.Cells(2, 17).NumberFormat = "0.00%"
                   ws.Cells(2, 16).Value = maxptick
                End If
                
                If ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(myRange) Then
                    minperc = ws.Cells(i, 11).Value
                    minptick = ws.Cells(i, 9).Value
                    ws.Cells(3, 17).Value = minperc
                        ws.Cells(3, 17).NumberFormat = "0.00%"
                    ws.Cells(3, 16).Value = minptick
                End If
                
                If ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(myrange2) Then
                    maxtot = ws.Cells(i, 12).Value
                    maxttick = ws.Cells(i, 9).Value
                    ws.Cells(4, 17).Value = maxtot
                    ws.Cells(4, 16).Value = maxttick
                End If
         Next i
        Next
        Application.StatusBar = False
End Sub
