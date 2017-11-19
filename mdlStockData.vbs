
Sub Button1_Click()
    ' For each worksheet in workbook, run the statistics generator
    For Each ws In Worksheets
        Call GenerateStatistics(ws)
    Next
End Sub

Function GenerateStatistics(ByVal ws As Worksheet)
    ' Initialize summary table variables
    Dim dblOpeningPrice, dblClosingPrice, dblGreatestPercentIncrease, dblGreatestPercentDecrease As Double
    Dim arrTickers(), strTickerGreatestIncrease, strTickerGreatestDecrease, strTickerGreatestVolume As String
    Dim arrYearlyChange(), arrPercentChange() As Double
    Dim arrYearlyVolume(), dblGreatestVolume As Double
    
    ' Initialize utility variables
    Dim lRow, summaryRowIncrement As Long
    Dim dblTickerVolume As Double
     
    ' Reset row increment variable, total ticker volume variable, greatest percent increase / decrease, greatest volume variables
    summaryRowIncrement = 1
    dblTickerVolume = 0
    dblGreatestPercentIncrease = 0
    dblGreatestPercentDecrease = 0
    dblGreatestVolume = 0
    
    ' Get the number of the last row in the sheet
    lRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Loop through the dataset
    For i = 2 To lRow
        ' Increment the ticker-specific total volume
        dblTickerVolume = dblTickerVolume + ws.Cells(i, 7).Value
        
        ' If we are on the first row of a new ticker set, determine the opening price for this ticker
        If i = 2 Or ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            dblOpeningPrice = ws.Cells(i, 3).Value
        End If
        
       ' Determine if the stock ticker is about to change to a new set
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            ' Determine the closing price for this set
            dblClosingPrice = ws.Cells(i, 6).Value
            
            ' Expand our array variables to the new upper bounds for the summary table
            ReDim Preserve arrTickers(summaryRowIncrement), arrYearlyChange(summaryRowIncrement), arrYearlyVolume(summaryRowIncrement), arrPercentChange(summaryRowIncrement)
            
            ' Save the ticker name to array variable for summary table
            arrTickers(summaryRowIncrement) = ws.Cells(i, 1).Value
            
            ' Determine the yearly change for this ticker
            arrYearlyChange(summaryRowIncrement) = dblClosingPrice - dblOpeningPrice
            
            ' Determine the percentage change for this ticker, if greater than zero
            If dblOpeningPrice > 0 Then
                arrPercentChange(summaryRowIncrement) = 100 * (arrYearlyChange(summaryRowIncrement) / dblOpeningPrice)
            End If
                        
            ' Save the total yearly volume for this ticker
            arrYearlyVolume(summaryRowIncrement) = dblTickerVolume
            
            ' Set greatest percent increase/decrease variable to the current percent value, if it's greater/less than the previous one
            If arrPercentChange(summaryRowIncrement) > dblGreatestPercentIncrease Then
                dblGreatestPercentIncrease = arrPercentChange(summaryRowIncrement)
                strTickerGreatestIncrease = arrTickers(summaryRowIncrement)
            End If
            If arrPercentChange(summaryRowIncrement) < dblGreatestPercentDecrease Then
                dblGreatestPercentDecrease = arrPercentChange(summaryRowIncrement)
                strTickerGreatestDecrease = arrTickers(summaryRowIncrement)
            End If
            ' Set greatest total volume variable
            If arrYearlyVolume(summaryRowIncrement) > dblGreatestVolume Then
                dblGreatestVolume = arrYearlyVolume(summaryRowIncrement)
                strTickerGreatestVolume = arrTickers(summaryRowIncrement)
            End If
            
            ' Increment array counter for our summary table
            summaryRowIncrement = summaryRowIncrement + 1
            
            ' Reset the ticker-specific total volume variable
            dblTickerVolume = 0
        End If
        
    Next i
       
    ' Reset all formatting in the summary table
    ws.Range("I2:L" & UBound(arrTickers)).Value = ""
    ws.Range("I2:L" & UBound(arrTickers)).Interior.ColorIndex = 0
    
    ' *** Main Summary Table ***
    ' Headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ' Table values - loop through all arrTicker array values (unique ticker list with statistics)
    For i = 1 To UBound(arrTickers)
        ' Values
        ws.Cells(i + 1, 9).Value = arrTickers(i)
        ws.Cells(i + 1, 10).Value = arrYearlyChange(i)
        ws.Cells(i + 1, 11).Value = Round(arrPercentChange(i), 2) & "%"
        ws.Cells(i + 1, 12).Value = arrYearlyVolume(i)
        
        ' Set the background color of the Yearly Change value to red if negative, green if positive, and white if no change
        If arrYearlyChange(i) < 0 Then
            ws.Cells(i + 1, 10).Interior.ColorIndex = 3
        ElseIf arrYearlyChange(i) > 0 Then
            ws.Cells(i + 1, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(i + 1, 10).Interior.ColorIndex = 0
        End If
    Next i

    ' ****** Greatest % increase, decrease, volume table: HEADERS ********
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Range("P1:Q1").Interior.ColorIndex = 15  ' light gray background
    ws.Range("O2:O4").Font.FontStyle = "Italic"
    
    ' ****** Greatest % increase, decrease, volume table: VALUES ********
    ws.Cells(2, 16).Value = strTickerGreatestIncrease
    ws.Cells(3, 16).Value = strTickerGreatestDecrease
    ws.Cells(4, 16).Value = strTickerGreatestVolume
    ws.Cells(2, 17).Value = Round(dblGreatestPercentIncrease, 2) & "%"
    ws.Cells(3, 17).Value = Round(dblGreatestPercentDecrease, 2) & "%"
    ws.Cells(4, 17).Value = dblGreatestVolume
    ws.Cells(4, 17).NumberFormat = "#,###"
     
End Function

