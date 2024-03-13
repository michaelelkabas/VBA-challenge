Attribute VB_Name = "Module3"
Sub StockAnalysis()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryRow As Long
    Dim i As Long ' Declare i as Long data type
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    
    ' Initialize greatest increase, decrease, and volume to ensure consistent behavior throughout the program
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0
    
    ' Loop through each sheet in the workbook
    For Each ws In ThisWorkbook.Sheets
    ' Skip sheets other than "2018", "2019", and "2020"
        If ws.Name <> "2018" And ws.Name <> "2019" And ws.Name <> "2020" Then
        End If
        
        ' Find the last row of data in the current sheet
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Initialize summary table headers for the current sheet
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"
        
        ' Initialize Greatest increase / decrease / total volume table headers for the current sheet
        ws.Cells(2, 15).Value = "Greatest % increase"
        ws.Cells(3, 15).Value = "Greatest % decrease"
        ws.Cells(4, 15).Value = "Greatest total volume"
        
        ' Initialize Greatest Ticker and Value table headers for the current sheet
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
 
        
        ' Initialize summary table row for the current sheet
        summaryRow = 2
        
        ' Loop through each row of data in the current sheet
        For i = 2 To lastRow
            ' Check if the ticker symbol has changed
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Or i = 2 Then
                ' If not the first row, calculate and output results
                If i <> 2 Then
                    ws.Cells(summaryRow, 9).Value = ticker
                    ws.Cells(summaryRow, 10).Value = yearlyChange
                    ws.Cells(summaryRow, 11).Value = percentChange
                    ws.Cells(summaryRow, 12).Value = totalVolume
                    summaryRow = summaryRow + 1
                End If
                ' Set new ticker symbol and reset variables and ensures that the variable like closePrce starts with a clean slate for each stock
                ticker = ws.Cells(i, 1).Value
                openPrice = ws.Cells(i, 3).Value
                closePrice = 0
                yearlyChange = 0
                percentChange = 0
                totalVolume = 0
            End If
            
            ' Accumulate the total volume located in Column 7
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            ' Update close price at the end of the year
            If (ws.Cells(i, 2).Value) <> (ws.Cells(i + 1, 2).Value) Then
                closePrice = ws.Cells(i, 6).Value
                ' Calculate yearly change and percent change
                yearlyChange = closePrice - openPrice
                If openPrice <> 0 Then
                    percentChange = yearlyChange / openPrice
                End If
            End If
                
                ' Update greatest % increase, % decrease, and total volume
                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    greatestIncreaseTicker = ticker
                End If
                If percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    greatestDecreaseTicker = ticker
                End If
                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    greatestVolumeTicker = ticker
                End If
            Next i
        
        ' Output results for the last ticker symbol in the current sheet
        ws.Cells(summaryRow, 9).Value = ticker
        ws.Cells(summaryRow, 10).Value = yearlyChange
        ws.Cells(summaryRow, 11).Value = FormatPercent(percentChange, 2)
        ws.Cells(summaryRow, 12).Value = FormatNumber(totalVolume)
    Next ws
    
    ' Output greatest % increase, % decrease, and total volume in specified cells in each sheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "2018" Then
            ws.Cells(2, 16).Value = greatestIncreaseTicker
            ws.Cells(3, 16).Value = greatestDecreaseTicker
            ws.Cells(4, 16).Value = greatestVolumeTicker
            ws.Cells(2, 17).Value = FormatPercent(greatestIncrease, 2)
            ws.Cells(3, 17).Value = FormatPercent(greatestDecrease, 2)
            ws.Cells(4, 17).Value = FormatNumber(greatestVolume)
        ElseIf ws.Name = "2019" Then
            ws.Cells(2, 16).Value = greatestIncreaseTicker
            ws.Cells(3, 16).Value = greatestDecreaseTicker
            ws.Cells(4, 16).Value = greatestVolumeTicker
            ws.Cells(2, 17).Value = FormatPercent(greatestIncrease, 2)
            ws.Cells(3, 17).Value = FormatPercent(greatestDecrease, 2)
            ws.Cells(4, 17).Value = FormatNumber(greatestVolume)
        ElseIf ws.Name = "2020" Then
            ws.Cells(2, 16).Value = greatestIncreaseTicker
            ws.Cells(3, 16).Value = greatestDecreaseTicker
            ws.Cells(4, 16).Value = greatestVolumeTicker
            ws.Cells(2, 17).Value = FormatPercent(greatestIncrease, 2)
            ws.Cells(3, 17).Value = FormatPercent(greatestDecrease, 2)
            ws.Cells(4, 17).Value = FormatNumber(greatestVolume)
        End If
  Next ws
End Sub

