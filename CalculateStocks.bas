Attribute VB_Name = "Module1"
Sub CalculateStocks()

'Create variables
Dim lastRow As Double
Dim currentTicker As String
Dim nextTicker As String
Dim openPrice As Double
Dim closePrice As Double
Dim yearlyChange As Double
Dim percentChange As Double
Dim totalStock As Double
Dim currentStock As Double
Dim resultRow As Long
Dim firstRowInSet As Long

For Each ws In Worksheets

'Begin Creating Summary Table

    'Identify last row
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Setup result headers
    ws.Cells(1, 9).value = "Ticker"
    ws.Cells(1, 10).value = "Yearly Change"
    ws.Cells(1, 11).value = "Percent Change"
    ws.Cells(1, 12).value = "Total Stock Volume"
    
    'Initialize summary table values
    resultRow = 2
    openPrice = 0
    closePrice = 0
    yearlyChange = 0
    percentChange = 0
    totalStock = 0
    currentStock = 0

    firstRowInSet = 1
    
    'Loop through each row
    For i = 2 To lastRow
        'Set currentTicker and nextTicker values
        currentTicker = ws.Cells(i, 1).value
        nextTicker = ws.Cells(i + 1, 1).value
        currentStock = ws.Cells(i, 7).value
            
        'Check to see if currentTicker and nextTicker are different
        If currentTicker <> nextTicker Then
            closePrice = ws.Cells(i, 6).value
            yearlyChange = closePrice - openPrice
            totalStock = totalStock + currentStock
            
            'set percent change value to 0 if the openPrice started at 0
            If openPrice = 0 Then
                percentChange = 0
            Else
                percentChange = yearlyChange / openPrice
            End If
            
            'Capture results in result summary table
            ws.Cells(resultRow, 9).value = currentTicker
            ws.Cells(resultRow, 10).value = yearlyChange
            ws.Cells(resultRow, 11).value = Format(percentChange, "Percent")
            ws.Cells(resultRow, 12).value = totalStock
            
            
            If yearlyChange < 0 Then
                ws.Cells(resultRow, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(resultRow, 10).Interior.ColorIndex = 4
            End If
            
            resultRow = resultRow + 1
            
            firstRowInSet = 1
            
            totalStock = 0
            
        Else
            totalStock = totalStock + currentStock
            
            If firstRowInSet = 1 Then
                openPrice = ws.Cells(i, 3).value
            End If
            
            firstRowInSet = firstRowInSet + 1
            
        End If
        
    Next i
    
'Begin Creating Bonus Summary Table

    'Identify last row
    lastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'Initialize variables
    Dim currentPer As Double
    Dim currentVol As Double
    Dim greatestInc As Double
    Dim greatestDec As Double
    Dim greatestVol As Double
    Dim greatestIncTicker As String
    Dim greatestDecTicker As String
    Dim greatestVolTicker As String
        
        'Setup result headers
    ws.Cells(1, 15).value = "Ticker"
    ws.Cells(1, 16).value = "Value"
    ws.Cells(2, 14).value = "Greatest % Increase"
    ws.Cells(3, 14).value = "Greatest % Decrease"
    ws.Cells(4, 14).value = "Greatest Total Volume"
    
    
    greatestInc = 0
    greatestDec = 0
    greatestVol = 0
    
    For i = 2 To lastRow
        currentPer = ws.Cells(i, 11).value
        currentVol = ws.Cells(i, 12).value
        
        If currentPer > greatestInc Then
            'If current percent change is greater than previous, then replace the greatest value
            greatestInc = currentPer
            greatestIncTicker = ws.Cells(i, 9).value
            
        End If
        
        If currentPer < greatestDec Then
            'If current percent change is less than previous, then replace the least value
            greatestDec = currentPer
            greatestDecTicker = ws.Cells(i, 9).value
            
        End If
        
        If currentVol > greatestVol Then
            'If current volumn is greater than previous, then replace the greatest volume
            greatestVol = currentVol
            greatestVolTicker = ws.Cells(i, 9).value
            
        End If
        
        
    Next i
    
    
    'Fill in bonus summary
    ws.Cells(2, 15).value = greatestIncTicker
    ws.Cells(2, 16).value = Format(greatestInc, "Percent")
    
    ws.Cells(3, 15).value = greatestDecTicker
    ws.Cells(3, 16).value = Format(greatestDec, "Percent")

    ws.Cells(4, 15).value = greatestVolTicker
    ws.Cells(4, 16).value = greatestVol
    
    ws.Cells.EntireColumn.AutoFit

Next ws

MsgBox ("Calculations Complete!")

End Sub
