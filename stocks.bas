Option Explicit
Sub StocksAnalysis()
    
   Dim ws As Worksheet
   
   For Each ws In ThisWorkbook.Worksheets
   
   
    Dim lastRow As Long
    Dim ticker As String
    Dim nextRowTicker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim changeFraction As Double
    Dim totalVolume As LongLong
    Dim inputRow As Long
    Dim summaryRow As Long
    
    
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Set up summary table headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Volume"
    
    summaryRow = 2
    openingPrice = ws.Cells(2, 3).Value
    totalVolume = 0
    ' Loop through all the stocks
    For inputRow = 2 To lastRow
        totalVolume = totalVolume + ws.Cells(inputRow, 7).Value
        ticker = ws.Cells(inputRow, 1).Value
        nextRowTicker = ws.Cells(inputRow + 1, 1).Value
        If ticker <> nextRowTicker Then
        'Last row of current stock
            'Input
            closingPrice = ws.Cells(inputRow, 6).Value
            ' Calculations
            yearlyChange = closingPrice - openingPrice
            
            If openingPrice <> 0 Then
                changeFraction = (yearlyChange / openingPrice)
            Else
                changeFraction = 0
            End If
            
            ' Output
            ws.Cells(summaryRow, 9).Value = ticker
            ws.Cells(summaryRow, 10).Value = yearlyChange
            If yearlyChange >= 0 Then
                ws.Cells(summaryRow, 10).Interior.ColorIndex = 4 ' Green color
            Else
                ws.Cells(summaryRow, 10).Interior.ColorIndex = 3 ' Red color
            End If
            ws.Cells(summaryRow, 11).Value = FormatPercent(changeFraction)
            ws.Cells(summaryRow, 12).Value = totalVolume
 
           ' Prepare for next stock
            summaryRow = summaryRow + 1
            openingPrice = ws.Cells(inputRow + 1, 3).Value
            totalVolume = 0
        End If
    Next inputRow
    
    ' Auto-fit the columns in the summary table
    ws.Columns("I:K").AutoFit
    
    Next ws
    
End Sub

