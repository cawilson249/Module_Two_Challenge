Attribute VB_Name = "Module1"
Sub StockAnalysis():
' Alpha Testing
    ' total stock volume
        Dim total As Double
        
    ' loop control variable going through all sheets
        Dim row, rowCount As Long
        
    ' variable to hold yearly change for each stocks in a sheet
        Dim yearlyChange As Double
        
    ' variable that holds percent change for each stock
        Dim percentChange As Double
        
    ' holds the row of summary table row
        Dim summaryTableRow As Long
        
    ' holds where the stock starts in sheet
        Dim StockStart As Long
   
    ' loop through all worksheets
        For Each ws In Worksheets
        
            
        ' set title row
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        ' initialize values
        summaryTableRow = 0
        total = 0
        yearlyChange = 0
        StockStart = 2 ' first stock in sheet starts on 2nd row
        
        ' get value of last row
        rowCount = ws.Cells(Rows.Count, "A").End(xlUp).row
        
        'loop until the end of sheet is reached
        For row = 2 To rowCount
        
            ' check to see if there are changes in column A
            If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
            
            ' calculate the total one last time
                total = total + ws.Cells(row, 7).Value
            
            ' check if total Volume is 0
            If total = 0 Then
                ws.Range("I" & 2 + summaryTableRow).Value = ws.Cells(row, 1).Value
                ws.Range("J" & 2 + summaryTableRow).Value = 0
                ws.Range("K" & 2 + summaryTableRow).Value = 0
                ws.Range("L" & 2 + summaryTableRow).Value = 0
            Else
                ' find first non zero starting value
                If ws.Cells(StockStart, 3).Value = 0 Then
                    For FindValue = StockStart To row
                        If ws.Cells(FindValue, 3).Value <> 0 Then
                            StockStart = FindValue
                            Exit For
                        End If
                    Next FindValue
                End If
                
                ' Find the yearly change= last closed - first open
                yearlyChange = (ws.Cells(row, 6).Value - ws.Cells(StockStart, 3).Value)
                ' find percent change (yearly change /first open)
                percentChange = yearlyChange / ws.Cells(StockStart, 3).Value
                
                ' print outcome into columns
                ws.Range("I" & 2 + summaryTableRow).Value = ws.Cells(row, 1).Value
                ws.Range("J" & 2 + summaryTableRow).Value = yearlyChange
                ws.Range("J" & 2 + summaryTableRow).NumberFormat = "0.00"
                ws.Range("K" & 2 + summaryTableRow).Value = percentChange
                ws.Range("K" & 2 + summaryTableRow).NumberFormat = "0.00"
                ws.Range("L" & 2 + summaryTableRow).Value = total
                ws.Range("L" & 2 + summaryTableRow).NumberFormat = "#,###"
                
                ' format the yearly change column
                If yearlyChange > 0 Then
                    ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 4 ' changes to green
                ElseIf yearlyChange < 0 Then
                    ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 3 ' changes to red
                Else
                    ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 0 ' changes to white
                End If
                
            End If
            
            ' reset total to 0
            total = 0
            ' resent yearly change to 0
            yealryChange = 0
            ' move to summary table
            
            summaryTableRow = summaryTableRow + 1
            
            
            ' if ticker is the same
            Else
                total = total + ws.Cells(row, 7).Value
                
            End If
        
        Next row
        
        ' after looping through rows, find max and min to place in Q2 - 4
        ' greatest inc
        ws.Range("Q2").Value = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowCount)) * 100
        ' greatest dec
        ws.Range("Q3").Value = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowCount)) * 100
        ' value of greatest Max volume
         ws.Range("Q4").Value = "%" & WorksheetFunction.Max(ws.Range("L2:L" & rowCount)) * 100
         ws.Range("Q4").NumberFormat = "#,###"
         
              
        
        ' match ticker names with values
        ' matching the increase
        increaseNumber = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        ws.Range("P2").Value = ws.Cells(increaseNumber + 1, 9)
        ' matching decrease
        decreaseNumber = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        ws.Range("P3").Value = ws.Cells(decreaseNumber + 1, 9)
        ' matching greatest volume
        volNumber = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0)
        ws.Range("P4").Value = ws.Cells(volNumber + 1, 9)
        ws.Range("P4").NumberFormat = "#,###"
                
        
        ' Autofit the title Cells
        ws.Columns("A:Q").AutoFit
    
    Next ws
    
End Sub
