Attribute VB_Name = "Module1"
Sub summaryTable()

    'define the variables
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim stockVolume As Single
    
    'loop through each worksheet
    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
        'keep track of the location for each row in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2

        'define the last row variable
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'create the column headers for the summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'create the column headers for the 2nd summary table
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        For i = 2 To lastrow
        
            stockVolume = stockVolume + ws.Cells(i, 7).Value
            
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                openingPrice = ws.Cells(i, 3).Value
            
            'check if the cell in the next row is different
            ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value And openingPrice > 0 Then
                       
                ticker = ws.Cells(i, 1).Value
                
                closingPrice = ws.Cells(i, 6).Value
                
                yearlyChange = yearlyChange + closingPrice - openingPrice
                
                percentChange = yearlyChange / openingPrice
        
                'print the ticker in the summary table
                ws.Range("I" & Summary_Table_Row).Value = ticker
                
                'print the yearly change in the summary table
                ws.Range("J" & Summary_Table_Row).Value = yearlyChange
        
                'print the percent change in the summary table
                ws.Range("K" & Summary_Table_Row).Value = percentChange
                ws.Range("K" & Summary_Table_Row).Style = "Percent"
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        
                'print the total stock volume in the summary table
                ws.Range("L" & Summary_Table_Row).Value = stockVolume
                
                'add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                
                yearlyChange = 0
                
                percentChange = 0
                
                stockVolume = 0
                
            End If
        
        Next i
        
        Dim maxIncrease As Double
        maxIncrease = WorksheetFunction.Max(ws.Range("K:K"))
        
        Dim maxDecrease As Double
        maxDecrease = WorksheetFunction.Min(ws.Range("K:K"))
        
        Dim maxStockVolume As Single
        maxStockVolume = WorksheetFunction.Max(ws.Range("L:L"))
        
        For j = 2 To lastrow
        
            If ws.Cells(j, 11) = maxIncrease Then
            
            ws.Cells(2, 16).Value = ws.Cells(j, 9).Value
            ws.Cells(2, 17).Value = maxIncrease
            ws.Cells(2, 17).Style = "Percent"
            ws.Cells(2, 17).NumberFormat = "0.00%"
            
            ElseIf ws.Cells(j, 11) = maxDecrease Then
            
            ws.Cells(3, 16).Value = ws.Cells(j, 9).Value
            ws.Cells(3, 17).Value = maxDecrease
            ws.Cells(3, 17).Style = "Percent"
            ws.Cells(3, 17).NumberFormat = "0.00%"
            
            ElseIf ws.Cells(j, 12) = maxStockVolume Then
            
            ws.Cells(4, 16).Value = ws.Cells(j, 9).Value
            ws.Cells(4, 17).Value = maxStockVolume
            
            End If
            
        Next j
        
            ws.Activate
            
            ws.Columns("J:J").Select
            Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
                Formula1:="=0"
            Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
            With Selection.FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .Color = 5287936
                .TintAndShade = 0
            End With
            Selection.FormatConditions(1).StopIfTrue = False
            Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
                Formula1:="=0"
            Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
            With Selection.FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .Color = 255
                .TintAndShade = 0
            End With
            Selection.FormatConditions(1).StopIfTrue = False
            ws.Range("J1").Select
            Selection.FormatConditions.Delete
    
    Next ws
    
End Sub