Sub summaryTable()

    Dim ticker As String
    
    Dim openingPrice As Double
    
    Dim closingPrice As Double
    
    Dim yearlyChange As Double
    
    Dim percentChange As Double
    
    Dim stockVolume As Single
    
    Dim maxIncrease As Double
    
    Dim maxDecrease As Double
    
    Dim maxStockVolume As Single
    
    'define the last row variable
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'create the column headers for the summary table
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    'create the column headers for the 2nd summary table
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    'keep track of the location for each row in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    Dim Summary_Table As Range
    
            For i = 2 To lastrow
            
            stockVolume = stockVolume + Cells(i, 7).Value
            
            If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
            
            openingPrice = Cells(i, 3).Value
            
            'check if the cell in the next row is different
            ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value And openingPrice > 0 Then
                       
            ticker = Cells(i, 1).Value
            
            closingPrice = Cells(i, 6).Value
            
            yearlyChange = yearlyChange + closingPrice - openingPrice
            
            percentChange = yearlyChange / openingPrice
    
            'print the ticker in the summary table
            Range("I" & Summary_Table_Row).Value = ticker
            
            'print the yearly change in the summary table
            Range("J" & Summary_Table_Row).Value = yearlyChange
    
            'print the percent change in the summary table
            Range("K" & Summary_Table_Row).Value = percentChange
            Range("K" & Summary_Table_Row).Style = "Percent"
            Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
    
            'print the total stock volume in the summary table
            Range("L" & Summary_Table_Row).Value = stockVolume
            
            'add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            
            yearlyChange = 0
            
            percentChange = 0
            
            stockVolume = 0
            
            End If
            
        Next i
        
    'print the 2nd summary table
    Cells(2, 17).Value = WorksheetFunction.Max(Range("K:K"))
    Cells(2, 17).Style = "Percent"
    Cells(2, 17).NumberFormat = "0.00%"
    Cells(3, 17).Value = WorksheetFunction.Min(Range("K:K"))
    Cells(3, 17).Style = "Percent"
    Cells(3, 17).NumberFormat = "0.00%"
    Cells(4, 17).Value = WorksheetFunction.Max(Range("L:L"))
        
    'add conditional formatting
    Columns("J:J").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
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
    Range("J1").Select
    Selection.FormatConditions.Delete
    
End Sub
