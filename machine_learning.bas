Attribute VB_Name = "machine_learning"
Sub classify()
'This sub aims to populate the classification column of the data on the lower region.
'Most data here are hard-coded (e.g. Range dimensions) since this is for presentation purposes only.
'The focus is to showcase a classification algorithm
    
    'Store data on variants
    Dim Y, x_dates, x_amounts As Variant
    Y = Range("B2:E5").Value

    x_dates = Range("B10:B16").Value
    x_amounts = Range("D10:D16").Value
    
    
    Dim cost, memo(1 To 7) As Variant
    Dim index As Integer
    
    For i = 1 To 4
        'Populate memo with cost (h(x) - y)^2
        For j = 1 To 7
        
            f1 = Application.WorksheetFunction.Power(DateDiff("d", Y(i, 1), x_dates(j, 1)), 2)
            
            'f2 (USD Equivalent) is scaled to % change
            f2 = (-Y(i, 3) - x_amounts(j, 1)) / -Y(i, 3)
            f2 = Application.WorksheetFunction.Power(f2, 2)
            
            cost = f1 + f2
            memo(j) = cost
        Next j

        'Get the least amount of cost
        index = WorksheetFunction.Match(Application.WorksheetFunction.Min(memo), memo, 0)
        
        'Output
        Range("E" & 9 + index).Value = Y(i, 4)
    Next i
    

End Sub
