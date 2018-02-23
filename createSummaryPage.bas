Sub createSummaryPage()
    Dim Rng As Range

'top headers
    Range("A1:D1").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Description"

    Range("E1:G1").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Volume"
    
    Range("H1:M1").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Previous Close to Close"
    
    Range("N1:S1").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Previous Open to Open"
    
    Range("T1:Y1").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Previous Close to Open"
        
    Range("Z1:AE1").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Intraday Open to Close"
    
'Sub headers
    Range("E2:G2").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Actual"
    
    Range("H2:J2").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Actual"
    
    Range("K2:M2").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Percent"
    
    Range("N2:P2").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Actual"
    
    Range("Q2:S2").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Percent"
    
    Range("T2:V2").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Actual"
    
    Range("W2:Y2").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Percent"
    
    Range("Z2:AB2").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Actual"
    
    Range("AC2:AE2").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Percent"
    
'Line headers

    Set Rng = Range("A3")
    Rng.Value = "Symbol"
    Rng.Offset(0, 1) = "Date Data Start"
    Rng.Offset(0, 2) = "Date Data End"
    Rng.Offset(0, 3) = "Sample Size"
    Rng.Offset(0, 4) = "Minimum"
    Rng.Offset(0, 5) = "Maximum"
    Rng.Offset(0, 6) = "Average"
    Rng.Offset(0, 7) = "Minimum"
    Rng.Offset(0, 8) = "Maximum"
    Rng.Offset(0, 9) = "Average"
    Rng.Offset(0, 10) = "Minimum"
    Rng.Offset(0, 11) = "Maximum"
    Rng.Offset(0, 12) = "Average"
    Rng.Offset(0, 13) = "Minimum"
    Rng.Offset(0, 14) = "Maximum"
    Rng.Offset(0, 15) = "Average"
    Rng.Offset(0, 16) = "Minimum"
    Rng.Offset(0, 17) = "Maximum"
    Rng.Offset(0, 18) = "Average"
    Rng.Offset(0, 19) = "Minimum"
    Rng.Offset(0, 20) = "Maximum"
    Rng.Offset(0, 21) = "Average"
    Rng.Offset(0, 22) = "Minimum"
    Rng.Offset(0, 23) = "Maximum"
    Rng.Offset(0, 24) = "Average"
    Rng.Offset(0, 25) = "Minimum"
    Rng.Offset(0, 26) = "Maximum"
    Rng.Offset(0, 27) = "Average"
    Rng.Offset(0, 28) = "Minimum"
    Rng.Offset(0, 29) = "Maximum"
    Rng.Offset(0, 30) = "Average"

'center and auto-width those three rows
    
    Range("A1:AE3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    Selection.Columns.AutoFit
    
'big borders
    Range("D:D,G:G,M:M,S:S,Y:Y,AE:AE").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

Range("A1").Select

    
End Sub
