Function orderDataForGraphing()
    
    'order the columns for graphing

    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("K:K").Select
    Selection.Cut
    Columns("B:B").Select
    ActiveSheet.Paste
    Columns("J:J").Select
    Selection.Cut
    Columns("C:C").Select
    ActiveSheet.Paste
    Columns("H:H").Select
    Selection.Cut
    Columns("D:D").Select
    ActiveSheet.Paste
    Columns("I:I").Select
    Selection.Cut
    Columns("E:E").Select
    ActiveSheet.Paste
    Columns("G:G").Select
    Application.CutCopyMode = False
    Selection.Cut
    Columns("F:F").Select
    ActiveSheet.Paste
    
    'add commas and dollar signs
    
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Style = "Comma"
    Selection.NumberFormat = "_(* #,##0.0_);_(* (#,##0.0);_(* ""-""??_);_(@_)"
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    Range("C2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Style = "Currency"
    Range("A1").Select
    
End Function


Sub manipulateData()
    Dim Rng As Range
    Dim LastRow As Integer
    
    'find the last row
    
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    LastRow = Selection.Rows.Count
    
    'day average average
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Day Average"
    Range("G2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=AVERAGE(RC[-4]:RC[-1])"
    Range("G2").Select
    Selection.AutoFill Destination:=Range("G2:G" & LastRow)
    
    'data manipulations
    Range("H1").Value = "Previous Close to Close"
    Range("H3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-2]-R[-1]C[-2]"
    Range("I3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]/R[-1]C[-3]"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Previous Open to Open"
    Range("J3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-7]-R[-1]C[-7]"
    Range("K3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]/R[-1]C[-8]"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Previous Close to Open"
    Range("L3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]/R[-1]C[-7]"
    Range("M3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-10]/R[-1]C[-7]"
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "Intraday Open to Close"
    Range("N3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-8]-RC[-11]"
    Range("O3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]/RC[-9]"
    Range("H3:O3").Select
    Selection.AutoFill Destination:=Range("H3:O" & LastRow)
   
End Sub
