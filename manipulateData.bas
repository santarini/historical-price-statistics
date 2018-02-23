Sub manipulateData()
    Dim Rng As Range
    Dim LastRow As Integer
    
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    LastRow = Selection.Rows.Count

    Range("G1").Value = "Previous Close to Close"
    Range("G3").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]-R[-1]C[-1]"
    'populate more formulas
    
    
    
    
    Range("G3").Select
    Selection.AutoFill Destination:=Range("G3:G" & LastRow)
End Sub
