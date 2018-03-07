Sub summaryPageFormat()
    Columns("D:D").Select
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"

    Columns("E:Q").Select
    Selection.Style = "Comma"
    
    Columns("R:T").Select
    Selection.NumberFormat = "0.00"
    
    Columns("U:U").Select
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"

    Columns("V:AH").Select
    Selection.Style = "Percent"


    Columns("AI:AK").Select
    Selection.NumberFormat = "0.00"
    
    Columns("AL:AL").Select
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    
End Sub
