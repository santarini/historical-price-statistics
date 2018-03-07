Sub summaryPageFormat()
'volume formatting
    Columns("E:Q").Select
    Selection.Style = "Comma"
    
    Columns("R:T").Select
    Selection.NumberFormat = "0.00"
    
'N formatting
    Columns("D:D").Select
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    Columns("U:U").Select
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    Columns("BC:BC").Select
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    Columns("CK:CK").Select
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    Columns("DS:DS").Select
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    
'Actual dollar formatting
    Columns("V:AH").Select
    Columns("BD:BP").Select
    Columns("CL:CX").Select
    Columns("DT:EF").Select
    
'Percent data formatting
    Columns("AM:AY").Select
    Columns("BU:CG").Select
    Columns("DC:DO").Select
    Columns("EK:EW").Select
    
'CV, Kurtosis, Skewness
    Columns("R:T").Select

    Columns("AI:AK").Select

    Columns("AZ:BB").Select

    Columns("BQ:BS").Select

    Columns("CH:CJ").Select

    Columns("CY:DA").Select

    Columns("DP:DR").Select

    Columns("EG:EI").Select

    Columns("EX:EZ").Select
    
End Sub
