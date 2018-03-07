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
    Selection.Style = "Currency"
    Columns("BD:BP").Select
    Selection.Style = "Currency"
    Columns("CL:CX").Select
    Selection.Style = "Currency"
    Columns("DT:EF").Select
    Selection.Style = "Currency"
    
'Percent data formatting
    Columns("AM:AY").Select
    Selection.Style = "Percent"
    Columns("BU:CG").Select
    Selection.Style = "Percent"
    Columns("DC:DO").Select
    Selection.Style = "Percent"
    Columns("EK:EW").Select
    Selection.Style = "Percent"
    
'CV, Kurtosis, Skewness
    Columns("R:T").Select
    Selection.NumberFormat = "0.00"
    Columns("AI:AK").Select
    Selection.NumberFormat = "0.00"
    Columns("AZ:BB").Select
    Selection.NumberFormat = "0.00"
    Columns("BQ:BS").Select
    Selection.NumberFormat = "0.00"
    Columns("CH:CJ").Select
    Selection.NumberFormat = "0.00"
    Columns("CY:DA").Select
    Selection.NumberFormat = "0.00"
    Columns("DP:DR").Select
    Selection.NumberFormat = "0.00"
    Columns("EG:EI").Select
    Selection.NumberFormat = "0.00"
    Columns("EX:EZ").Select
    Selection.NumberFormat = "0.00"
End Sub
