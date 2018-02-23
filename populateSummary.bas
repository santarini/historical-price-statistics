Function populateSummary(SummaryRng As Range)

    Dim WS As Worksheet
    Dim StartDate, EndDate As Date
    Dim Count, LastRow As Integer
    Dim VolNActual, VolMinimumActual, VolFirstQuintileActual, VolFirstDecileActual, VolLowerQuartileActual, VolMedianActual, VolUpperQuartileActual, VolLastDecileActual, VolLastQuintileActual, VolMaximumActual, VolModeActual, VolArithmeticMeanActual, VolVarianceActual, VolStandardDeviationActual, VolCoefficientOfVariationActual, VolKurtosisActual, VolSkewnessActual As Double
    Dim CtoCNActual, CtoCMinimumActual, CtoCFirstQuintileActual, CtoCFirstDecileActual, CtoCLowerQuartileActual, CtoCMedianActual, CtoCUpperQuartileActual, CtoCLastDecileActual, CtoCLastQuintileActual, CtoCMaximumActual, CtoCModeActual, CtoCArithmeticMeanActual, CtoCVarianceActual, CtoCStandardDeviationActual, CtoCCoefficientOfVariationActual, CtoCKurtosisActual, CtoCSkewnessActual, CtoCNPercent, CtoCMinimumPercent, CtoCFirstQuintilePercent, CtoCFirstDecilePercent, CtoCLowerQuartilePercent, CtoCMedianPercent, CtoCUpperQuartilePercent, CtoCLastDecilePercent, CtoCLastQuintilePercent, CtoCMaximumPercent, CtoCModePercent, CtoCArithmeticMeanPercent, CtoCVariancePercent, CtoCStandardDeviationPercent, CtoCCoefficientOfVariationPercent, CtoCKurtosisPercent, CtoCSkewnessPercent As Double
    Dim OtoONActual, OtoOMinimumActual, OtoOFirstQuintileActual, OtoOFirstDecileActual, OtoOLowerQuartileActual, OtoOMedianActual, OtoOUpperQuartileActual, OtoOLastDecileActual, OtoOLastQuintileActual, OtoOMaximumActual, OtoOModeActual, OtoOArithmeticMeanActual, OtoOVarianceActual, OtoOStandardDeviationActual, OtoOCoefficientOfVariationActual, OtoOKurtosisActual, OtoOSkewnessActual, OtoONPercent, OtoOMinimumPercent, OtoOFirstQuintilePercent, OtoOFirstDecilePercent, OtoOLowerQuartilePercent, OtoOMedianPercent, OtoOUpperQuartilePercent, OtoOLastDecilePercent, OtoOLastQuintilePercent, OtoOMaximumPercent, OtoOModePercent, OtoOArithmeticMeanPercent, OtoOVariancePercent, OtoOStandardDeviationPercent, OtoOCoefficientOfVariationPercent, OtoOKurtosisPercent, OtoOSkewnessPercent As Double
    Dim OtoCNActual, OtoCMinimumActual, OtoCFirstQuintileActual, OtoCFirstDecileActual, OtoCLowerQuartileActual, OtoCMedianActual, OtoCUpperQuartileActual, OtoCLastDecileActual, OtoCLastQuintileActual, OtoCMaximumActual, OtoCModeActual, OtoCArithmeticMeanActual, OtoCVarianceActual, OtoCStandardDeviationActual, OtoCCoefficientOfVariationActual, OtoCKurtosisActual, OtoCSkewnessActual, OtoCNPercent, OtoCMinimumPercent, OtoCFirstQuintilePercent, OtoCFirstDecilePercent, OtoCLowerQuartilePercent, OtoCMedianPercent, OtoCUpperQuartilePercent, OtoCLastDecilePercent, OtoCLastQuintilePercent, OtoCMaximumPercent, OtoCModePercent, OtoCArithmeticMeanPercent, OtoCVariancePercent, OtoCStandardDeviationPercent, OtoCCoefficientOfVariationPercent, OtoCKurtosisPercent, OtoCSkewnessPercent As Double
    Dim Rng As Range
    
    Set WS = ActiveSheet
    
    WS.Activate
    
    'get dates
    Set Rng = Range("A2")
    Rng.Select
    StartDate = Selection
    Selection.End(xlDown).Select
    EndDate = Selection

    'paste dates
    Worksheets("Summary").Activate
    SummaryRng.Value = WS.Name
    SummaryRng.Offset(0, 1).Value = StartDate
    SummaryRng.Offset(0, 2).Value = EndDate
    
    'define volume actual range
    WS.Activate
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Set Rng = Selection
    LastRow = Selection.Rows.Count
    
    'calculate volume actual stats
    VolN = LastRow
    VolMinimumVal = Application.WorksheetFunction.Min(Rng)
    VolFirstQuintile = Application.WorksheetFunction.Percentile(Rng, 0.05)
    VolFirstDecile = Application.WorksheetFunction.Percentile(Rng, 0.1)
    VolLowerQuartile = Application.WorksheetFunction.Percentile(Rng, 0.25)
    VolMedian = Application.WorksheetFunction.Median(Rng)
    VolUpperQuartile = Application.WorksheetFunction.Percentile(Rng, 0.75)
    VolLastDecile = Application.WorksheetFunction.Percentile(Rng, 0.9)
    VolLastQuintile = Application.WorksheetFunction.Percentile(Rng, 0.95)
    VolMaximumVal = Application.WorksheetFunction.Max(Rng)
    VolModeVal = Application.WorksheetFunction.Mode(Rng)
    VolArithmeticMean = Application.WorksheetFunction.Average(Rng)
    VolVariance = VolStandardDeviation * VolStandardDeviation
    VolStandardDeviation = Application.WorksheetFunction.StDev_P(Rng)
    VolCoefficientOfVariation = VolStandardDeviation / VolArithmeticMean
    VolKurtosis = Application.WorksheetFunction.Kurt(Rng)
    VolSkewness = Application.WorksheetFunction.Skew_p(Rng)

    'paste volume actual stats
    Worksheets("Summary").Activate
    SummaryRng.Offset(0, 3).Value = VolN
    SummaryRng.Offset(0, 4).Value = VolMinimumVal
    SummaryRng.Offset(0, 5).Value = VolFirstQuintile
    SummaryRng.Offset(0, 6).Value = VolFirstDecile
    SummaryRng.Offset(0, 7).Value = VolLowerQuartile
    SummaryRng.Offset(0, 8).Value = VolMedian
    SummaryRng.Offset(0, 9).Value = VolUpperQuartile
    SummaryRng.Offset(0, 10).Value = VolLastDecile
    SummaryRng.Offset(0, 11).Value = VolLastQuintile
    SummaryRng.Offset(0, 12).Value = VolMaximumVal
    SummaryRng.Offset(0, 13).Value = VolModeVal
    SummaryRng.Offset(0, 14).Value = VolArithmeticMean
    SummaryRng.Offset(0, 15).Value = VolGeometricMean
    SummaryRng.Offset(0, 16).Value = VolHarmonicMean
    SummaryRng.Offset(0, 17).Value = VolVariance
    SummaryRng.Offset(0, 18).Value = VolStandardDeviation
    SummaryRng.Offset(0, 19).Value = VolCoefficientOfVariation
    SummaryRng.Offset(0, 20).Value = VolKurtosis
    SummaryRng.Offset(0, 21).Value = VolSkewness
    
    
    
    'define Previous Close to Close actual Range
    WS.Activate
    Range("H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Set Rng = Selection
    LastRow = Selection.Rows.Count
    
    'calculate Previous Close to Close actual stats
    CtoCNActual = LastRow
    CtoCMinimumValActual = Application.WorksheetFunction.Min(Rng)
    CtoCFirstQuintileActual = Application.WorksheetFunction.Percentile(Rng, 0.05)
    CtoCFirstDecileActual = Application.WorksheetFunction.Percentile(Rng, 0.1)
    CtoCLowerQuartileActual = Application.WorksheetFunction.Percentile(Rng, 0.25)
    CtoCMedianActual = Application.WorksheetFunction.Median(Rng)
    CtoCUpperQuartileActual = Application.WorksheetFunction.Percentile(Rng, 0.75)
    CtoCLastDecileActual = Application.WorksheetFunction.Percentile(Rng, 0.9)
    CtoCLastQuintileActual = Application.WorksheetFunction.Percentile(Rng, 0.95)
    CtoCMaximumValActual = Application.WorksheetFunction.Max(Rng)
    CtoCModeValActual = Application.WorksheetFunction.Mode(Rng)
    CtoCArithmeticMeanActual = Application.WorksheetFunction.Average(Rng)
    CtoCVarianceActual = CtoCStandardDeviationActual * CtoCStandardDeviationActual
    CtoCStandardDeviationActual = Application.WorksheetFunction.StDev_P(Rng)
    CtoCCoefficientOfVariationActual = CtoCStandardDeviationActual / CtoCArithmeticMeanActual
    CtoCKurtosisActual = Application.WorksheetFunction.Kurt(Rng)
    CtoCSkewnessActual = Application.WorksheetFunction.Skew_p(Rng)

    'paste Previous Close to Close actual stats
    Worksheets("Summary").Activate
    SummaryRng.Offset(0, 22).Value = CtoCNActual
    SummaryRng.Offset(0, 23).Value = CtoCMinimumValActual
    SummaryRng.Offset(0, 24).Value = CtoCFirstQuintileActual
    SummaryRng.Offset(0, 25).Value = CtoCFirstDecileActual
    SummaryRng.Offset(0, 26).Value = CtoCLowerQuartileActual
    SummaryRng.Offset(0, 27).Value = CtoCMedianActual
    SummaryRng.Offset(0, 28).Value = CtoCUpperQuartileActual
    SummaryRng.Offset(0, 29).Value = CtoCLastDecileActual
    SummaryRng.Offset(0, 30).Value = CtoCLastQuintileActual
    SummaryRng.Offset(0, 31).Value = CtoCMaximumValActual
    SummaryRng.Offset(0, 32).Value = CtoCModeValActual
    SummaryRng.Offset(0, 33).Value = CtoCArithmeticMeanActual
    SummaryRng.Offset(0, 34).Value = CtoCGeometricMeanActual
    SummaryRng.Offset(0, 35).Value = CtoCHarmonicMeanActual
    SummaryRng.Offset(0, 36).Value = CtoCVarianceActual
    SummaryRng.Offset(0, 37).Value = CtoCStandardDeviationActual
    SummaryRng.Offset(0, 38).Value = CtoCCoefficientOfVariationActual
    SummaryRng.Offset(0, 39).Value = CtoCKurtosisActual
    SummaryRng.Offset(0, 40).Value = CtoCSkewnessActual
    
    'define Previous Close to Close percent Range
    WS.Activate
    Range("I3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Set Rng = Selection
    LastRow = Selection.Rows.Count
    
    'calculate Previous Close to Close percent stats
    CtoCNPercent = LastRow
    CtoCMinimumValPercent = Application.WorksheetFunction.Min(Rng)
    CtoCFirstQuintilePercent = Application.WorksheetFunction.Percentile(Rng, 0.05)
    CtoCFirstDecilePercent = Application.WorksheetFunction.Percentile(Rng, 0.1)
    CtoCLowerQuartilePercent = Application.WorksheetFunction.Percentile(Rng, 0.25)
    CtoCMedianPercent = Application.WorksheetFunction.Median(Rng)
    CtoCUpperQuartilePercent = Application.WorksheetFunction.Percentile(Rng, 0.75)
    CtoCLastDecilePercent = Application.WorksheetFunction.Percentile(Rng, 0.9)
    CtoCLastQuintilePercent = Application.WorksheetFunction.Percentile(Rng, 0.95)
    CtoCMaximumValPercent = Application.WorksheetFunction.Max(Rng)
    CtoCArithmeticMeanPercent = Application.WorksheetFunction.Average(Rng)
    CtoCVariancePercent = CtoCStandardDeviationPercent * CtoCStandardDeviationPercent
    CtoCStandardDeviationPercent = Application.WorksheetFunction.StDev_P(Rng)
    CtoCCoefficientOfVariationPercent = CtoCStandardDeviationPercent / CtoCArithmeticMeanPercent
    CtoCKurtosisPercent = Application.WorksheetFunction.Kurt(Rng)
    CtoCSkewnessPercent = Application.WorksheetFunction.Skew_p(Rng)

    'paste Previous Close to Close percent stats
    Worksheets("Summary").Activate
    SummaryRng.Offset(0, 41).Value = CtoCNPercent
    SummaryRng.Offset(0, 42).Value = CtoCMinimumValPercent
    SummaryRng.Offset(0, 43).Value = CtoCFirstQuintilePercent
    SummaryRng.Offset(0, 44).Value = CtoCFirstDecilePercent
    SummaryRng.Offset(0, 45).Value = CtoCLowerQuartilePercent
    SummaryRng.Offset(0, 46).Value = CtoCMedianPercent
    SummaryRng.Offset(0, 47).Value = CtoCUpperQuartilePercent
    SummaryRng.Offset(0, 48).Value = CtoCLastDecilePercent
    SummaryRng.Offset(0, 49).Value = CtoCLastQuintilePercent
    SummaryRng.Offset(0, 50).Value = CtoCMaximumValPercent
    SummaryRng.Offset(0, 51).Value = CtoCModeValPercent
    SummaryRng.Offset(0, 52).Value = CtoCArithmeticMeanPercent
    SummaryRng.Offset(0, 53).Value = CtoCGeometricMeanPercent
    SummaryRng.Offset(0, 54).Value = CtoCHarmonicMeanPercent
    SummaryRng.Offset(0, 55).Value = CtoCVariancePercent
    SummaryRng.Offset(0, 56).Value = CtoCStandardDeviationPercent
    SummaryRng.Offset(0, 57).Value = CtoCCoefficientOfVariationPercent
    SummaryRng.Offset(0, 58).Value = CtoCKurtosisPercent
    SummaryRng.Offset(0, 59).Value = CtoCSkewnessPercent


    'define Previous Open to Open actual Range
    WS.Activate
    Range("J3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Set Rng = Selection
    LastRow = Selection.Rows.Count
    
    'calculate Previous Open to Open actual stats
    OtoONActual = LastRow
    OtoOMinimumValActual = Application.WorksheetFunction.Min(Rng)
    OtoOFirstQuintileActual = Application.WorksheetFunction.Percentile(Rng, 0.05)
    OtoOFirstDecileActual = Application.WorksheetFunction.Percentile(Rng, 0.1)
    OtoOLowerQuartileActual = Application.WorksheetFunction.Percentile(Rng, 0.25)
    OtoOMedianActual = Application.WorksheetFunction.Median(Rng)
    OtoOUpperQuartileActual = Application.WorksheetFunction.Percentile(Rng, 0.75)
    OtoOLastDecileActual = Application.WorksheetFunction.Percentile(Rng, 0.9)
    OtoOLastQuintileActual = Application.WorksheetFunction.Percentile(Rng, 0.95)
    OtoOMaximumValActual = Application.WorksheetFunction.Max(Rng)
    OtoOModeValActual = Application.WorksheetFunction.Mode(Rng)
    OtoOArithmeticMeanActual = Application.WorksheetFunction.Average(Rng)
    OtoOVarianceActual = OtoOStandardDeviationActual * OtoOStandardDeviationActual
    OtoOStandardDeviationActual = Application.WorksheetFunction.StDev_P(Rng)
    OtoOCoefficientOfVariationActual = OtoOStandardDeviationActual / OtoOArithmeticMeanActual
    OtoOKurtosisActual = Application.WorksheetFunction.Kurt(Rng)
    OtoOSkewnessActual = Application.WorksheetFunction.Skew_p(Rng)

    
    'paste Previous Open to Open actual stats
    Worksheets("Summary").Activate
    SummaryRng.Offset(0, 60).Value = OtoONActual
    SummaryRng.Offset(0, 61).Value = OtoOMinimumValActual
    SummaryRng.Offset(0, 62).Value = OtoOFirstQuintileActual
    SummaryRng.Offset(0, 63).Value = OtoOFirstDecileActual
    SummaryRng.Offset(0, 64).Value = OtoOLowerQuartileActual
    SummaryRng.Offset(0, 65).Value = OtoOMedianActual
    SummaryRng.Offset(0, 66).Value = OtoOUpperQuartileActual
    SummaryRng.Offset(0, 67).Value = OtoOLastDecileActual
    SummaryRng.Offset(0, 68).Value = OtoOLastQuintileActual
    SummaryRng.Offset(0, 69).Value = OtoOMaximumValActual
    SummaryRng.Offset(0, 70).Value = OtoOModeValActual
    SummaryRng.Offset(0, 71).Value = OtoOArithmeticMeanActual
    SummaryRng.Offset(0, 72).Value = OtoOGeometricMeanActual
    SummaryRng.Offset(0, 73).Value = OtoOHarmonicMeanActual
    SummaryRng.Offset(0, 74).Value = OtoOVarianceActual
    SummaryRng.Offset(0, 75).Value = OtoOStandardDeviationActual
    SummaryRng.Offset(0, 76).Value = OtoOCoefficientOfVariationActual
    SummaryRng.Offset(0, 77).Value = OtoOKurtosisActual
    SummaryRng.Offset(0, 78).Value = OtoOSkewnessActual

    'define Previous Open to Open percent Range
    WS.Activate
    Range("K3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Set Rng = Selection
    LastRow = Selection.Rows.Count
    
    'calculate Previous Open to Open percent stats
    OtoONPercent = LastRow
    OtoOMinimumValPercent = Application.WorksheetFunction.Min(Rng)
    OtoOFirstQuintilePercent = Application.WorksheetFunction.Percentile(Rng, 0.05)
    OtoOFirstDecilePercent = Application.WorksheetFunction.Percentile(Rng, 0.1)
    OtoOLowerQuartilePercent = Application.WorksheetFunction.Percentile(Rng, 0.25)
    OtoOMedianPercent = Application.WorksheetFunction.Median(Rng)
    OtoOUpperQuartilePercent = Application.WorksheetFunction.Percentile(Rng, 0.75)
    OtoOLastDecilePercent = Application.WorksheetFunction.Percentile(Rng, 0.9)
    OtoOLastQuintilePercent = Application.WorksheetFunction.Percentile(Rng, 0.95)
    OtoOMaximumValPercent = Application.WorksheetFunction.Max(Rng)
    OtoOModeValPercent = Application.WorksheetFunction.Mode(Rng)
    OtoOArithmeticMeanPercent = Application.WorksheetFunction.Average(Rng)
    OtoOVariancePercent = OtoOStandardDeviationPercent * OtoOStandardDeviationPercent
    OtoOStandardDeviationPercent = Application.WorksheetFunction.StDev_P(Rng)
    OtoOCoefficientOfVariationPercent = OtoOStandardDeviationPercent / OtoOArithmeticMeanPercent
    OtoOKurtosisPercent = Application.WorksheetFunction.Kurt(Rng)
    OtoOSkewnessPercent = Application.WorksheetFunction.Skew_p(Rng)

    
    'paste Previous Open to Open percent stats
    Worksheets("Summary").Activate
    SummaryRng.Offset(0, 79).Value = OtoONPercent
    SummaryRng.Offset(0, 80).Value = OtoOMinimumValPercent
    SummaryRng.Offset(0, 81).Value = OtoOFirstQuintilePercent
    SummaryRng.Offset(0, 82).Value = OtoOFirstDecilePercent
    SummaryRng.Offset(0, 83).Value = OtoOLowerQuartilePercent
    SummaryRng.Offset(0, 84).Value = OtoOMedianPercent
    SummaryRng.Offset(0, 85).Value = OtoOUpperQuartilePercent
    SummaryRng.Offset(0, 86).Value = OtoOLastDecilePercent
    SummaryRng.Offset(0, 87).Value = OtoOLastQuintilePercent
    SummaryRng.Offset(0, 88).Value = OtoOMaximumValPercent
    SummaryRng.Offset(0, 89).Value = OtoOModeValPercent
    SummaryRng.Offset(0, 90).Value = OtoOArithmeticMeanPercent
    SummaryRng.Offset(0, 91).Value = OtoOGeometricMeanPercent
    SummaryRng.Offset(0, 92).Value = OtoOHarmonicMeanPercent
    SummaryRng.Offset(0, 93).Value = OtoOVariancePercent
    SummaryRng.Offset(0, 94).Value = OtoOStandardDeviationPercent
    SummaryRng.Offset(0, 95).Value = OtoOCoefficientOfVariationPercent
    SummaryRng.Offset(0, 96).Value = OtoOKurtosisPercent
    SummaryRng.Offset(0, 97).Value = OtoOSkewnessPercent

    'define Previous Close to Open actual Range
    WS.Activate
    Range("L3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Set Rng = Selection
    LastRow = Selection.Rows.Count
    
    'calculate Previous Close to Open actual stats
    CtoONActual = LastRow
    CtoOMinimumValActual = Application.WorksheetFunction.Min(Rng)
    CtoOFirstQuintileActual = Application.WorksheetFunction.Percentile(Rng, 0.05)
    CtoOFirstDecileActual = Application.WorksheetFunction.Percentile(Rng, 0.1)
    CtoOLowerQuartileActual = Application.WorksheetFunction.Percentile(Rng, 0.25)
    CtoOMedianActual = Application.WorksheetFunction.Median(Rng)
    CtoOUpperQuartileActual = Application.WorksheetFunction.Percentile(Rng, 0.75)
    CtoOLastDecileActual = Application.WorksheetFunction.Percentile(Rng, 0.9)
    CtoOLastQuintileActual = Application.WorksheetFunction.Percentile(Rng, 0.95)
    CtoOMaximumValActual = Application.WorksheetFunction.Max(Rng)
    CtoOModeValActual = Application.WorksheetFunction.Mode(Rng)
    CtoOArithmeticMeanActual = Application.WorksheetFunction.Average(Rng)
    CtoOVarianceActual = CtoOStandardDeviationActual * CtoOStandardDeviationActual
    CtoOStandardDeviationActual = Application.WorksheetFunction.StDev_P(Rng)
    CtoOCoefficientOfVariationActual = CtoOStandardDeviationActual / CtoOArithmeticMeanActual
    CtoOKurtosisActual = Application.WorksheetFunction.Kurt(Rng)
    CtoOSkewnessActual = Application.WorksheetFunction.Skew_p(Rng)

    
    'paste Previous Close to Open actual stats
    Worksheets("Summary").Activate
    SummaryRng.Offset(0, 98).Value = CtoONActual
    SummaryRng.Offset(0, 99).Value = CtoOMinimumValActual
    SummaryRng.Offset(0, 100).Value = CtoOFirstQuintileActual
    SummaryRng.Offset(0, 101).Value = CtoOFirstDecileActual
    SummaryRng.Offset(0, 102).Value = CtoOLowerQuartileActual
    SummaryRng.Offset(0, 103).Value = CtoOMedianActual
    SummaryRng.Offset(0, 104).Value = CtoOUpperQuartileActual
    SummaryRng.Offset(0, 105).Value = CtoOLastDecileActual
    SummaryRng.Offset(0, 106).Value = CtoOLastQuintileActual
    SummaryRng.Offset(0, 107).Value = CtoOMaximumValActual
    SummaryRng.Offset(0, 108).Value = CtoOModeValActual
    SummaryRng.Offset(0, 109).Value = CtoOArithmeticMeanActual
    SummaryRng.Offset(0, 110).Value = CtoOGeometricMeanActual
    SummaryRng.Offset(0, 111).Value = CtoOHarmonicMeanActual
    SummaryRng.Offset(0, 112).Value = CtoOVarianceActual
    SummaryRng.Offset(0, 113).Value = CtoOStandardDeviationActual
    SummaryRng.Offset(0, 114).Value = CtoOCoefficientOfVariationActual
    SummaryRng.Offset(0, 115).Value = CtoOKurtosisActual
    SummaryRng.Offset(0, 116).Value = CtoOSkewnessActual

    'define Previous Close to Open percent Range
    WS.Activate
    Range("M3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Set Rng = Selection
    LastRow = Selection.Rows.Count
    
    'calculate Previous Close to Open percent stats
    CtoONPercent = LastRow
    CtoOMinimumValPercent = Application.WorksheetFunction.Min(Rng)
    CtoOFirstQuintilePercent = Application.WorksheetFunction.Percentile(Rng, 0.05)
    CtoOFirstDecilePercent = Application.WorksheetFunction.Percentile(Rng, 0.1)
    CtoOLowerQuartilePercent = Application.WorksheetFunction.Percentile(Rng, 0.25)
    CtoOMedianPercent = Application.WorksheetFunction.Median(Rng)
    CtoOUpperQuartilePercent = Application.WorksheetFunction.Percentile(Rng, 0.75)
    CtoOLastDecilePercent = Application.WorksheetFunction.Percentile(Rng, 0.9)
    CtoOLastQuintilePercent = Application.WorksheetFunction.Percentile(Rng, 0.95)
    CtoOMaximumValPercent = Application.WorksheetFunction.Max(Rng)
    CtoOModeValPercent = Application.WorksheetFunction.Mode(Rng)
    CtoOArithmeticMeanPercent = Application.WorksheetFunction.Average(Rng)
    CtoOVariancePercent = CtoOStandardDeviationPercent * CtoOStandardDeviationPercent
    CtoOStandardDeviationPercent = Application.WorksheetFunction.StDev_P(Rng)
    CtoOCoefficientOfVariationPercent = CtoOStandardDeviationPercent / CtoOArithmeticMeanPercent
    CtoOKurtosisPercent = Application.WorksheetFunction.Kurt(Rng)
    CtoOSkewnessPercent = Application.WorksheetFunction.Skew_p(Rng)

    
    'paste Previous Close to Open percent stats
    Worksheets("Summary").Activate
    SummaryRng.Offset(0, 117).Value = CtoONPercent
    SummaryRng.Offset(0, 118).Value = CtoOMinimumValPercent
    SummaryRng.Offset(0, 119).Value = CtoOFirstQuintilePercent
    SummaryRng.Offset(0, 120).Value = CtoOFirstDecilePercent
    SummaryRng.Offset(0, 121).Value = CtoOLowerQuartilePercent
    SummaryRng.Offset(0, 122).Value = CtoOMedianPercent
    SummaryRng.Offset(0, 123).Value = CtoOUpperQuartilePercent
    SummaryRng.Offset(0, 124).Value = CtoOLastDecilePercent
    SummaryRng.Offset(0, 125).Value = CtoOLastQuintilePercent
    SummaryRng.Offset(0, 126).Value = CtoOMaximumValPercent
    SummaryRng.Offset(0, 127).Value = CtoOModeValPercent
    SummaryRng.Offset(0, 128).Value = CtoOArithmeticMeanPercent
    SummaryRng.Offset(0, 129).Value = CtoOGeometricMeanPercent
    SummaryRng.Offset(0, 130).Value = CtoOHarmonicMeanPercent
    SummaryRng.Offset(0, 131).Value = CtoOVariancePercent
    SummaryRng.Offset(0, 132).Value = CtoOStandardDeviationPercent
    SummaryRng.Offset(0, 133).Value = CtoOCoefficientOfVariationPercent
    SummaryRng.Offset(0, 134).Value = CtoOKurtosisPercent
    SummaryRng.Offset(0, 135).Value = CtoOSkewnessPercent

    
    'define Intraday Open to Close actual Range
    WS.Activate
    Range("N3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Set Rng = Selection
    LastRow = Selection.Rows.Count
    
    'calculate Intraday Open to Close actual stats
    OtoCNActual = LastRow
    OtoCMinimumValActual = Application.WorksheetFunction.Min(Rng)
    OtoCFirstQuintileActual = Application.WorksheetFunction.Percentile(Rng, 0.05)
    OtoCFirstDecileActual = Application.WorksheetFunction.Percentile(Rng, 0.1)
    OtoCLowerQuartileActual = Application.WorksheetFunction.Percentile(Rng, 0.25)
    OtoCMedianActual = Application.WorksheetFunction.Median(Rng)
    OtoCUpperQuartileActual = Application.WorksheetFunction.Percentile(Rng, 0.75)
    OtoCLastDecileActual = Application.WorksheetFunction.Percentile(Rng, 0.9)
    OtoCLastQuintileActual = Application.WorksheetFunction.Percentile(Rng, 0.95)
    OtoCMaximumValActual = Application.WorksheetFunction.Max(Rng)
    OtoCModeValActual = Application.WorksheetFunction.Mode(Rng)
    OtoCArithmeticMeanActual = Application.WorksheetFunction.Average(Rng)
    OtoCVarianceActual = OtoCStandardDeviationActual * OtoCStandardDeviationActual
    OtoCStandardDeviationActual = Application.WorksheetFunction.StDev_P(Rng)
    OtoCCoefficientOfVariationActual = OtoCStandardDeviationActual / OtoCArithmeticMeanActual
    OtoCKurtosisActual = Application.WorksheetFunction.Kurt(Rng)
    OtoCSkewnessActual = Application.WorksheetFunction.Skew_p(Rng)

    
    'paste Intraday Open to Close actual stats
    Worksheets("Summary").Activate
    SummaryRng.Offset(0, 136).Value = OtoCNActual
    SummaryRng.Offset(0, 137).Value = OtoCMinimumValActual
    SummaryRng.Offset(0, 138).Value = OtoCFirstQuintileActual
    SummaryRng.Offset(0, 139).Value = OtoCFirstDecileActual
    SummaryRng.Offset(0, 140).Value = OtoCLowerQuartileActual
    SummaryRng.Offset(0, 141).Value = OtoCMedianActual
    SummaryRng.Offset(0, 142).Value = OtoCUpperQuartileActual
    SummaryRng.Offset(0, 143).Value = OtoCLastDecileActual
    SummaryRng.Offset(0, 144).Value = OtoCLastQuintileActual
    SummaryRng.Offset(0, 145).Value = OtoCMaximumValActual
    SummaryRng.Offset(0, 146).Value = OtoCModeValActual
    SummaryRng.Offset(0, 147).Value = OtoCArithmeticMeanActual
    SummaryRng.Offset(0, 148).Value = OtoCGeometricMeanActual
    SummaryRng.Offset(0, 149).Value = OtoCHarmonicMeanActual
    SummaryRng.Offset(0, 150).Value = OtoCVarianceActual
    SummaryRng.Offset(0, 151).Value = OtoCStandardDeviationActual
    SummaryRng.Offset(0, 152).Value = OtoCCoefficientOfVariationActual
    SummaryRng.Offset(0, 153).Value = OtoCKurtosisActual
    SummaryRng.Offset(0, 154).Value = OtoCSkewnessActual
    
    'define Intraday Open to Close percent Range
    WS.Activate
    Range("O3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Set Rng = Selection
    LastRow = Selection.Rows.Count
    
    'calculate Previous Open to Close percent stats
    OtoCNPercent = LastRow
    OtoCMinimumValPercent = Application.WorksheetFunction.Min(Rng)
    OtoCFirstQuintilePercent = Application.WorksheetFunction.Percentile(Rng, 0.05)
    OtoCFirstDecilePercent = Application.WorksheetFunction.Percentile(Rng, 0.1)
    OtoCLowerQuartilePercent = Application.WorksheetFunction.Percentile(Rng, 0.25)
    OtoCMedianPercent = Application.WorksheetFunction.Median(Rng)
    OtoCUpperQuartilePercent = Application.WorksheetFunction.Percentile(Rng, 0.75)
    OtoCLastDecilePercent = Application.WorksheetFunction.Percentile(Rng, 0.9)
    OtoCLastQuintilePercent = Application.WorksheetFunction.Percentile(Rng, 0.95)
    OtoCMaximumValPercent = Application.WorksheetFunction.Max(Rng)
    OtoCModeValPercent = Application.WorksheetFunction.Mode(Rng)
    OtoCArithmeticMeanPercent = Application.WorksheetFunction.Average(Rng)
    OtoCVariancePercent = OtoCStandardDeviationPercent * OtoCStandardDeviationPercent
    OtoCStandardDeviationPercent = Application.WorksheetFunction.StDev_P(Rng)
    OtoCCoefficientOfVariationPercent = OtoCStandardDeviationPercent / OtoCArithmeticMeanPercent
    OtoCKurtosisPercent = Application.WorksheetFunction.Kurt(Rng)
    OtoCSkewnessPercent = Application.WorksheetFunction.Skew_p(Rng)

    
    'paste Intraday Open to Close percent stats
    Worksheets("Summary").Activate
    SummaryRng.Offset(0, 155).Value = OtoCNPercent
    SummaryRng.Offset(0, 156).Value = OtoCMinimumValPercent
    SummaryRng.Offset(0, 157).Value = OtoCFirstQuintilePercent
    SummaryRng.Offset(0, 158).Value = OtoCFirstDecilePercent
    SummaryRng.Offset(0, 159).Value = OtoCLowerQuartilePercent
    SummaryRng.Offset(0, 160).Value = OtoCMedianPercent
    SummaryRng.Offset(0, 161).Value = OtoCUpperQuartilePercent
    SummaryRng.Offset(0, 162).Value = OtoCLastDecilePercent
    SummaryRng.Offset(0, 163).Value = OtoCLastQuintilePercent
    SummaryRng.Offset(0, 164).Value = OtoCMaximumValPercent
    SummaryRng.Offset(0, 165).Value = OtoCModeValPercent
    SummaryRng.Offset(0, 166).Value = OtoCArithmeticMeanPercent
    SummaryRng.Offset(0, 167).Value = OtoCGeometricMeanPercent
    SummaryRng.Offset(0, 168).Value = OtoCHarmonicMeanPercent
    SummaryRng.Offset(0, 169).Value = OtoCVariancePercent
    SummaryRng.Offset(0, 170).Value = OtoCStandardDeviationPercent
    SummaryRng.Offset(0, 171).Value = OtoCCoefficientOfVariationPercent
    SummaryRng.Offset(0, 172).Value = OtoCKurtosisPercent
    SummaryRng.Offset(0, 173).Value = OtoCSkewnessPercent
    
End Function
