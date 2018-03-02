'go to summary page
'set Company1 and Company2 to A4
'count rows in column

'create CorrelationPage
'set TrgtRng

'for i to count
    'define Rng1 path using Company1
    'go to Company1 Sheet
    'set Rng1 Range
    'for j to count
        'define Rng2 path using Company2
        'go to Company2 Sheet
        'set Rng2 Range
        'find the correlation of the two variables
        'CorrelationVar = Application.WorksheetFunction.Correl(Rng1, Rng2)
        'navigate to the CorrelationPage
        'paste the CorrelationVar value into TrgtRng
        'navigate to summary page
        'offset Company2
    'next j
    'navigate to summary page
    'offset Company1
'next i
