
'go to summary page
'a4 select
'select all filled cells beneath A4
'count rows in selection
'copy selection

'create CorrelationPage
'A2 select
'paste values
'Set Company1 to A2

'b1 select
'paste values (trasnposed)
'Set Company2 to B1

'b2 select
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
        'navigate to CorrelationPage page
        'offset Company2
    'next j
    'navigate to CorrelationPage page
    'offset Company1
'next i
