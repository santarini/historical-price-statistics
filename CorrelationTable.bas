Option Explicit
Function CorrelationTable()

Dim Count As Integer
Dim Company1 As Range
Dim Company2 As Range
Dim TrgtRng As Range
Dim i As Integer
Dim j As Integer










'go to summary page, a4 select
Worksheets("Summary").Select
Range("A4").Select
Range(Selection, Selection.End(xlDown)).Select
Count = Selection.Rows.Count
Selection.Copy

MsgBox Count

Sheets.Add.Name = "CorrelationPage"
Worksheets("CorrelationPage").Select
Range("A2").Select
ActiveSheet.Paste
Range("A2").Select
Set Company1 = Selection

Range("B1").Select
Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
Range("B1").Select
Set Company2 = Selection

Range("B2").Select
Set TrgtRng = Selection

j = 0

For i = 0 To Count:
    Company1.Select
    Range("O3").Select
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
Next i


End Function
