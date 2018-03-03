Option Explicit
Function CorrelationTable()

Dim Count As Integer
Dim Rng1 As Range
Dim Rng2 As Range
Dim Rng3 As Range
Dim Rng4 As Range
Dim Company1 As String
Dim Company2 As String
Dim TrgtRng As Range
Dim i As Integer
Dim j As Integer
Dim CorrelationVar As Single

Worksheets("Summary").Select
Range("A4").Select
Range(Selection, Selection.End(xlDown)).Select
Count = Selection.Rows.Count
Selection.Copy


Sheets.Add.Name = "CorrelationPage"
Worksheets("CorrelationPage").Select
Range("A2").Select
ActiveSheet.Paste
Range("A2").Select
Set Rng1 = Selection
Company1 = Rng1.Value


Range("B1").Select
Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
Range("B1").Select
Set Rng2 = Selection
Company2 = Rng1.Value

Range("B2").Select
Set TrgtRng = Selection

For i = 1 To Count:
    Worksheets(Company1).Select
    Range("O3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Set Rng3 = Selection
    For j = 1 To Count
    On Error Resume Next
        Worksheets(Company2).Select
        Range("O3").Select
        Range(Selection, Selection.End(xlDown)).Select
        Set Rng4 = Selection
        CorrelationVar = Application.WorksheetFunction.Correl(Rng3, Rng4)
        Worksheets("CorrelationPage").Select
        TrgtRng = CorrelationVar
        TrgtRng.Offset(0, 1).Select
        Set TrgtRng = Selection
        Rng2.Offset(0, 1).Select
        Set Rng2 = Selection
        Company2 = Rng2
    Next j
    Worksheets("CorrelationPage").Select
    Range("B1").Select
    Set Rng2 = Selection
    Company2 = Rng1.Value
    Rng1.Offset(1, 1).Select
    Set TrgtRng = Selection
    Rng1.Offset(1, 0).Select
    Set Rng1 = Selection
    Company1 = Rng1
Next i


End Function
