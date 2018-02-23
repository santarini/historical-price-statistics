Sub orderDataForGraphing()
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
End Sub
