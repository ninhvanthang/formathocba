Sub Macro2()
'
' Macro2 Macro
'
' Keyboard Shortcut: Ctrl+y
'
    ActiveWindow.SmallScroll Down:=18
    Range("A31:D31").Select
    Selection.Copy
    ActiveWindow.SmallScroll Down:=-9
    Range("G20:J20").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveWindow.SmallScroll Down:=18
    Range("H37:J37").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Lên l" & ChrW(7899) & "p 8"
    Range("H38:J39").Select
    ActiveWindow.SmallScroll Down:=3
    Range("A42:G42").Select
    ActiveCell.FormulaR1C1 = "'- Có ch" & ChrW(7913) & "ng ch" & ChrW(7881) & " Ngh" & ChrW(7873) & " ph" & ChrW(7893) & " thông: Không"
    Range("H42:J42").Select
    Selection.ClearContents
    Range("A43:J44").Select
    ActiveCell.FormulaR1C1 = _
        "'- " & ChrW(272) & ChrW(432) & ChrW(7907) & "c gi" & ChrW(7843) & "i th" & ChrW(432) & ChrW(7903) & "ng trong các k" & ChrW(7923) & " thi t" & ChrW(7915) & " c" & ChrW(7845) & "p huy" & ChrW(7879) & "n tr" & ChrW(7903) & " lên: Không"
    Range("A45:J46").Select
    ActiveWindow.SmallScroll Down:=6
    Range("A57:J57").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Selection.UnMerge
    Range("A57").Select
    Selection.Copy
    Range("G57:I57").Select
    Application.CutCopyMode = False
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("G57:I57").Select
    ActiveCell.FormulaR1C1 = ""
    Range("B57").Select
    Selection.Copy
    Range("A57").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("G57:I57").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = ""
    Range("A57").Select
    Selection.Copy
    Range("G57:I57").Select
    ActiveSheet.Paste
    Range("A57").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("A61:J63").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Selection.Font.Italic = True
End Sub
