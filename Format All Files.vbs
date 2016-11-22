Sub ProcessFiles()
    Dim Filename, Pathname As String
    Dim wb As Workbook
    Dim XlsFolder As String
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    XlsFolder = "C:\Users\nholt2\Desktop\Automation\Tables from R\Formatted\"
    Pathname = "C:\Users\nholt2\Desktop\Automation\Tables from R\"
    Filename = Dir(Pathname & "*.csv")
    Do While Filename <> ""
        Set wb = Workbooks.Open(Pathname & Filename)
        DoWork wb
        wb.SaveAs XlsFolder & Replace(Filename, ".csv", ""), ThisWorkbook.FileFormat
        wb.Close SaveChanges:=True
        Filename = Dir()
    On Error GoTo 0
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Loop
End Sub

Sub DoWork(wb As Workbook)
    With wb
        wb.Application.CutCopyMode = True
        Application.Run ("FormatMacro")
        wb.Application.CutCopyMode = False
    End With
End Sub

Sub FormatMacro()
'
' FormatMacro Macro
' Formats Tables for Word Document Display
'
' Keyboard Shortcut: Ctrl+m
'

    Range("M1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("M2").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("N1:P1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("Q1:S1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("N1:P1").Select
    ActiveCell.FormulaR1C1 = "Num. of Endorsements"
    Range("Q1:S1").Select
    ActiveCell.FormulaR1C1 = "Percentage who Endorsed"
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "Female"
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "Male"
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "Female"
    Range("R2").Select
    ActiveCell.FormulaR1C1 = "Male"
    Range("S2").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("M3").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("M4").Select
    ActiveCell.FormulaR1C1 = "Latin America & Caribbean"
    Range("M5").Select
    ActiveCell.FormulaR1C1 = "Colombia"
    Range("M6").Select
    ActiveCell.FormulaR1C1 = "Dominican Republic (Santo Domingo)"
    Range("M7").Select
    ActiveCell.FormulaR1C1 = "Ecuador"
    Range("M8").Select
    ActiveCell.FormulaR1C1 = "Guayaquil"
    Range("M9").Select
    ActiveCell.FormulaR1C1 = "Quito"
    Range("M10").Select
    ActiveCell.FormulaR1C1 = "Guatemala"
    Range("M11").Select
    ActiveCell.FormulaR1C1 = "Honduras (San Pedro Sula)"
    Range("M12").Select
    ActiveCell.FormulaR1C1 = "Mexico (Jalisco)"
    Range("M13").Select
    ActiveCell.FormulaR1C1 = "Asia & Africa"
    Range("M14").Select
    ActiveCell.FormulaR1C1 = "India"
    Range("M15").Select
    ActiveCell.FormulaR1C1 = "Delhi"
    Range("M16").Select
    ActiveCell.FormulaR1C1 = "Sahay"
    Range("M17").Select
    ActiveCell.FormulaR1C1 = "Philippines"
    Range("M18").Select
    ActiveCell.FormulaR1C1 = "Bicol"
    Range("M19").Select
    ActiveCell.FormulaR1C1 = "Manila"
    Range("M20").Select
    ActiveCell.FormulaR1C1 = "Quezon City"
    Range("M21").Select
    ActiveCell.FormulaR1C1 = "Zambia"
    Range("M22").Select
    ActiveCell.FormulaR1C1 = "USA (Little Rock)"
    Range("M23").Select
    Columns("M:M").EntireColumn.AutoFit
    Range("M3:M22").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("N1:S2").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    Range("M1:S22").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("N3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC13, R[-1]C[-13]:R[17]C[-3],2,FALSE)"
    Range("N3").Select
    Selection.AutoFill Destination:=Range("N3:N22"), Type:=xlFillDefault
    Range("N3:N22").Select
    Range("N4").Select
    Range("N3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC13, R2C1:R20C11,2,FALSE)"
    Range("N3").Select
    Selection.AutoFill Destination:=Range("N3:N22"), Type:=xlFillDefault
    Range("N3:N22").Select
    Range("N4").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC13, R2C1:R20C11,2,FALSE)"
    Range("N3").Select
    Selection.Copy
    Range("O3").Select
    ActiveSheet.Paste
    Range("P3").Select
    ActiveSheet.Paste
    Range("Q3").Select
    ActiveSheet.Paste
    Range("R3").Select
    ActiveSheet.Paste
    Range("S3").Select
    ActiveSheet.Paste
    Range("O3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC13, R2C1:R20C11,3,FALSE)"
    Range("P3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC13, R2C1:R20C11,4,FALSE)"
    Range("P3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC13, R2C1:R20C11,5,FALSE)"
    Range("Q3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC13, R2C1:R20C11,9,FALSE)"
    Range("R3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC13, R2C1:R20C11,10,FALSE)"
    Range("S3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC13, R2C1:R20C11,11,FALSE)"
    Range("O3:S3").Select
    Selection.AutoFill Destination:=Range("O3:S22"), Type:=xlFillDefault
    Range("O3:S22").Select
    Range("Q3:S22").Select
    Selection.NumberFormat = "0%"
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("N3:P22").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("N1:S1").Select
    Selection.Font.Bold = True
    Range("M3:M4").Select
    Selection.Font.Bold = True
    Range("M13").Select
    Selection.Font.Bold = True
    Range("M22").Select
    Selection.Font.Bold = True
    Range("M3:S3").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    Range("M4:S4").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
    Range("M13:S13").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
    Range("M22:S22").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
    Range("M1:S22").Select
    Selection.Copy
    
End Sub

