Attribute VB_Name = "Module1"
Sub CBI_BANK_STAT()
    Range("A1:A24").Delete Shift:=xlUp
    Columns("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(11, 1), Array(23, 1), Array(53, 1), Array(66, 1), _
        Array(90, 1), Array(112, 1)), TrailingMinusNumbers:=True
    Columns("A:G").EntireColumn.AutoFit
    Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1:G1").Value = Array("DATE", "DATE", "Perticulars", "CHEQ No.", "Debit", "credit", "balance")
    With Range("A1:G1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("G:G").Replace What:="dr", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    With Application.ReplaceFormat.Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    FindAndDeleteRows
    newFindAndDeleteRows
    closedelete
    newcode1
    gapfinder
    Range("A1:I474").Select
    CreateTable1
    
    
End Sub

Sub FindAndDeleteRows()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim firstAddress As String

    Set ws = ActiveSheet
    Set rng = ws.Range("C:C").Find("carried forward", LookIn:=xlValues, LookAt:=xlPart)
    If rng Is Nothing Then Exit Sub
    firstAddress = rng.Address
    Do
        rng.Select
        DeleteSixRows
        Set rng = ws.Range("C:C").FindNext(rng)
    Loop While Not rng Is Nothing And rng.Address <> firstAddress
End Sub

Sub DeleteSixRows()
    Dim startRow As Integer
    Dim cell As Range

    For Each cell In Selection.Cells
        startRow = cell.Row + 1
        Rows(startRow & ":" & startRow + 11).EntireRow.Delete
    Next cell
End Sub

Sub newFindAndDeleteRows()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim firstAddress As String
    Dim rngCollection As New Collection

    Set ws = ActiveSheet
    Set rng = ws.Range("C:C").Find("carried forward", LookIn:=xlValues, LookAt:=xlPart)
    If rng Is Nothing Then Exit Sub
    firstAddress = rng.Address
    rngCollection.Add rng
    Do
        Set rng = ws.Range("C:C").FindNext(rng)
        If rng Is Nothing Or rng.Address = firstAddress Then Exit Do
        rngCollection.Add rng
    Loop
    For i = rngCollection.Count To 1 Step -1
        rngCollection(i).EntireRow.Delete
    Next i
End Sub

Sub closedelete()
Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim firstAddress As String

    Set ws = ActiveSheet
    Set rng = ws.Range("C:C").Find("closing balance", LookIn:=xlValues, LookAt:=xlPart)
    If rng Is Nothing Then Exit Sub
    firstAddress = rng.Address
    Do
        rng.Select
        DeleteSixRows
        Set rng = ws.Range("C:C").FindNext(rng)
    Loop While Not rng Is Nothing And rng.Address <> firstAddress

End Sub
Sub gapfinder()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim firstAddress As String
    Dim rngCollection As New Collection

    Set ws = ActiveSheet
    Set rng = ws.Range("A:A").Find("", LookIn:=xlValues, LookAt:=xlPart)
    If rng Is Nothing Then Exit Sub
    firstAddress = rng.Address
    rngCollection.Add rng
    Do
        Set rng = ws.Range("A:A").FindNext(rng)
        If rng Is Nothing Or rng.Address = firstAddress Then Exit Do
        rngCollection.Add rng
    Loop
    For i = rngCollection.Count To 1 Step -1
        rngCollection(i).EntireRow.Delete
    Next i
End Sub

Sub newcode1()
Columns("D:D").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Columns("D:D").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("D1").FormulaR1C1 = "TO"
Range("E1").FormulaR1C1 = "TRF"
Range("D2:D1042").FormulaR1C1 = "=IF(R[1]C[-2]="""",R[1]C[-1],"""")"
Range("E2:E1042").FormulaR1C1 = "=IF(R[2]C[-3]="""",IF(R[2]C[-2]=R[1]C[-1],R[1]C[-1],""""),"""")"
    Range("D2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("D7").Select
    Application.CutCopyMode = False
    Selection.ClearContents
End Sub
Sub CreateTable1()
    Dim rng As Range
    Dim tbl As ListObject
    Dim tblName As String
    Dim i As Integer
    i = 1
    tblName = "MyTable" & i
    While WorksheetExists(tblName)
        i = i + 1
        tblName = "MyTable" & i
    Wend

    ' Create a table on the specified range
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1:I474"), , xlYes)

    ' Assign the unique name to the table
    tbl.Name = tblName
End Sub

Function WorksheetExists(sName As String) As Boolean
    WorksheetExists = Evaluate("ISREF('" & sName & "'!A1)")
End Function




