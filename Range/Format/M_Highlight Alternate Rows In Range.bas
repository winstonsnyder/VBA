Option Explicit


Sub HighlightAlternateRowsInRange()

Dim wb As Workbook
Dim ws As Worksheet
Dim rng As Range
Dim RangeForFill As Range
Dim i As Long

Set wb = ThisWorkbook
Set ws = wb.Worksheets(1)
Set rng = ws.Range("A1").CurrentRegion
Set rng = rng.Offset(1).Resize(rng.Rows.Count - 1)

'Set row style -- bottom up
For i = rng.Rows.Count To 1 Step -2
    Set RangeForFill = rng.Rows(i)
    RangeForFill.Style = "20% - Accent3"
    Set RangeForFill = Nothing
Next i

Set RangeForFill = Nothing
Set rng = Nothing
Set ws = Nothing
Set wb = Nothing
End Sub
