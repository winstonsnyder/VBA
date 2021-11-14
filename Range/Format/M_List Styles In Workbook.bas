Option Explicit

Sub GetListOfStylesInWorkbook()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    
    Dim i As Long
    
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets(1)
    Set rng = ws.Range("A1").CurrentRegion
    
    For i = 1 To wb.Styles.Count
        ws.Cells(i, 1) = wb.Styles(i).Name
    Next
    
    
    Set rng = Nothing
    Set ws = Nothing
    Set wb = Nothing
End Sub