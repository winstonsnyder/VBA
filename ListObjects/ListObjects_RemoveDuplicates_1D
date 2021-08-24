
Option Explicit

Sub ListObjects_RemoveDuplicates_1D()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets(1)
    Set lo = ws.ListObjects(1)
    
    lo.Range.RemoveDuplicates Columns:=Array(1), Header:=xlYes
    
    Set lo = Nothing
    Set ws = Nothing
    Set wb = Nothing
End Sub
