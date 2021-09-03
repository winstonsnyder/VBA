Option Explicit

Sub RangeToArray_1D()

    Dim wb As Workbook
    Dim lo As ListObject
    Dim arr() As Variant
    Dim r As Long               'Last row
    Dim rng As Range
    Dim i As Long
    
    Set wb = ThisWorkbook
    Set lo = Data.ListObjects(1)
    r = GetLastRow(ws:=Crit)
    
    With Crit
        Set rng = .Range(.Cells(2, 1), .Cells(r, 1))
    End With
    
    arr = rng
    
    For i = LBound(arr) To UBound(arr)
        Debug.Print i, arr(i, 1)
    Next i
    

    Erase arr
    Set rng = Nothing
    Set lo = Nothing
    Set wb = Nothing
End Sub
