Attribute VB_Name = "M_Array_2D_1D"
Option Explicit
Option Base 1

Sub Load_Array_2D_1D()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    Dim arr As Variant
    Dim animal() As String
    Dim quant() As Long
    Dim i As Long
    Dim j As Long
    Dim x
    
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("Sheet1")
    Set rng = ws.Range("A1").CurrentRegion
    arr = rng.Value
    x = 1
    
    ReDim animal(UBound(arr, 1), 1)
    ReDim quant(UBound(arr, 2), 1)
    
    For i = LBound(arr, 1) To UBound(arr, 1)
        animal(x, 1) = CStr(arr(x, 1))
        quant(x, 1) = CLng(arr(x, 2))
        x = x + 1
    Next i
    
    For i = LBound(animal) To UBound(animal)
        Debug.Print "animal: ", animal(i, 1)
        Debug.Print "quant: ", quant(i, 1)
    Next i
    
    ReDim gv_animal(UBound(animal, 1), 1)
    ReDim gv_quant(UBound(quant, 1), 1)
    
    gv_animal = animal
    gv_quant = quant
    
    Debug.Print "Globals"
    Debug.Print "================================"
    For i = LBound(gv_animal) To UBound(gv_animal)
        Debug.Print "animal: ", gv_animal(i, 1)
        Debug.Print "quant: ", gv_quant(i, 1)
    Next i
    
    
    Set rng = Nothing
    Set ws = Nothing
    Set wb = Nothing
End Sub
