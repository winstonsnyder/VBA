Attribute VB_Name = "M_Test"
Option Explicit

Sub DeleteCols()

    Dim wb As Workbook
    
    Set wb = ThisWorkbook
    
    With wb
        If Not .Worksheets("x") Is Nothing Then
            .Worksheets("x").Delete
        End If
    End With

    Set wb = Nothing
End Sub
