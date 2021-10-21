Attribute VB_Name = "M_Fx_FilterDelete"
Option Explicit

Sub FilterOutUnwantedValues()

    'Objects
        Dim wb As Workbook
        Dim ws As Worksheet
        Dim wsCrit As Worksheet
        Dim rng As Range
        
    'Variables
        Dim x As Long
        Dim r As Long
        Dim rowNumber As Long
        Dim i As Long
        Dim y As Long
        Dim z As Long
        Dim acriteria() As Variant
        Dim acriteria1D() As String
    
    'Constants
        Const ndx As Long = 7
        
    'Initialize
        Set wb = ThisWorkbook
        Set ws = wb.Worksheets("Flat")
        Set wsCrit = wb.Worksheets("crit")
        Set rng = ws.Range("A4").CurrentRegion
        rowNumber = 1
        
    'Get count of used rows on worksheet
    'Assume column 1
        r = GetRows(ws:=wsCrit)
        
    'Now that we know the number of rows
    'Redimension the array
        ReDim acriteria(1 To r, 1)
        y = LBound(acriteria)
        z = UBound(acriteria)
        ReDim acriteria1D(y To z)
    
    'Populate array
    'Read in values from the worksheet
        For i = LBound(acriteria, 1) To UBound(acriteria, 1)
            acriteria(i, 1) = wsCrit.Cells(rowNumber, 1).Value
            rowNumber = rowNumber + 1
        Next i
        
    'Convert 2D array to 1D array
        For i = LBound(acriteria, 1) To UBound(acriteria, 1)
            acriteria1D(i) = acriteria(i, 1)
        Next i
        
    'Check contents of 1D Array
        For i = LBound(acriteria1D) To UBound(acriteria1D)
        Next i

        
    'Call the filter function
    'Pass the array as the filter criteria
        x = GetFilterDeleteRows(ws:=ws, _
                                rng:=rng, _
                                FilterCriteria:=acriteria1D, _
                                ColNumber:=ndx)
                                
    'Destroy objects
        Set rng = Nothing
        Set ws = Nothing
        Set wsCrit = Nothing
        Set wb = Nothing
                                
End Sub

Public Function GetFilterDeleteRows(ws As Worksheet, _
                                    rng As Range, _
                                    FilterCriteria As Variant, _
                                    ColNumber As Long) As Long
    Dim i As Long
    Dim rngDelete As Range
    
    rng.AutoFilter _
        Field:=ColNumber, _
        Criteria1:=FilterCriteria, _
        Operator:=xlFilterValues
        
    Set rngDelete = rng.Offset(1, 0) _
                       .Resize(rng.Rows.Count - 1, rng.Columns.Count) _
                       .SpecialCells(xlCellTypeVisible) _
                       .EntireRow
                       
    rngDelete.Delete
    
    ws.ShowAllData
    
    Set rngDelete = Nothing
    
    GetFilterDeleteRows = 0
                                         
End Function

