Attribute VB_Name = "M_AutoFilter"
Option Explicit

Sub caller()

    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("Sheet2")
    
    Set_AutoFilter_Properties ws:=ws
    
    Set ws = Nothing
    Set wb = Nothing
    
End Sub

Public Sub Set_AutoFilter_Properties(ws As Worksheet)

    'This sub checks the worksheet to see if the worksheet has any listobjects
    'If there are listobjects, the sub will call the ListObject_AutoFilter Sub
    'If there are no listobjects, this sub will call the Worksheet_AutoFilter Sub.
    
    'Private Functions used :
    '---------------------------------
    '1.) ListObject_AutoFilter_ShowAll
    '2.) Worksheet_AutoFilter_ShowAll
    
    'Public Functions used:
    '---------------------------------
    '1.) GetLastUsedRow
    '2.) GetLastUsedColumn
    '3.) GetRangeFromCells
    '
    '===============================================================================
    '===============================================================================
    
    Dim lo As ListObject
    Dim r As Long
    Dim c As Long
    Dim rng As Range
    
    With ws
        If .ListObjects.Count > 0 Then
            For Each lo In .ListObjects
                ListObject_AutoFilter_ShowAll lo:=lo
            Next lo
        Else
            r = GetLastUsedRow(ws:=ws)
            c = GetLastUsedColumn(ws:=ws)
            Set rng = GetRangeFromCells(ws:=ws, _
                                        EndRow:=r, _
                                        EndCol:=c)
            Worksheet_AutoFilter_ShowAll ws:=ws
        End If
    End With
    
End Sub

Private Sub ListObject_AutoFilter_ShowAll(lo As ListObject)

    'Purpose    :   This sub works with listobjects (Tables) only.
    '               This sub will show all data if the listobject is filtered
    '               If the listobject is not filtered, this sub will turn on the filter
    'Parameters :
    '1.) lo     :   Required parameter. A listobject
    '
    '===================================================================================
    
    With lo
        If Not .AutoFilter Is Nothing Then
            .AutoFilter.ShowAllData
        Else
            .ShowAutoFilter = True
        End If
    End With
    
End Sub

Private Sub Worksheet_AutoFilter_ShowAll(rng As Range, _
                                         ws As Worksheet)

    'Purpose    :   This sub works with worksheets
    '               This sub will show all data if the worksheet is filtered
    '               If the worksheet is not filtered, this sub will turn on the filter
    'Parameters :
    '1.)    rng     :   Required parameter. A Range Object
    '2.)    ws      :   Required parameter. A Worksheet Object
    '
    '===================================================================================
                                
    With ws
        If Not .AutoFilter Is Nothing Then
            .AutoFilter.ShowAllData
        Else
            rng.AutoFilter
        End If
    End With
    
End Sub

Public Function GetLastUsedRow(ws As Worksheet) As Long

    'Purpose    :   A function to return the last used row on a worksheet
    'Parameters :
    '1)ws       :   Required parameter. A worksheet object
    '
    '=====================================================================
    
    GetLastUsedRow = ws.Range("A1").SpecialCells(xlCellTypeLastCell).Row
    
End Function

Public Function GetLastUsedColumn(ws As Worksheet) As Long

    'Purpose    :   A function to return the last used column on a worksheet
    'Parameters :
    '1)ws       :   Required parameter. A worksheet object
    '
    '=====================================================================
    
    GetLastUsedColumn = ws.Range("A1").SpecialCells(xlCellTypeLastCell).Column
    
End Function

Public Function GetRangeFromCells(ws As Worksheet, _
                                  EndRow As Long, _
                                  EndCol As Long, _
                                  Optional ByVal BeginRow As Long = 1, _
                                  Optional ByVal BeginCol As Long = 1) As Range
                                  
    'Purpose        :   A function to return a Range Object given 4 cells
    '
    'Parameters     :
    '----------------------------------------------------------------------------------------------------------------------------
    '
    '1.) EndRow     :   Required parameter. A value of Long Data Type. The last row in the range.
    '2.) EndCol     :   Required parameter. A value of Long Data Type. The last column in the range.
    '3.) BeginRow   :   Optional parameter. A value of Long Data Type. The first row in the range. Assumes range begins as A1.
    '4.) BeginCol   :   Optional parameter. A value of Long Data Type. The first column in the range. Assumes range begins at A1.
    '
    '=============================================================================================================================
    '=============================================================================================================================
    
    With ws
        Set GetRangeFromCells = .Range(.Cells(BeginRow, BeginCol), .Cells(EndRow, EndCol))
    End With

End Function

