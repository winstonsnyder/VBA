Attribute VB_Name = "M_Fx"
Option Explicit

Public Function DeleteRange(ws As Worksheet, _
                            FirstRow As Long, _
                            FirstColumn As Long, _
                            LastColumn As Long) As Long
                            
    Dim LastRow As Long
    Dim rng As Range
    
    LastRow = GetRows(ws:=ws)
    
    With ws
        Set rng = .Range(.Cells(FirstRow, FirstColumn), .Cells(LastRow, LastColumn))
    End With
    
    rng.Clear
    
    Set rng = Nothing
    
    DeleteRange = 0

End Function

Public Function GetTrimRangeA1(ws As Worksheet) As Long

    Dim rng As Range
    Dim cell As Range
    
    Set rng = ws.Range("A1").CurrentRegion
    
    For Each cell In rng
        cell.Value = Trim(cell.Value)
    Next cell
    
    Set rng = Nothing
    
End Function

Public Function CopyRangeToWorksheet(wsSource As Worksheet, _
                                     wsDestination As Worksheet, _
                                     ToRow As Long, _
                                     ToColumn) As Long
                                      
    'Objects
        Dim rngSource As Range
        Dim rngDestination As Range
    
    'Variables
        Dim FirstRowData As Long
        Dim LastRowData As Long
        
    'Get last row from source sheet
        FirstRowData = GetFirstNonEmptyCellA1(ws:=wsSource)
        LastRowData = GetRows(ws:=wsSource)
        
    'Create range object for source data
        With wsSource
            Set rngSource = .Range(.Cells(FirstRowData, 1), .Cells(LastRowData, 1))
        End With
        
    'Create range object for destination
        With wsDestination
            Set rngDestination = .Range(.Cells(ToRow, ToColumn), .Cells(ToRow, ToColumn))
        End With

    'Copy the source range to the destination range
        rngSource.Copy
        rngDestination.PasteSpecial xlPasteValuesAndNumberFormats
        rngDestination.PasteSpecial xlPasteFormats
        
    'Destroy objects
        Set rngSource = Nothing
        Set rngDestination = Nothing
        
    'Update function
        CopyRangeToWorksheet = 0

End Function


Public Function CopyRangeToWorksheetOld(wsSource As Worksheet, _
                                        wsDestination As Worksheet, _
                                        lngDestinationRow As Long) As Long
                                      
    'Objects
        Dim rngSource As Range
        Dim rngDestination As Range
    
    'Variables
        Dim FirstRowData As Long
        Dim LastRowData As Long
        Dim r As Long
        Dim c As Long
        
    'Get last row from source sheet
        FirstRowData = GetFirstNonEmptyCellA1(ws:=wsSource)
        LastRowData = GetRows(ws:=wsSource)
        
    'Create range object for source data
        With wsSource
            Set rngSource = .Range(.Cells(FirstRowData, 1), .Cells(LastRowData, 1))
        End With
        
    'Get dimensions of source range
        With rngSource
            r = .Rows.Count
            c = .Columns.Count
        End With
        
    'Create a destination range
        With wsDestination
            Set rngDestination = .Cells(lngDestinationRow, 1)
        End With
        
    'Resie the destination range so same size as source range
        Set rngDestination = rngDestination.Resize(r, c)
        
    'Copy the source range to the destination range
        rngDestination.Value = rngSource.Value
        
    'Destroy objects
        Set rngSource = Nothing
        Set rngDestination = Nothing
        
    'Update function
        CopyRangeToWorksheet = 0

End Function

Public Function GetPeriodMonth(ControlName As String) As String

    Dim z As String

    Select Case ControlName
        Case "optSeptember"
            z = "Period 01"
        Case "optOctober"
            z = "Period 02"
        Case "optNovember"
            z = "Period 03"
        Case "optDecember"
            z = "Period 04"
        Case "optJanuary"
            z = "Period 05"
        Case "optFebruary"
            z = "Period 06"
        Case "optMarch"
            z = "Period 07"
        Case "optApril"
            z = "Period 08"
        Case "optMay"
            z = "Period 09"
        Case "optJune"
            z = "Period 10"
        Case "optJuly"
            z = "Period 11"
        Case "optAugust"
            z = "Period 12"
    End Select
    
    GetPeriodMonth = z

End Function


Public Function GetRows(ws As Worksheet, _
                        Optional ByVal ColNumber As Long = 1) As Long
 
    'ws             :   Worksheet
    'ColNumber      :   Column number to be used to determine the last row, default is Column 1 (A)
    'Output         :   A row number of type long
     
     

    With ws
        GetRows = .Cells(Rows.Count, ColNumber).End(xlUp).Row
    End With

         
End Function

Public Function GetLastColumn(ws As Worksheet, _
                              Optional ByVal rowNumber As Long = 1) As Long
 
    'ws         :   A Worksheet Object
    'Output     :   A column number of type long
    'Row        :   Optional, Assume row 1
    

     GetLastColumn = ws.Cells(rowNumber, Columns.Count).End(xlToLeft).Column

         
End Function

Public Function GetFirstNonEmptyCellA1(ws As Worksheet) As Long

    Dim i As Long
    Dim rng As Range
    Dim cell As Range
    
    Set rng = ws.Range("A1").CurrentRegion
    i = 1
    
    For Each cell In rng
        If Not (IsEmpty(cell)) Then
            GetFirstNonEmptyCellA1 = i
            Exit For
        End If
        i = i + 1
    Next cell
    
    Set rng = Nothing
    
End Function

