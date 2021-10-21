Attribute VB_Name = "M_Fx_FormatHeaderRow"
Option Explicit

Public Function FormatHeaderRow(ws As Worksheet) As Long

    'Objects
        Dim rng As Range
        
    'Variables
        Dim LastColumn As Long

    'Get last used column on sheet
        LastColumn = GetLastColumn(ws:=ws, _
                                   rowNumber:=4)
                                   
        Debug.Print "Header row last column : " & LastColumn
        
    'Create range object for header row based on A1: Last Column
        With ws
            Set rng = .Range(.Cells(4, 1), .Cells(4, LastColumn))
        End With
    
    'Freeze panes
        ws.Activate
        With ActiveWindow
            .SplitColumn = 0
            .SplitRow = 4
            .FreezePanes = True
        End With
        
    'Set color for header row
        rng.Interior.Color = RGB(255, 192, 0)
    
    'Turn on filter arrows
        rng.AutoFilter
        
    'Move cursor to A5
        ActiveSheet.Cells(5, 1).Select
        
    'Destroy objects
        Set rng = Nothing
        
    'Return function
        FormatHeaderRow = 0
    
End Function
