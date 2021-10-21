Attribute VB_Name = "M_Fx_ClearSheetPriorUse"
Option Explicit

Public Function ClearHeaderRow(ws As Worksheet) As Long

    'Clear freeze panes
           ws.Activate
           ActiveWindow.FreezePanes = False
        
    'Filtermode off
        If ws.AutoFilterMode Then
            ws.AutoFilterMode = False
        End If
        
    'Return function
        ClearHeaderRow = 0

End Function

Public Function ClearSheetPriorUse(ws As Worksheet) As Long

    On Error Resume Next

    'Clear freeze panes
        ws.Activate
        ActiveWindow.FreezePanes = False
        
    'Filtermode off
        If ws.AutoFilterMode Then
            ws.AutoFilterMode = False
        End If
        
    'Clear data and formats
        ws.UsedRange.Clear
        
    'Return to A1
         Application.Goto Reference:=ws.Range("A1"), _
                          Scroll:=True
         
    'Reset columnwidth
        ws.UsedRange.Columns.ColumnWidth = 8.43
         
    ClearSheetPriorUse = 0
End Function
