Attribute VB_Name = "M_Table_Sort"
Option Explicit

Sub Sort_Table()

    'Objects
        Dim wb As Workbook
        Dim ws As Worksheet
        Dim lo As ListObject
    
    'Initialize objects
        Set wb = ThisWorkbook
        Set ws = wb.Worksheets(1)
        Set lo = ws.ListObjects(1)
    
    'Pause Excel Environment
        With Application
            .Calculation = xlCalculationManual
            .ScreenUpdating = False
            .DisplayStatusBar = False
            .EnableEvents = False
            .DisplayAlerts = False
        End With
        
    '----------------
    'Sort Table Range
        lo.Range.Sort key1:=lo.ListColumns(1), _
                      order1:=xlAscending, _
                      Header:=xlYes, _
                      MatchCase:=False, _
                      Orientation:=xlTopToBottom
                      
    '----------------
    
    'Tidy up
        'Destroy objects
            Set lo = Nothing
            Set ws = Nothing
            Set wb = Nothing
            
        'Restore Excel Environment
            With Application
                .Calculation = xlCalculationAutomatic
                .ScreenUpdating = True
                .DisplayStatusBar = True
                .EnableEvents = True
                .DisplayAlerts = True
            End With
    


End Sub
