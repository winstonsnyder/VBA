Attribute VB_Name = "M_Pivot_Table"
Private Sub Get_Pivot_Table(wb As Workbook)

    'Dependencies:
        '1.) GetPivotCache      :   Public Sub      :   'Public Function // Module : M_Public_Fx_PivotTables // GetPivotCache
        '2.) GetPivotTable      :   Public Sub      :   'Public Function // Module : M_Public_Fx_PivotTables // GetPivotTable
        '3.) AddFieldsToPivot   :   Private Sub     :   'Private Sub // Module : This Module // AddFieldsToPivot

    'Objects
        Dim ws As Worksheet
        Dim wsPivotTable As Worksheet
        Dim wsDataSource As Worksheet
        Dim rngDataSource As Range
        Dim my_pivotcache As PivotCache
        Dim my_pivot As PivotTable
        Dim pf As PivotField
    
    'Data source for Pivot Cache
        With wb
            For Each ws In .Worksheets
                If ws.Name Like "*crosstab*" Then
                    Set wsDataSource = wb.Worksheets(ws.Name)
                    Set rngDataSource = ws.Range("A1").CurrentRegion
                Else
                    MsgBox prompt:="The data source sheet for the Pivot Cache could not be found" & _
                                   "Exiting procedure.", _
                           Title:="Pivot Data Source Warning", _
                           Buttons:=vbOKOnly
                    Exit Sub
                End If
            Next ws
        End With
        
    'Add a new worksheet for the Pivot Table
        Set wsPivotTable = wb.Worksheets.Add(after:=wb.Worksheets(wb.Worksheets.Count))
        wsPivotTable.Name = "Pivot"
    
    'Get Pivot Cache
        'Public Function // Module : M_Public_Fx_PivotTables // GetPivotCache
            Set my_pivotcache = GetPivotCache(wb:=wb, _
                                              rng:=rngDataSource)
                                              
    'Where to place the Pivot Table on the worksheet
    
    
    'Get Pivot Table
        'Public Function // Module : M_Public_Fx_PivotTables // GetPivotCache
            Set my_pivot = GetPivotTable(pc:=my_pivotcache, _
                                         ws:=wsPivotTable, _
                                         strPivotTableName:="Master_Pivot", _
                                         lngRowPlacement:=4, _
                                         lngColPlacement:=1)
                                             
    'Add fields to Pivot Table
        'Private Sub // Module : This module // AddPivotTableFields
            AddFieldsToPivot pt:=my_pivot
        
    'Pivot table formatting
        With my_pivot
            .RowAxisLayout xlTabularRow
            .PivotFields("Sum of Commit (USD)").NumberFormat = "$#,##0.00"
            .RepeatAllLabels xlRepeatLabels
        End With
        
    'No subtotals
        With my_pivot
            For Each pf In .PivotFields
                pf.Subtotals(1) = False
            Next pf
        End With
            
    'Tidy up
        Set my_pivot = Nothing
        Set my_pivotcache = Nothing
        Set rngDataSource = Nothing
        Set wsDataSource = Nothing
        Set wsPivotTable = Nothing
        Set wb = Nothing

End Sub

Private Sub AddFieldsToPivot(pt As PivotTable)
  
    'Error handler
        On Error GoTo ErrHandler
  
    'Add fields to pivot table
        With pt
        
            'Page Fields (Filters)
                .PivotFields("Supplier Name").Orientation = xlPageField
                .PivotFields("Supplier Name").Position = 1
                
                .PivotFields("PO Number").Orientation = xlPageField
                .PivotFields("PO Number").Position = 2
                
            'Row fields
                .PivotFields("WBS Number").Orientation = xlRowField
                .PivotFields("WBS Number").Position = 1
                
                .PivotFields("GL").Orientation = xlRowField
                .PivotFields("GL").Position = 2
            
            'Column fields
                .PivotFields("CR Type").Orientation = xlColumnField
                .PivotFields("CR Type").Position = 1
            
            'Value fields
                .AddDataField .PivotFields("Commit (USD)"), _
                    Caption:="Sum of Commit (USD)", _
                    Function:=xlSum
        End With
  
ErrHandler:
    If Err.Number > 0 Then _
        MsgBox Err.Description, vbMsgBoxHelpButton, "Get pivot table fields", Err.HelpFile, Err.HelpContext
        Err.Clear
  
End Sub

Option Explicit

Public Function GetPivotCache(wb As Object, _
                              rng As Object) As PivotCache
  
    'Declare Objects
        Dim pc As Object
     
    'Declare variables
        Dim strPivotCacheSource As String
  
    'Error handler
        On Error GoTo ErrHandler
  
    'Pivot cache source
        strPivotCacheSource = rng.Parent.Name & "!" & _
                              rng.Address(ReferenceStyle:=xlR1C1)
  
    'Create pivot cache
        Set pc = wb.PivotCaches.Create( _
                        SourceType:=xlDatabase, _
                        SourceData:=strPivotCacheSource)
  
    'Pass object to function
        Set GetPivotCache = pc
  
ErrHandler:
    If Err.Number > 0 Then _
        MsgBox Err.Description, vbMsgBoxHelpButton, "Get pivot cache", Err.HelpFile, Err.HelpContext
  
    'Tidy up
        Set pc = Nothing
  
End Function

Public Function GetPivotTable(pc As Object, _
                              ws As Object, _
                              strPivotTableName As String, _
                              Optional ByVal lngRowPlacement As Long = 3, _
                              Optional ByVal lngColPlacement As Long = 3)
  
    'Declare Objects
        Dim pt As Object
        Dim rng As Object
  
    'Declare variables
        Dim strPivotPlacement As String
  
    'Error handler
        On Error GoTo ErrHandler
  
    'Create range
        Set rng = ws.Cells(lngRowPlacement, lngColPlacement)
  
    'Pivot table placement
        strPivotPlacement = ws.Name & "!" & _
                            rng.Address(ReferenceStyle:=xlR1C1)
  
    'Create pivot table
        Set pt = pc.CreatePivotTable( _
                    TableDestination:=strPivotPlacement, _
                    TableName:=strPivotTableName)
  
    'Pass object to function
        Set GetPivotTable = pt
  
ErrHandler:
    If Err.Number > 0 Then _
        MsgBox Err.Description, vbMsgBoxHelpButton, "Get pivot table", Err.HelpFile, Err.HelpContext
  
    'Tidy up
        Set rng = Nothing
        Set pt = Nothing
  
End Function


