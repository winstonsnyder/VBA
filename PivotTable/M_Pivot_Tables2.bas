Attribute VB_Name = "M_Pivot_Tables2"
 Sub GetPivotTable()
 

 
 'Get Pivot Cache
        'Public Function // Module : M_Public_Fx_PivotTables // GetPivotCache
            Set my_pivotcache = GetPivotCache(wb:=wb, _
                                              rng:=rngDataSource)
    
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
        
    'No Subtotals/Grnad Totals
        With my_pivot
            For Each pf In .PivotFields
                pf.Subtotals(1) = False
            Next pf
            .ColumnGrand = False
            .RowGrand = False
        End With
                                         
End Sub
                                         

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


