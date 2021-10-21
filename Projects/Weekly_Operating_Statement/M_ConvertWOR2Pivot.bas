Attribute VB_Name = "M_ConvertWOR2Pivot"
Option Explicit

Sub ConvertWOR2Pivot()
    
    'Declare objects
        Dim xlApp As Object
        Dim xlBook As Object
        Dim xlSheet As Object
        Dim xlRange As Object
        Dim xlLO As Object
        Dim wbWOR As Workbook
        Dim wsWOR As Worksheet
        Dim wbImport As Workbook
        Dim wsImport As Worksheet
        Dim GetRange As Range
        Dim GetUnPivotWorksheet As Worksheet
        
    'Declare variables
        Dim strFile As String
        Dim r As Long
        Dim C As Long
        Dim a As Long
        Dim rngRows As Long
        Dim rngCols As Long
        
    'Excel environment
        With Application
            .ScreenUpdating = False
            .DisplayAlerts = False
            .EnableEvents = False
            .Calculation = xlCalculationManual
        End With
        
    'Initialize objects
        Set wbWOR = ThisWorkbook
        Set wsWOR = wbWOR.Worksheets(1)
        
    'Get file for import
        strFile = GetFDObjectName(strDialogType:="File", _
                                  strTitle:="Select A File For Import")
    'Open the Excel workbook
        Set wbImport = Workbooks.Open(strFile)
        Set wsImport = wbImport.Worksheets(1)
        
    'Get range from workbook
        Set GetRange = GetExcelRangeFunc(ws:=wsImport, _
                                         lngRows:=7, _
                                         lngCols:=2)
                                         
    'Copy / paste range
        GetRange.Copy
        wsWOR.Range("A1").PasteSpecial xlPasteAll

    'Trim imported data
        Call TrimAll(rng:=GetRange)
        
    'Close workbooks / Destroy objects
        wbImport.Close
        Set GetRange = Nothing
        Set wsImport = Nothing
        Set wbImport = Nothing
        
    'Get cost center
        Call GetCostCenter(ws:=wsWOR)
        
    'Unstack account labels
        Call UnstackAccountLabels(ws:=wsWOR)

    'Delete rows
        Call DeleteRows(ws:=wsWOR)
        
    'Fill Labels
        Call FillLabels(ws:=wsWOR)
        
    'Delete columns pcnt
        Call DeleteColumnsPcnt(ws:=wsWOR)
        
    'Delete blank rows
        Call DeleteBlankRows(ws:=wsWOR)
        
    'Lookup account labels
        Call LookupAccountLabels(ws:=wsWOR)
        
    'Get parent account
        Call GetParentAccount(ws:=wsWOR)
        
    'Delete rows - week and period
        Call DeleteSpecificLabels(ws:=wsWOR)
        
    'Add headers
        Call AddHeaders(ws:=wsWOR)
        
    'UnPivot data
        Set GetUnPivotWorksheet = UnPivotData(wb:=wbWOR, _
                                              ws:=wsWOR)
                                              
    'Add cost center name
        Call AddCostCenterName(ws:=GetUnPivotWorksheet)
        
    'Create new Excel app
        Set xlApp = GetApplication(strApplication:="Excel")
        xlApp.Visible = True
        
    'Add workbook to new instance
        Set xlBook = xlApp.Workbooks.Add
        
    'Create worksheet object in new instance
        Set xlSheet = xlBook.Worksheets(1)
        
    'Get current region
        Set GetRange = Nothing
        Set GetRange = GetCurrentRegion(ws:=GetUnPivotWorksheet)
        
    'Get rows and columns of region
        With GetRange
            rngRows = .Rows.Count
            rngCols = .Columns.Count
        End With
        
    'Resize destination range
        With xlSheet
            Set xlRange = .Range("A1")
            Set xlRange = xlRange.Resize(rngRows, rngCols)
        End With
        
    'Transfer range values
        xlRange.Value = GetRange.Value
        
    'Add a listobject
        Call AddListObject(ws:=xlSheet, _
                           rng:=xlRange)
                                     
    'Tidy up
        'Destroy objects
            Set xlRange = Nothing
            Set GetRange = Nothing
            Set xlSheet = Nothing
            Set wsWOR = Nothing
            Set wsImport = Nothing
            Set GetUnPivotWorksheet = Nothing
            Set xlBook = Nothing
            Set wbImport = Nothing
            Set wbWOR = Nothing
            Set xlApp = Nothing
            
        'Excel environment
            With Application
                .ScreenUpdating = True
                .DisplayAlerts = True
                .EnableEvents = True
                .Calculation = xlCalculationAutomatic
            End With

End Sub

Public Function GetExcelRangeFunc(ws As Worksheet, _
                                  lngRows As Long, _
                                  lngCols As Long) As Range
    'Declare objects
        Dim rng As Range
        
    'Declare variables
        Dim lngMeRows As Long
        Dim lngMeCols As Long
        
    'Get last column - use row 7
        lngMeCols = GetLast(ws:=ws, _
                            RC:="c", _
                            lngRowColumn:=lngRows)
                                
    'Get last row - use col 2
        lngMeRows = GetLast(ws:=ws, _
                            RC:="r", _
                            lngRowColumn:=lngCols)
                                
    'Create range
        With ws
            Set rng = .Range(.Cells(1, 1), .Cells(lngMeRows, lngMeCols))
        End With
        
    'Pass object to function
        Set GetExcelRangeFunc = rng
        
    'Tidy up
        Set rng = Nothing
   
End Function

Private Sub TrimAll(rng As Range)

    Dim C As Range

'    'Also Treat CHR 0160, as a space (CHR 032)
'        rng.Select
'        Selection.Replace what:=Chr(160), Replacement:=Chr(32), _
'        LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False


    If Not rng Is Nothing Then
        
        'Try:

        'Also Treat CHR 0160, as a space (CHR 032)
            rng.Replace _
                what:=Chr(160), _
                Replacement:=Chr(32), _
                LookAt:=xlPart, _
                SearchOrder:=xlByRows, _
                MatchCase:=False
    
            'Trim in Excel removes extra internal spaces, VBA does not
        '        On Error Resume Next   'in case no text cells in selection
        '        For Each cell In Intersect(Selection, _
        '            Selection.SpecialCells(xlConstants, xlTextValues))
        '            cell.Value = Application.Trim(cell.Value)
        '        Next cell
        '        On Error GoTo 0
        '        rng.Cells(1, 1).Select
    
            For Each C In rng
                C.Value = Trim(C.Value)
            Next C
            
    Else
        'Catch
            MsgBox "The Range does not exist"
    End If
                
End Sub


Private Sub GetCostCenter(ws As Worksheet)

    'Declare objects
        Dim C As Range
    
    'Declare variables
        Dim lngMeRows As Long
        Dim lngMeValue As Long
        
    'Declare constants
        Const strFORMULA = "=IFERROR(VALUE(TRIM(MID(B1,FIND("":"",B1)+2,8))),0)"
        
    'Initialize variables
        lngMeValue = 0
    
    'Get max rows
        lngMeRows = GetLast(ws:=ws, _
                            RC:="r", _
                            lngRowColumn:=2)

    'Get cost center numbers
        With ws
            'Get cost center formula
                .Range("A1").EntireColumn.Insert
                .Range("A1:A" & lngMeRows).Formula = strFORMULA
            
            'Flatten formulas to values
                For Each C In .Range("A1:A" & lngMeRows)
                    C.Value = C.Value
                Next C
            
            'Fill cost center down
                For Each C In .Range("A1:A" & lngMeRows)
                    If C.Value = 0 Then
                        C.Value = lngMeValue
                    Else
                        lngMeValue = C.Value
                    End If
                Next C
        End With
        
    'Set Column Width
        ws.Columns("A").ColumnWidth = 12
     
End Sub

Private Sub DeleteRows(ws As Worksheet)

    'Declare variables
        Dim lngMeRows As Long
        Dim i As Long
                                
    'Get last row - use col 2
        lngMeRows = GetLast(ws:=ws, _
                            RC:="r")
                            
    'Delete rows
        With ws
            For i = lngMeRows To 1 Step -1
                If .Cells(i, 7).Value = "SODEXO" Or _
                    .Cells(i, 7).Value = "Weekly Operating Report" Or _
                    .Cells(i, 7).Value = "Cost Center Detail" Or _
                    InStr(.Cells(i, 12), "Period") Or _
                    InStr(.Cells(i, 12), "W/E") Or _
                    InStr(.Cells(i, 15), "Page") Or _
                    .Cells(i, 3).Value = "Account" Or _
                    .Cells(i, 2).Value = "COSTS" Or _
                    .Cells(i, 2).Value = "AMORT AND IMPAIRMENT" Or _
                    .Cells(i, 2).Value = "DIRECT COSTS" Or _
                    .Cells(i, 2).Value = "PROCESSING COSTS" Or _
                    .Cells(i, 2).Value = "CONTRIBUTION" Or _
                    .Cells(i, 2).Value = "PROFIT" Or _
                    .Cells(i, 2).Value = "PERSONNEL COSTS" Or _
                    .Cells(i, 3).Value = "Number" Then
                    .Cells(i, 1).EntireRow.Delete shift:=xlUp
                End If
            Next i
        End With

    'Get last row
        lngMeRows = GetLast(ws:=ws, _
                            RC:="r")
                            
    'Delete rows - 2nd pass
        With ws
            For i = lngMeRows To 1 Step -1
                If IsEmpty(.Cells(i, 2)) And _
                    IsEmpty(.Cells(i, 3)) Then
                    .Cells(i, 1).EntireRow.Delete shift:=xlUp
                End If
            Next i
        End With
        
End Sub

Private Sub UnstackAccountLabels(ws As Worksheet)

    'Declare variables
        Dim lngMeRows As Long
        Dim i As Long
        Dim strValue As String
                                
    'Get last row - use col 2
        lngMeRows = GetLast(ws:=ws, _
                            RC:="r")
                            
    'Delete rows
        With ws
            For i = lngMeRows To 1 Step -1
                If .Cells(i, 2).Value = "PER0000 - OPERATING" Or _
                    .Cells(i, 2).Value = "DAI0000 - OPERATING DEPR" Or _
                    .Cells(i, 2).Value = "TPC9999 - OPERATING" Or _
                    .Cells(i, 2).Value = "EBI9000 - UNIT OPERATING" Or _
                    .Cells(i, 2).Value = "FCO9999 - OPERATING" Then
                    .Cells(i, 2).Value = .Cells(i, 2).Value & " " & .Cells(i, 2).Offset(1, 0).Value
                ElseIf InStr(.Cells(i, 2), "ODC0000") Then
                  .Cells(i, 2).Value = Left(.Cells(i, 2).Value, 25) & " " & .Cells(i, 2).Offset(1, 0).Value
                  .Cells(i, 2).Offset(0, 1).Value = "Week"
                ElseIf InStr(.Cells(i, 2), "GRP9999 - OPERATING GROSS") Then
                  .Cells(i, 2).Value = Left(.Cells(i, 2).Value, 25) & " " & .Cells(i, 2).Offset(1, 0).Value
                  .Cells(i, 2).Offset(0, 1).Value = "Week"
                ElseIf InStr(.Cells(i, 2), "64501101") Then
                  strValue = .Cells(i, 2).Value
                  .Cells(i, 2).Value = Left(strValue, 27)
                  .Cells(i, 2).Offset(0, 1).Value = Right(strValue, 8)
                End If
            Next i
        End With
End Sub

Private Sub FillLabels(ws As Worksheet)

    'Declare variables
        Dim lngMeRows As Long
        Dim i As Long
        Dim x As Long
        Dim varValue As Variant
                                
    'Get last row - use col 2
        lngMeRows = GetLast(ws:=ws, _
                            RC:="r")
                            
        x = lngMeRows
        varValue = vbNullString
                            
    'Fill values down
        With ws
            For i = 1 To lngMeRows
                If IsEmpty(.Cells(i, 2)) Then
                    .Cells(i, 2).Value = varValue
                Else
                    varValue = .Cells(i, 2).Value
                End If
            Next i
            
            varValue = vbNullString
            
            For i = 1 To x
                If IsEmpty(.Cells(i, 3)) Then
                    .Cells(i, 3).Value = varValue
                Else
                    varValue = .Cells(i, 3).Value
                End If
            Next i
        End With
End Sub

Private Sub DeleteColumnsPcnt(ws As Worksheet)

    'Declare objects
        Dim rng As Range
    
    'Create range object
        With ws
            Set rng = Union(.Range("E1"), _
                            .Range("G1"), _
                            .Range("I1"), _
                            .Range("K1"), _
                            .Range("M1"), _
                            .Range("O1"))
        End With
        
    'Delete columns in range
        rng.Columns.EntireColumn.Delete shift:=xlLeft
    
    'Tidy up
        Set rng = Nothing
        
End Sub

Private Sub DeleteBlankRows(ws As Worksheet)

    'Declare variables
        Dim lngMeRows As Long
        Dim i As Long
                                
    'Get last row - use col 2
        lngMeRows = GetLast(ws:=ws, _
                            RC:="r")
                            
    'Delete blank rows
        With ws
            For i = lngMeRows To 1 Step -1
                If IsEmpty(.Cells(i, 4)) Then
                    .Cells(i, 1).EntireRow.Delete
                End If
            Next i
        End With
        
End Sub

Private Sub LookupAccountLabels(ws As Worksheet)

    'Declare objects
        Dim C As Range
    
    'Declare variables
        Dim lngMeRows As Long
        Dim lngMeValue As Long
        
    'Declare constants
        Const strFORMULA = "=IFERROR(IF(OR(C1=""Week"",C1=""Period""),B1,VLOOKUP(C1,tblAccounts,2,FALSE)),0)"
        
    'Initialize variables
        lngMeValue = 0
    
    'Get max rows
        lngMeRows = GetLast(ws:=ws, _
                            RC:="r", _
                            lngRowColumn:=2)

    'Get cost center numbers
        With ws
            'Get cost center formula
                .Range("D1").EntireColumn.Insert
                .Range("D1:D" & lngMeRows).Formula = strFORMULA
            
            'Flatten formulas to values
                For Each C In .Range("D1:D" & lngMeRows)
                    C.Value = C.Value
                Next C
        End With
        
    'Set Column Width
        ws.Columns("D").ColumnWidth = 12
End Sub

Private Sub GetParentAccount(ws As Worksheet)

    'Declare objects
        Dim C As Range
    
    'Declare variables
        Dim lngMeRows As Long
        Dim lngMeValue As Long
        Dim varValue As Variant
        Dim i As Long
        
    'Declare constants
        Const strFORMULA = "=IF(ISNUMBER(C1),0,D1)"
        
    'Initialize variables
        lngMeValue = 0
        varValue = vbNullString
    
    'Get max rows
        lngMeRows = GetLast(ws:=ws, _
                            RC:="r", _
                            lngRowColumn:=2)

    'Get cost center numbers
        With ws
            'Get cost center formula
                .Range("E1").EntireColumn.Insert
                .Range("E1:E" & lngMeRows).Formula = strFORMULA
            
            'Flatten formulas to values
                For Each C In .Range("E1:E" & lngMeRows)
                    C.Value = C.Value
                Next C
        End With
        
    'Fill values up
        With ws
            For i = lngMeRows To 1 Step -1
                If .Cells(i, 5).Value = 0 Then
                    .Cells(i, 5).Value = varValue
                Else
                    varValue = .Cells(i, 5).Value
                End If
            Next i
        End With
        
    'Set Column Width
        ws.Columns("E").ColumnWidth = 12

End Sub


Private Sub DeleteSpecificLabels(ws As Worksheet)

    'Declare variables
        Dim lngMeRows As Long
        Dim i As Long
                                
    'Get last row - use col 2
        lngMeRows = GetLast(ws:=ws, _
                            RC:="r")
                            
    'Delete rows
        With ws
            For i = lngMeRows To 1 Step -1
                If .Cells(i, 3).Value = "Period" Or _
                    .Cells(i, 3).Value = "Week" Then
                    .Cells(i, 1).EntireRow.Delete shift:=xlUp
                End If
            Next i
        End With

    'Get last row
        lngMeRows = GetLast(ws:=ws, _
                            RC:="r")
                            
    'Delete rows - 2nd pass
        With ws
            For i = lngMeRows To 1 Step -1
                If IsEmpty(.Cells(i, 2)) And _
                    IsEmpty(.Cells(i, 3)) Then
                    .Cells(i, 1).EntireRow.Delete shift:=xlUp
                End If
            Next i
        End With
        
End Sub

Private Sub AddHeaders(ws As Worksheet)

    With ws
        .Range("A1").EntireRow.Insert shift:=xlDown
        .Range("A1").Value = "CostCenter"
        .Range("B1").Value = "AccountElement"
        .Range("C1").Value = "AccountNumber"
        .Range("D1").Value = "AccountDescription"
        .Range("E1").Value = "ParentAccount"
        .Range("F1").Value = "Week1"
        .Range("G1").Value = "Week2"
        .Range("H1").Value = "Week3"
        .Range("I1").Value = "Week4"
        .Range("J1").Value = "Week5"
        .Range("K1").Value = "PeriodTotal"
    End With
    
End Sub
Private Function UnPivotData(wb As Workbook, _
                             ws As Worksheet) As Worksheet

    'Declare objects
        Dim wsUnPivot As Worksheet
        
    'Declare variables
        Dim MaxRows As Long
        Dim MaxColumns As Long
        Dim i As Long   'rows to unpivot
        Dim j As Long   'columns to unpivot
        Dim k As Long   'output rows
        
    'Declare constants
        Const strSheetName As String = "UnPivot"
        
    'Initialize variables
        k = 2
        
    'Add worksheet for UnPivot output
        Set wsUnPivot = AddWorksheet(wb:=wb, _
                                     strSheetName:=strSheetName)
                                     
    'Rows for UnPivot
        MaxRows = GetLast(ws:=ws, _
                          RC:="r")
                          
    'Columns for UnPivot
        MaxColumns = GetLast(ws:=ws, _
                             RC:="c")
                             
    'UnPivot
        With wsUnPivot
            For i = 2 To MaxRows 'Begin with first row of (Measures)
                For j = 6 To MaxColumns 'Begin with first column of data (Measures)
                    .Cells(k, 1).Value = ws.Cells(i, 1).Value  'Cost Center
                    .Cells(k, 2).Value = ws.Cells(i, 2).Value  'Account Element
                    .Cells(k, 3).Value = ws.Cells(i, 3).Value  'Account Number
                    .Cells(k, 4).Value = ws.Cells(i, 4).Value  'Account Description
                    .Cells(k, 5).Value = ws.Cells(i, 5).Value  'Account Parent
                    .Cells(k, 6).Value = ws.Cells(1, j).Value  'Time
                    .Cells(k, 7).Value = ws.Cells(i, j).Value  'Value
                    k = k + 1
                Next j
            Next i
        End With
        
    'Add Cost Center Name
        
    'Add headers
        With wsUnPivot
            .Range("A1").Value = "CostCenter"
            .Range("B1").Value = "AccountElement"
            .Range("C1").Value = "AccountNumber"
            .Range("D1").Value = "AccountDescription"
            .Range("E1").Value = "AccountParent"
            .Range("F1").Value = "Time"
            .Range("G1").Value = "Amount"
        End With
        
    'Pass object to function
        Set UnPivotData = wsUnPivot
        
    'Tidy up
        Set wsUnPivot = Nothing

End Function

Private Sub AddCostCenterName(ws As Worksheet)

    'Declare objects
        Dim C As Range
    
    'Declare variables
        Dim lngMeRows As Long
        Dim lngMeValue As Long
        
    'Declare constants
        Const strFORMULA = "=IFERROR(VLOOKUP(A2,tblUnits,2,FALSE),0)"
        
    'Initialize variables
        lngMeValue = 0
    
    'Get max rows
        lngMeRows = GetLast(ws:=ws, _
                            RC:="r")

    'Get cost center numbers
        With ws
            'Get cost center formula
                .Range("D1").EntireColumn.Insert
                .Range("D2:D" & lngMeRows).Formula = strFORMULA
            
            'Flatten formulas to values
                For Each C In .Range("D1:D" & lngMeRows)
                    C.Value = C.Value
                Next C
        End With
        
    'Add header
        ws.Range("D1").Value = "CostCenterName"
        
    'Set Column Width
        ws.Columns("D").ColumnWidth = 12
End Sub


Private Function GetCurrentRegion(ws As Worksheet) As Range

    'Declare variables
        Dim rng As Range
        
    'Get current region
        Set rng = ws.Range("A1").CurrentRegion
        
    'Pass object to function
        Set GetCurrentRegion = rng
        
    'Tidy up
        Set rng = Nothing
        
End Function

Private Sub AddListObject(ws As Worksheet, _
                          rng As Range)
                          
    'Declare objects
        Dim lo As ListObject
        
    'Add listobject
        Set lo = ws.ListObjects.Add( _
                                    SourceType:=xlSrcRange, _
                                    Source:=rng, _
                                    Destination:=ws.Range("A1"))

End Sub

