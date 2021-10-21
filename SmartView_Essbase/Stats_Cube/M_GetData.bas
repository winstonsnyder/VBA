Attribute VB_Name = "M_GetData"
Option Explicit

Sub GetData()

    'Declare objects
        Dim wb As Workbook
        Dim wsRtrv As Worksheet
        Dim wsUnits As Worksheet
        Dim wsAdmin As Worksheet
        Dim wsInputs As Worksheet
        Dim wsFinal As Worksheet
        Dim wsAccounts As Worksheet
        Dim rngAccountLabels As range
        Dim rngFilter As range
        Dim rngReplace As range
        Dim rngCopy As range 'Value for data type coercion
        Dim rngForCopy As range
        Dim rngOrgName As range
        Dim rngTimeDimension As range
        Dim rngFunctionDimension As range
        Dim C As range
        
        Dim wbNew As Workbook
        Dim ws As Worksheet
        Dim wsConnect As Worksheet
        Dim wsDisconnect As Worksheet

    
    'Declare variables
        Dim lngRowsUnits As Long
        Dim lngRowsRetrieve As Long
        Dim lngRowsFinal As Long
        Dim lngRowsOrgName As Long
        Dim lngLastRow As Long
        Dim lngRowsChk As Long
        Dim i As Long
        Dim y As Long   'EssConnect
        Dim z As Long   'EssZoomIn
        Dim w As Long   'EssDisconnect
        Dim blnFlag As Boolean
        Dim blnFlagOrg As Boolean
        
        Dim lngSheetsCount As Long
        Const lngACCOUNTNUMBER As Long = 1
        Const strPath As String = "C:\Delete\"
        Const strFileName As String = "Davis_Stats.xlsx"
        Dim blnFolderExists As Boolean
        Dim blnFileExists As Boolean
        i = 1

    'Excel environment
        With Application
            .ScreenUpdating = False
            .EnableEvents = False
            .Calculation = xlCalculationManual
        End With
        
    'Initialize objects
        Set wb = ThisWorkbook
        Set wsAccounts = wb.Worksheets("infAccounts")
        Set wsRtrv = wb.Worksheets("Rtrv")
        Set wsUnits = wb.Worksheets("infUnits")
        Set wsAdmin = wb.Worksheets("infAdmin")
        Set wsInputs = wb.Worksheets("infInputs")
        Set wsFinal = wb.Worksheets("dataFinal")
        Set rngCopy = wsInputs.range("B3")

    'Initialize variables
        lngRowsUnits = GetLast(ws:=wsUnits, _
                               RC:="r")
        blnFlag = True
        blnFlagOrg = True
        i = 1
        
    'Check if directory for working files exists, if not, create the directory
        blnFolderExists = CreateDirectory(strPath:=strPath)
        
    'Delete file if it exists - previous use
        blnFileExists = DeleteFiles(strPath:=strPath)
        
    'Create workbook and add worksheets for retrieve------------------------------------------------------------------
        'Add a workbook
            Set wbNew = Workbooks.Add
            
        'Save the new workbook
            wbNew.SaveAs strPath & strFileName, _
                FileFormat:=51
            
        'Count number of units to determine number of sheets for retrieve
            lngRowsUnits = GetLast(ws:=wsUnits, _
                                   RC:="r")
                                   
        'Count number of sheet in workbook
            lngSheetsCount = wbNew.Sheets.Count
            
        'Add additional sheets to the workbook to equal total number of cost centers
            With wbNew
                .Worksheets.Add After:=.Worksheets(.Worksheets.Count), _
                    Count:=lngRowsUnits - lngSheetsCount
            End With
            
        'Rename sheets
            With wbNew
                For Each ws In .Worksheets
                    ws.Name = CStr(ws.Index)
                    ws.Tab.ColorIndex = 3
                Next ws
            End With
            
    'Setup sheets for retrieve----------------------------------------------------------------------------------------
        With wbNew
            For Each ws In .Worksheets
                With ws
                    .range("A1").Value = wsRtrv.range("A1").Value               'Scenario
                    .range("A2").Value = wsRtrv.range("A2").Value               'Currency
                    .range("A3").Value = wsRtrv.range("A3").Value               'Document Type
                    .range("A4").Value = wsUnits.Cells(i, 1).Value              'Organization
                    .range("B5").Value = wsRtrv.range("B5").Value               'Time
                    .range("A6").Value = wsRtrv.range("A6").Value               'Account
                End With
                i = i + 1
            Next ws
        End With
        
    'Retrieve/ZoomIn/Disconnect each sheet-------------------------------------------------------------------------------------
    
        'Add a worksheet for Essbase connection
            With wbNew
                Set wsConnect = .Worksheets.Add(After:=.Sheets(.Sheets.Count))
                wsConnect.Name = "EssConnect"
            End With
                     
        'Essbase Connect
            y = GetConnected(wsEssConnectionValues:=wsAdmin, _
                             wsEssConnect:=wsConnect)
        
        'Retrieve/ZoomIn each worksheet
            With wbNew
                For Each ws In .Worksheets
                    If ws.Tab.ColorIndex = 3 Then
                        'Essbase ZoomIn
                        'Level 3 is bottom level
                            z = GetZoomData(sheetName:=ws.Name, _
                                            range:=Null, _
                                            selection:=wsRtrv.Cells(6, 1), _
                                            level:=3, _
                                            across:=False)
                    End If
                Next ws
            End With
            
        'Disconnect Essbase
            With wbNew
                For Each ws In .Worksheets
                    w = GetDisconnected(ws:=ws)
                Next ws
            End With
                
    'Replace, filter, copy, paste -----------------------------------------------------------------------------------------------
    
        With wbNew
            For Each ws In .Worksheets
                If ws.Tab.ColorIndex = 3 Then
                
                    'Get the last row on the sheet
                        lngLastRow = GetLast(ws:=ws, _
                                             RC:="r")
                                      
                    'Create a range object
                        With ws
                            Set rngReplace = .range(.Cells(6, 2), .Cells(lngLastRow, 2))
                        End With
            
                    'Replace alpha and non-numerics
                        For Each C In rngReplace
                            If C.Value = 0 Or Not IsNumeric(C.Value) Then
                                C.Value = "Blank"
                            End If
                        Next C
                        
                    'Filter the range--------------------------------------------------------------------------------------------
                        'Update header row
                            ws.Cells(5, 1).Value = "Accounts"
                            
                        'Resize the range
                            With ws
                                Set rngReplace = .range(.Cells(5, 1), .Cells(lngLastRow - 1, 2))
                            End With
                            
                        'Filter the range
                            rngReplace.AutoFilter _
                                Field:=2, _
                                Criteria1:="<>Blank", _
                                VisibleDropDown:=True
                                
                     'Copy visible range to final sheet-------------------------------------------------------------------------------
                        'Determine last used row on final sheet
                            If blnFlag = True Then
                                lngRowsFinal = 2
                                blnFlag = False
                            Else
                                lngRowsFinal = GetLast(ws:=wsFinal, _
                                                       RC:="r") + 1
                            End If
                            
                        'Reshape the range to be copied to remove the header row
                            Set rngReplace = rngReplace.Offset(1, 0).Resize(rngReplace.Rows.Count - 1).SpecialCells(xlCellTypeVisible)
                            
                        '///Debug
                            Debug.Print ws.Name, "Rows", rngReplace.Rows.Count
                        
                        'Copy the range object
                            If rngReplace.Rows.Count = 2626 Then
                                'Do Nothing
                            Else
                                rngReplace.Copy
                                wsFinal.Cells(lngRowsFinal, 1).PasteSpecial (xlValues) 'Paste to final destination
                            
                                'Add organization name---------------------------------------------------------------------------------------------
                                    
                                    'Get last row of data on final worksheet
                                        lngRowsFinal = GetLast(ws:=wsFinal, _
                                                               RC:="r")
                                                       
                                    'Get next blank row for org name
                                        If blnFlagOrg = True Then
                                            lngRowsOrgName = 2
                                            blnFlagOrg = False
                                        Else
                                            lngRowsOrgName = GetLast(ws:=wsFinal, _
                                                                     RC:="r", _
                                                                     lngRowColumn:=3) + 1
                                        End If
            
                                    'Assign the org name from the retrieve sheet to the final sheet
                                        If IsEmpty(wsFinal.Cells(lngRowsOrgName, 3).Offset(0, -1)) Then
                                            'Do nothing
                                        Else
                                            wsFinal.Cells(lngRowsOrgName, 3).Value = ws.Cells(4, 1).Value
                                        End If
                                
                                    'Create a range object for the org name
                                        With wsFinal
                                            Set rngOrgName = .range(.Cells(lngRowsOrgName, 3), .Cells(lngRowsFinal, 3))
                                        End With
                                
                                    'Fill the org name range down
                                        If lngRowsFinal = lngRowsOrgName Then
                                            'Do nothing
                                        ElseIf IsEmpty(wsFinal.Cells(lngRowsOrgName, 3).Offset(0, -1)) Then
                                            'Do nothing
                                        Else
                                            If Not rngOrgName Is Nothing Then
                                                On Error Resume Next
                                                rngOrgName.FillDown
                                            End If
                                        End If
                            End If
                                    
                End If
            Next ws
        End With
        
    'Transformations----------------------------------------------------------------------------------------------------
        'Move organization member---------------------------------------------------------------------------------------
            'Insert column
                With wsFinal
                    .range("A1").EntireColumn.Insert Shift:=xlToRight
                End With
            
            'Get max columns org names
                lngRowsOrgName = GetLast(ws:=wsFinal, _
                                         RC:="r", _
                                         lngRowColumn:=4)
                                     
            'Create a range object for the org name
                With wsFinal
                    Set rngOrgName = .range(.Cells(2, 4), .Cells(lngRowsOrgName, 4))
                End With
            
            'Copy the org name to the first column
                rngOrgName.Copy
                wsFinal.range("A2").PasteSpecial xlPasteValues
            
            'Clear the original org name
                rngOrgName.Clear
                
        'Trim account labels---------------------------------------------------------------------------------------
            
            'Get max rows account labels
                lngRowsFinal = GetLast(ws:=wsFinal, _
                                       RC:="r")
                                     
            'Create a range object for the account labels
                With wsFinal
                    Set rngAccountLabels = .range(.Cells(2, 2), .Cells(lngRowsFinal, 2))
                End With
                
            'Trim labels
                For Each C In rngAccountLabels
                    C.Value = Trim(C.Value)
                Next C
                
        'Add time dimension---------------------------------------------------------------------------------------
            'Insert column
                With wsFinal
                    .range("B1").EntireColumn.Insert Shift:=xlToRight
                End With
            
            'Get max rows
                lngRowsFinal = GetLast(ws:=wsFinal, _
                                       RC:="r")
                                     
            'Create a range object for the time dimension
                With wsFinal
                    Set rngTimeDimension = .range(.Cells(2, 2), .Cells(lngRowsFinal, 2))
                End With
            
            'Copy the time dimension to the range
                rngTimeDimension.Value = wsRtrv.Cells(5, 2).Value
                
            'Fill down
                rngTimeDimension.FillDown

    'Tidy up
        'Add headers to final data
            With wsFinal
                .range("A1") = "Organization"
                .range("B1") = "Time"
                .range("C1") = "Account"
                .range("D1") = "Measure"
            End With
            
        'Destroy objects
            Set rngAccountLabels = Nothing
            Set rngFilter = Nothing
            Set rngReplace = Nothing
            Set rngCopy = Nothing
            Set rngForCopy = Nothing
            Set rngOrgName = Nothing
            Set rngTimeDimension = Nothing
            Set rngFunctionDimension = Nothing
            Set C = Nothing
            Set wsRtrv = Nothing
            Set wsUnits = Nothing
            Set wsAdmin = Nothing
            Set wsInputs = Nothing
            Set wsFinal = Nothing
            Set ws = Nothing
            Set wsConnect = Nothing
            Set wsDisconnect = Nothing
            Set wbNew = Nothing
            Set wb = Nothing

        'Restore Excel environment
            With Application
                .ScreenUpdating = True
                .EnableEvents = True
                .Calculation = xlCalculationAutomatic
            End With
        
        
End Sub




