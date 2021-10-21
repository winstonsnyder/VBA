Attribute VB_Name = "M_Rewrite"
Sub GetEssbaseData()
    
    'Declare Objects
        Dim wb As Workbook
        Dim wsHyperion As Worksheet
        Dim wsEssbase As Worksheet
        Dim wsFlat As Worksheet
        Dim wsFacts As Worksheet
        
        Dim wsOrganization As Worksheet
        Dim wsAccounts As Worksheet
        Dim wsOrganization_Map As Worksheet
        Dim wsAccounts_Map As Worksheet
        Dim rngRetrieve As Range
        Dim cell As Range
        Dim rng As Range

    'Variables
        Dim x As Long                                                                   'Check Essbase Functions
        Dim i As Long                                                                   'Count number of members in organization
        Dim j As Long                                                                   'Counter for organization loop
        Dim z As Long                                                                   'Loop counter
        Dim r As Long                                                                   'Total number of rows on a sheet
        Dim y As Long                                                                   'For debugging only
        
    'Constants
        Const ColumnToFilter As Long = 7
        Const MyOption As Long = 1
        
    'Initialize objects
        Set wb = ThisWorkbook
        With wb
            Set wsHyperion = .Worksheets("Hyperion")
            Set wsEssbase = .Worksheets("Rtrv")
            Set wsFlat = .Worksheets("Flat")
            Set wsFacts = .Worksheets("Facts")
            
            Set wsOrganization = .Worksheets("Organization")
            Set wsAccounts = .Worksheets("Accounts")
            Set wsOrganization_Map = .Worksheets("Map_Organization")
            Set wsAccounts_Map = .Worksheets("Map_GL")
        End With
        
    'Initialize variables
        z = 5
        y = 1
        
    'Excel environment - speed things up
        With Application
            .ScreenUpdating = False
            .DisplayAlerts = False
            .EnableEvents = False
            .Calculation = xlCalculationManual
        End With
        
    'Cleanup residual prior run
        x = ClearSheetPriorUse(ws:=wsFlat)
        x = ClearSheetPriorUse(ws:=wsOrganization)
        x = ClearSheetPriorUse(ws:=wsAccounts)
        x = ClearSheetPriorUse(ws:=wsOrganization_Map)
        x = ClearSheetPriorUse(ws:=wsAccounts_Map)
        
    'Import Essbase data files
        x = ImportDataFile(wb:=wb, _
                           CallSub:=MyOption)
        
    'Suppress missing values and zeroes
        'Suppress missing values : Item 6
            x = SetSheetOption(ws:=wsEssbase, _
                               EssItem:=6, _
                               EssOption:=True)
                               
        'Suppress zero values : Item 7
            x = SetSheetOption(ws:=wsEssbase, _
                               EssItem:=7, _
                               EssOption:=True)
                               
    'Create loop counter based on count of children in organization sheet
        i = GetRows(ws:=wsOrganization)
        
    'Set date and time
        With wsFacts
            .Cells(1, 1).Value = Format(Now(), "yyyy/mm/dd")
            .Cells(2, 1).Value = Format(Now(), "h:mm AM/PM")
        End With
        
    'Get month/period from user
        Load frmPeriod
        frmPeriod.Show
                               
    'Essbase login credentials
        Load frmEssConnect
        frmEssConnect.Show
                
    'Create Essbase connection
        x = GetEssbaseConnection(wsConnect:=wsEssbase, _
                                 wsParameters:=wsHyperion, _
                                 EssUserName:=EssLogin, _
                                 EssUserPwd:=EssPassword, _
                                 lngBeginColumnNumber:=1, _
                                 lngBeginRowNumber:=1)
                             
        Debug.Print "GetEssbaseConnection : " & x
        
        If x <> 0 Then
            MsgBox "A connection to Essbase does not exist." & vbCrLf & _
                   "Please try again", Title:="Essbase Get Connection Error"
            x = GetEssDisconnect(wsDisconnect:=wsEssbase)
            x = HypRemoveConnection(vtFriendlyName:=gEssbaseFriendlyName)
            Exit Sub
        End If
    
    'Connect to Essbase
        x = HypConnect(vtSheetName:=wsEssbase, _
                       vtUserName:=EssLogin, _
                       vtPassword:=EssPassword, _
                       vtFriendlyName:=gEssbaseFriendlyName)
                       
        Debug.Print "Connect to Essbase : " & x
        
        If x <> 0 Then
            MsgBox "A connection to Essbase does not exist." & vbCrLf & _
                   "Please try again", Title:="Essbase Connection Error"
            Exit Sub
        End If
        
    'Retrieve each child member in organization sheet
    'Stop at row 2 to allow for header row
        For j = i To 2 Step -1

            'Move to Retrieve Sheet
                Application.Goto Reference:=wsEssbase.Range("A7"), _
                                 Scroll:=True
                                 
            'Get Organization member
                wsEssbase.Cells(6, 2).Value = wsOrganization.Cells(j, 1).Value
                
            'Update essbase time based on user selection
                wsEssbase.Cells(5, 2).Value = gPeriodMonth
                
            'Get Accounts to retrieve on from Accounts worksheet
                'Delete last run
                    x = DeleteRange(ws:=wsEssbase, _
                                    FirstRow:=7, _
                                    FirstColumn:=1, _
                                    LastColumn:=2)
                                    
            'Update account for ZoomIn
                wsEssbase.Cells(7, 1).Value = "TOTAL PROCESSING COSTS : TPC9999"

            'Essbase ZoomIn
                Set rng = Nothing
                Set rng = wsEssbase.Range("A7")
                x = GetEssZoomIn(ws:=wsEssbase, _
                                 rng:=rng, _
                                 lngLevel:=2)
                Debug.Print "ZoomIn Essbase : " & x
                
            'Reshape the data into a columnar layout
                x = ReshapeData(wsDataIn:=wsEssbase, _
                                wsDataOut:=wsFlat, _
                                wsMeta:=wsFacts, _
                                lngIndex:=z)
                                
            'Increment loop counter
            'Used to copy data from retrieve sheet to flat sheet
                z = z + 1
                y = y + 1
                
        Next j
        Set rng = Nothing

    'Disconnect Essbase
        x = GetEssDisconnect(wsDisconnect:=wsEssbase)
        Debug.Print "Disconnect Essbase : " & x
        
    'Remove Essbase connection
        x = HypRemoveConnection(vtFriendlyName:=gEssbaseFriendlyName)
        Debug.Print "Remove Friendly connection : " & x

    'Add headers
        With wsFlat
            .Cells(4, 1).Value = "Document Type"
            .Cells(4, 2).Value = "Functional Area"
            .Cells(4, 3).Value = "Currency"
            .Cells(4, 4).Value = "Scenario"
            .Cells(4, 5).Value = "Time"
            .Cells(4, 6).Value = "Organization"
            .Cells(4, 7).Value = "Account"
            .Cells(4, 8).Value = "Date"
            .Cells(4, 9).Value = "Time"
            .Cells(4, 10).Value = "Final Amount"
            .Cells(4, 11).Value = "Source"
        End With
    
    'Format Headers
        x = FormatHeaderRow(ws:=wsFlat)
        
    'Go to flat sheet
        Application.Goto Reference:=wsFlat.Range("A5"), _
                         Scroll:=True
                         
    'trim account members
        r = GetRows(ws:=wsFlat)
        With wsFlat
            Set rng = .Range(.Cells(5, 7), .Cells(r, 7))
        End With
        
        For Each cell In rng
            cell.Value = Trim(cell.Value)
        Next cell
        
  'Format values on flat worksheet
        With wsFlat
            .Range("H1").EntireColumn.NumberFormat = "yyyy/mm/dd"
            .Range("I1").EntireColumn.NumberFormat = "[$-x-systime]h:mm:ss AM/PM"
            .Range("J1").EntireColumn.NumberFormat = "#,##0.00_);(#,##0.00)"
        End With
        
    'Autofit Columnwidth
        wsFlat.Range("A4").CurrentRegion.Columns.AutoFit

    'Destroy objects
        Set wsHyperion = Nothing
        Set wsEssbase = Nothing
        Set wsFlat = Nothing
        Set wsOrganization = Nothing
        Set wsAccounts = Nothing
        Set wsFacts = Nothing
        Set wb = Nothing
        
    'Nullify global variables
        EssLogin = vbNullString
        EssPassword = vbNullString
        gPeriodMonth = vbNullString
        
    'Restore Excel environment
        With Application
            .ScreenUpdating = True
            .DisplayAlerts = True
            .EnableEvents = True
            .Calculation = xlCalculationAutomatic
        End With

End Sub




