Attribute VB_Name = "M_Test_Hyperion"
Option Explicit

Sub RemoveEssbaseConnection()

    Dim x As Long

    x = HypRemoveConnection(vtFriendlyName:=gEssbaseFriendlyName)
    Debug.Print "Remove Friendly connection : " & x
        
End Sub
Sub TestEssbaseRetrieve()
    
    'Declare Objects
        Dim wb As Workbook
        Dim wsHyperion As Worksheet
        Dim wsEssbase As Worksheet
        Dim wsOrganization As Worksheet
        
    'Variables
        Dim x As Long
        
    'Initialize objects
        Set wb = ThisWorkbook
        With wb
            Set wsHyperion = .Worksheets("Hyperion")
            Set wsEssbase = .Worksheets("Rtrv")
            Set wsOrganization = .Worksheets("Organization")
        End With
        
    'Suppress missing values and zeroes
        'Suppress missing values : Item 6
            x = SetSheetOption(ws:=wsEssbase, _
                               EssItem:=6, _
                               EssOption:=True)
                               
        'Suppress zero values : Item 7
            x = SetSheetOption(ws:=wsEssbase, _
                               EssItem:=7, _
                               EssOption:=True)
                               
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
        
        If x = 0 Then
            'Continue : 0 is expected values
        Else
            MsgBox "A connection to Essbase does not exist." & vbCrLf & _
                   "Please try again", Title:="Essbase Get Connection Error"
            x = GetEssDisconnect(wsDisconnect:=wsEssbase)
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
        
    'Get Organization member
        wsEssbase.Cells(6, 2).Value = wsOrganization.Cells(169, 1).Value
                
    'Update essbase time based on user selection
        wsEssbase.Cells(5, 2).Value = "Period 03"
        
    'Essbase Retrieve
        x = HypRetrieve(vtSheetName:=wsEssbase)
        Debug.Print "Essbase Retrieve Range : " & x
        
    'Disconnect Essbase
        x = GetEssDisconnect(wsDisconnect:=wsEssbase)
        Debug.Print "Disconnect Essbase : " & x
        
    'Remove Essbase connection
        x = HypRemoveConnection(vtFriendlyName:=gEssbaseFriendlyName)
        Debug.Print "Remove Friendly connection : " & x

    'Tidy up
        Set wsHyperion = Nothing
        Set wsEssbase = Nothing
        Set wsOrganization = Nothing
        Set wb = Nothing
        
End Sub
