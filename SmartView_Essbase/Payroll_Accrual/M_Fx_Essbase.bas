Attribute VB_Name = "M_Fx_Essbase"
Option Explicit

Sub RemoveEssbaseConnection()

    Dim x As Long

    'Remove connection
        x = HypRemoveConnection(vtFriendlyName:=gEssbaseFriendlyName)
        Debug.Print "Remove Connection : " & x
        
End Sub

Public Function GetEssZoomIn(ws As Worksheet, _
                             rng As Range, _
                             lngLevel As Long, _
                             Optional ByVal blnAcross As Boolean = False) As Long
                             
    '0 = Next level
    '1 = All levels
    '2 = Bottom level
    '3 = Siblings (available only for Essbase 11.1.2.1.102 or later connections using Oracle Hyperion Provider Services)
    '4 = Same Level (available only for Essbase 11.1.2.1.102 or later connections using Provider Services)
    '5 = Same generation (available only for Essbase 11.1.2.1.102 or later connections using Provider Services)
    '6 = Formulas (available only for Essbase 11.1.2.1.102 or later connections using Provider Services)
    
    'vtAcross is not used
                             
    GetEssZoomIn = HypZoomIn(vtSheetName:=ws, _
                             vtSelection:=rng, _
                             vtLevel:=lngLevel, _
                             vtAcross:=blnAcross)
                              
                              
End Function

Public Function GetRetrieveRange(ws As Worksheet, _
                                 UserOption As Long) As Long
                      
   'Objects
        Dim rng As Range
        
    'Add range
        Select Case UserOption
            Case 1
                With ws
                    Set rng = .Range(.Cells(1, 1), .Cells(19, 2))
                End With
            Case Else
                With ws
                    Set rng = .Range("A1").CurrentRegion
                End With
        End Select
        
    'Retrieve range
        GetRetrieveRange = HypRetrieveRange(vtSheetName:=ws, _
                                            vtRange:=rng, _
                                            vtFriendlyName:=gEssbaseFriendlyName)

End Function

Public Function SetSheetOption(ws As Worksheet, _
                               EssItem As Long, _
                               EssOption As Boolean) As Long
                               
    'Item   Option
    '========================
    '6      Suppress Missing
    '7      Suppress Zeros
                               
    SetSheetOption = HypSetSheetOption(vtSheetName:=ws, _
                                       vtItem:=EssItem, _
                                       vtOption:=EssOption)

End Function

Public Function SetEssbaseSheet(wsDestination As Worksheet, _
                                wsParameters As Worksheet) As Long
                                 
    'Variables
        Dim EssDocumentType As String
        Dim EssFunctionalArea As String
        Dim EssCurrency As String
        Dim EssScenario As String
        Dim EssTime As String
        Dim EssAccount As String
        Dim EssOrganization As String
        
    
    'Get input values
        With wsParameters
            EssDocumentType = .Cells(2, 4).Value
            EssFunctionalArea = .Cells(3, 4).Value
            EssCurrency = .Cells(4, 4).Value
            EssScenario = .Cells(5, 4).Value
            EssTime = .Cells(6, 4).Value
            EssAccount = .Cells(7, 4).Value
            EssOrganization = .Cells(8, 4).Value
        End With
        
    'Setup retrieve sheet
        With wsDestination
            .Cells(1, 2).Value = EssDocumentType
            .Cells(2, 2).Value = EssFunctionalArea
            .Cells(3, 2).Value = EssCurrency
            .Cells(4, 2).Value = EssScenario
            .Cells(5, 2).Value = EssTime
        End With
        
    'Process is complete
        SetEssbaseSheet = 0
    
End Function

Public Function GetRetrieveSheet(wsEssRetrieve As Worksheet, _
                                 wsParameters As Worksheet) As Long
                                 
    'Variables
        Dim EssDocumentType As String
        Dim EssFunctionalArea As String
        Dim EssCurrency As String
        Dim EssScenario As String
        Dim EssTime As String
        Dim EssAccount As String
        Dim EssOrganization As String
        Dim lngRetrieveSheetSetup As Long
    
    'Get input values
        With wsParameters
            EssDocumentType = .Cells(2, 11).Value
            EssFunctionalArea = .Cells(3, 11).Value
            EssCurrency = .Cells(4, 11).Value
            EssScenario = .Cells(5, 11).Value
            EssTime = .Cells(6, 11).Value
            EssAccount = .Cells(7, 11).Value
            EssOrganization = .Cells(8, 11).Value
        End With
        
    'Setup retrieve sheet
        With wsEssRetrieve
            .Cells(1, 2).Value = EssDocumentType
            .Cells(2, 2).Value = EssFunctionalArea
            .Cells(3, 2).Value = EssCurrency
            .Cells(4, 2).Value = EssScenario
            .Cells(5, 2).Value = EssTime
            .Cells(6, 2).Value = EssOrganization
            .Cells(7, 1).Value = EssAccount
        End With
        
    'Process is complete
        lngRetrieveSheetSetup = 0
        
    'Pass value to function
        GetRetrieveSheet = lngRetrieveSheetSetup
 

Public Function GetEssRetrieve(ws As Worksheet) As Long
                                 
    GetEssRetrieve = HypRetrieve(vtSheetName:=ws)

End Function


Public Function GetEssDisconnect(wsDisconnect As Worksheet, _
                                 Optional ByVal blnLogout As Boolean = True)
                                 
    GetEssDisconnect = HypDisconnect(vtSheetName:=wsDisconnect, _
                                     bLogoutUser:=blnLogout)

End Function


Public Function GetEssbaseConnection(wsConnect As Worksheet, _
                                     wsParameters As Worksheet, _
                                     EssUserName As String, _
                                     EssUserPwd As String, _
                                     lngBeginColumnNumber As Long, _
                                     lngBeginRowNumber As Long) As Long
                                     
    'Purpose : This is a custom function to create asn Essbase connection using the Smart View Function
    'HypCreateConnection
                                     
    'Documentation
    'Oracle® Smart View for Office Developer 's Guide
    
        'HypCreateConnection | 5-7 | 101 of 282
        'Friendly Name       | 5-2 |  96 of 282
        
        'vtFriendlyName: The connection name of the data provider. The friendly name
        '                parameter can accept either of the following:
        '                   • A connection name created using HypCreateConnection
        '                   • A connection string consisting of a URL, server name, application name, and
        '                     database name, in the format URL|server|app|db.

        '                   The URL component of the connection string follows the guidelines in Private
        '                   Connection URL Syntax in the Oracle Smart View for Office User's Guide.
    
    'Variables
        Dim EssURL As String
        Dim EssServer As String
        Dim EssApplication As String
        Dim EssDatabaseName As String
        
    'Get parameter values
        With wsParameters
            EssURL = .Cells(lngBeginRowNumber, lngBeginColumnNumber).Value
            EssServer = .Cells(lngBeginRowNumber + 1, lngBeginColumnNumber).Value
            EssApplication = .Cells(lngBeginRowNumber + 2, lngBeginColumnNumber).Value
            EssDatabaseName = .Cells(lngBeginRowNumber + 3, lngBeginColumnNumber).Value
        End With
        
    'Get connection
        GetEssbaseConnection = HypCreateConnection(vtSheetName:=Empty, _
                                                   vtUserName:=EssUserName, _
                                                   vtPassword:=EssUserPwd, _
                                                   vtProvider:=HYP_ESSBASE, _
                                                   vtProviderURL:=EssURL, _
                                                   vtServerName:=EssServer, _
                                                   vtApplicationName:=EssApplication, _
                                                   vtDatabaseName:=EssDatabaseName, _
                                                   vtFriendlyName:=gEssbaseFriendlyName, _
                                                   vtDescription:="Essbase_1")

End Function

