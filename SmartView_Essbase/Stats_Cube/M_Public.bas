Attribute VB_Name = "M_Public"
Public Function DeleteFiles(strPath As String) As Boolean

    'Declare objects
        Dim fso As Object
        Dim fsoFolder As Object
        Dim fsoFile As Object
        
    'Declare variables
        Dim blnFlag As Boolean
        
   'Initialize objects
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set fsoFolder = fso.GetFolder(strPath)
        
    'Initialize variables
        blnFlag = False
        
    'Delete any files
        For Each fsoFile In fsoFolder.Files
            If fso.fileExists(fsoFile) Then
                fso.DeleteFile (fsoFile)
                blnFlag = True
            End If
        Next fsoFile
        
    'Pass value to function
        DeleteFiles = blnFlag

    'Tidy up
        Set fsoFolder = Nothing
        Set fso = Nothing

End Function


Public Function CreateDirectory(strPath As String) As Boolean

    'Declare variables
        Dim blnFlag As Boolean
        
    'Check if directory exists
    'If the directory does not exist - create it
        If Not Dir(strPath, vbDirectory) = vbNullString Then
            blnFlag = True
        Else
            MkDir (strDirectory)
            blnFlag = False
        End If
    
    'Pass value to function
        CreateDirectory = blnFlag

End Function

Public Function GetDisconnected(ws As Worksheet) As Long
    
    'Declare variables
        Dim x As Long

    'Disconnect Essabse
        x = EssVDisconnect(sheetName:=ws.Name)
        
    'Pass value to function
        GetDisconnected = x

End Function


Public Function GetZoomData(ByVal sheetName As Variant, _
                            ByVal range As Variant, _
                            ByVal selection As Variant, _
                            ByVal level As Variant, _
                            ByVal across As Variant) As Long
                        
    'Documentation
        'VBA Level Constants -> http://docs.oracle.com/cd/E17236_01/epm.1112/esb_ss_user/frameset.htm?idh_essv_levelconstants.html
        'Level 2 is All Levels
        'Level 3 is Bottom Level
        
    'Declare variables
        Dim x As Long
        
    'ZoomIn
        x = EssVZoomIn(sheetName, _
                       range, _
                       selection, _
                       level, _
                       across)
    
    'Pass value to function
        GetZoomData = x

        
End Function


Public Function GetConnected(wsEssConnectionValues As Worksheet, _
                             wsEssConnect As Worksheet) As Long

    'Declare variables
        Dim x As Long
        Dim EssUserName As Variant
        Dim EssPassword As Variant
        Dim EssServer As Variant
        Dim EssApplication As Variant
        Dim EssDatabase As Variant
        
    'Get Essbase values for connection
        With wsEssConnectionValues
            EssUserName = .range("B2").Value
            EssPassword = .range("B3").Value
            EssServer = .range("B4").Value
            EssApplication = .range("B5").Value
            EssDatabase = .range("B6").Value
        End With
        
    'Connect to Essbase
        x = EssVConnect(sheetName:=wsEssConnect.Name, _
                        username:=EssUserName, _
                        password:=EssPassword, _
                        server:=EssServer, _
                        Application:=EssApplication, _
                        database:=EssDatabase)
                        
    'Pass value to function
        GetConnected = x
        
    
End Function

Public Function GetLast(ws As Worksheet, _
                        RC As String, _
                        Optional ByVal lngRowColumn As Long = 1) As Long
                        
    'Requirements :   ws - A worksheet object
    '                 RC - A string as either "r" or "c" to specify row or column
    '                 lngRowColumn - Either the row or column number to be used
    
    'Declare variables
        Dim x       As Long

    'Get last row or column
        Select Case RC
            Case "r"
                x = ws.Cells(Rows.Count, lngRowColumn).End(xlUp).Row
            Case Else
                x = ws.Cells(lngRowColumn, Columns.Count).End(xlToLeft).Column
        End Select
        
    'Pass value to function
        GetLast = x

End Function



