Attribute VB_Name = "M_MoveFiles_FSO"
Private Sub MoveFiles()

 'Purpose       :   Move files from users homepath to target folder
 'Comments      :   Objects use Late Binding -- Microsoft Scripting Runtime Library (scrrun.dll)
    
    'Objects
        Dim fso As Object
        Dim fldr As Object
        Dim f As Object
        
    'Variables
        Dim my_FileName As String
        Dim my_PathFileName_Source As String
        Dim my_PathFileName_Destination As String
        
    'Initialize
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set fldr = fso.getfolder(Environ("Homepath"))

    'Get name of PO Log file
        For Each f In fldr.Files
            If f.Name Like "Missing PO*" Then
                LogFileName = f.Name
                Exit For
            End If
        Next f
        
    'Source file & destination file
        PathFileName_Source = Environ("Homepath") & "\" & LogFileName
        PathFileName_Destination = gcsMissingPOLogs & "\" & LogFileName
        
    'Move the file
        fso.movefile PathFileName_Source, _
                     PathFileName_Destination

    'Tidy up
        Set fldr = Nothing
        Set fso = Nothing
End Sub

