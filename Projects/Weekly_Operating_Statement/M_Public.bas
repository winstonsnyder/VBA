Attribute VB_Name = "M_Public"
Option Explicit

Public Function GetApplication(strApplication As String)

    'Declare variables
        Dim MyApp As Object
    
    'Create a FileSystemObject
        Select Case strApplication
            Case "Excel"
                Set MyApp = CreateObject("Excel.Application")
            Case "Word"
                Set MyApp = CreateObject("Word.Application")
            Case "Powerpoint"
                Set MyApp = CreateObject("PowerPoint.Application")
        End Select
    
    'Pass the object to the function
        Set GetApplication = MyApp
        
    'Tidy up
        Set MyApp = Nothing
End Function

Public Function AddWorksheet(wb As Workbook, _
                             strSheetName As String) As Worksheet

    'Declare variables
        Dim ws As Worksheet
        Dim strMySheetName As String
        
    'Add worksheet
        With wb
            'If sheet exists, delete it
                On Error Resume Next
                If Not .Worksheets(strSheetName) Is Nothing Then
                    .Worksheets(strSheetName).Delete
                End If
    
            'Add worksheet
                    Set ws = .Sheets.Add(After:=.Sheets(wb.Sheets.Count))
                    ws.Name = strSheetName
        End With

        
    'Pass object to function
        Set AddWorksheet = ws
        
    'Tidy up
        Set ws = Nothing

End Function


Public Function GetLast(ws As Worksheet, _
                        RC As String, _
                        Optional lngRowColumn As Long = 1) As Long
                        
    'Requirements :   ws - A worksheet object
    '                 RC - A string as either "r" or "c" to specify row or column
    '                 lngRowColumn - Either the row or column number to be used
    
    'Declare variables
        Dim x       As Long

    'Get last used row
        Select Case RC
            Case "r"
                x = ws.Cells(Rows.Count, lngRowColumn).End(xlUp).Row
            Case Else
                x = ws.Cells(lngRowColumn, Columns.Count).End(xlToLeft).Column
        End Select
        
    'Pass value to function
        GetLast = x

End Function


Public Function GetFSO()

    'Declare variables
        Dim fso             As Object
    
    'Create a FileSystemObject
        Set fso = CreateObject("Scripting.FileSystemObject")
    
    'Pass the object to the function
        Set GetFSO = fso
        
    'Tidy up
        Set fso = Nothing
End Function

Public Function GetFDObjectName(strDialogType As String, _
                                strTitle As String) As String
    
    'Returns either the name of a folder or the name of a file based on the type passed into the function, "strDialogType"

    'Declare variables
        Dim fd As Object
        Dim strObjectName As String

    'Choose if user requested a folder dialog or other
        Select Case strDialogType
            Case "Folder"
                'Folder Dialog
                Set fd = Application.FileDialog(gclmsoFileDialogFolderPicker)
            Case Else
                'File Dialog
                Set fd = Application.FileDialog(gclmsoFileDialogFilePicker)
        End Select
        
    'Invoke filedialog
        With fd
            .Title = strTitle
            .AllowMultiSelect = False
            .Show
            strObjectName = .SelectedItems(1)
        End With

    'Pass value to function
        GetFDObjectName = strObjectName

    'Tidy up
        Set fd = Nothing
        
End Function

