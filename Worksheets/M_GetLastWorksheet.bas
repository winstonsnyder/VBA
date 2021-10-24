Attribute VB_Name = "M_GetLastWorksheet"
Sub TryIt()

    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    Dim ws As Worksheet
    Set ws = Get_Last_Worksheet(wb:=wb)
    
    Debug.Print "ws name: "; ws.Name
    
    Set ws = Nothing
    Set wb = Nothing
    
End Sub
Public Function Get_Last_Worksheet(wb As Workbook) As Worksheet

    'Get last worksheet in a workbook as a worksheet object
        
    'Get last worksheet
        If Not wb Is Nothing Then
            With wb
                Set Get_Last_Worksheet = .Worksheets(.Worksheets.Count)
            End With
        Else
            MsgBox Prompt:="The workbook object does not exist.", _
                   Title:="Get Last Worksheet Function Error", _
                   Buttons:=vbOKOnly + vbExclamation
            Exit Function
        End If
    'Tidy up
        Set wb = Nothing

End Function

