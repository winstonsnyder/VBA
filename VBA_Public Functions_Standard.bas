'//Function in this module
'//Get Last Column
'//Get Last Row
'//Add Worksheet
'//Create Range Object
'//Get User Range
=======================================================================
Attribute VB_Name = "M_Fx"
Public Function GetLastColumn(ws As Worksheet, _
                              row_number As Long) As Long

With ws
    GetLastColumn = .Cells(row_number, .Columns.Count).End(xlToLeft).Column
End With

End Function
Public Function GetLastRow(ws As Worksheet, _
                           column_number As Long) As Long

With ws
    GetLastRow = ws.Cells(.Rows.Count, column_number).End(xlUp).Row
End With

End Function

Public Function AddWorksheet(wb As Workbook, _
                             strSheetName As String) As Worksheet

    'Declare variables
        Dim ws As Worksheet
        Dim strMySheetName As String

    'Add worksheet if it does not exist
        On Error Resume Next
        Set ws = Sheets(strSheetName)
'        On Error GoTo 0
        If Not ws Is Nothing Then
            'The worksheet already exists
                ws.UsedRange.ClearContents
        Else
            'The worksheet does not exist
                Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
                ws.Name = strSheetName
        End If
        
    'Pass object to function
        Set AddWorksheet = ws
        
    'Tidy up
        Set ws = Nothing

End Function


Public Function CreateRangeObject(ws As Worksheet, _
                                  Optional ByVal RowBegin As Long = 1, _
                                  Optional ByVal ColumnBegin As Long = 1, _
  Optional ByVal ColumnEnd As Long = 1, _
  Optional ByVal RowEnd As Long = 1) As Range
                              
    'Declare variables
        Dim rng As Range
        Dim RowEnd As Long
        Dim ColumnEnd As Long
    
    'Get last row
        RowEnd = GetLast(ws:=ws, _
                         RC:="r", _
                         lngRowColumn:=ColumnReference)
                         
    'Get last column
        ColumnEnd = GetLast(ws:=ws, _
                            RC:="c", _
                            lngRowColumn:=RowReference)
                         
    'Create a range object
        With ws
            Set rng = .Range(.Cells(RowBegin, ColumnBegin), .Cells(RowEnd, ColumnEnd))
        End With
        
    'Pass object to function
        Set CreateRangeObject = rng
        
    'Tidy up
        Set rng = Nothing
                         
End Function
    
Public Function GetUserRange(ws As Worksheet) As Range

    'Prompt user to select a range of cells on a worksheet
      
    'Users - select a cell on a worksheet
        Set GetUserRange = Application.InputBox( _
                                                Prompt:="Please select a range on the worksheet", _
                                                Title:="Select a range", _
                                                Default:=ActiveCell.Address, _
                                                Type:=8) 'Range selection
 
 End Function

