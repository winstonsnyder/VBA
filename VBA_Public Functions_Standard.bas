'//Function in this module
'//Get Last Column
'//Get Last Row
'//Add Worksheet
'//Create Range Object
'//Get User Range
'//CleanMAX by Daniel Ferry
'//CustomSplit
'//ScrollToCell
'//AddWorksheet
'//GetLast
'//Workbooks

=======================================================================
For Each wb in Workbooks
If wb.Name Like "*CL_Inventory_Merge*" Then
  Set wb_Main = wb
    Exit For
End if
Next wb  

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
      
      
Function CleanMAX(r As Range)

    'Author :   Daniel Ferry
    'Date   :   4/3/2020
    'URL    :   https://www.linkedin.com/pulse/excel-vba-clean-data-easily-daniel-ferry/
    
    CleanMAX = Replace("trim(clean(substitute(|,char(160),"" "")))", "|", r.Address)
    If r.Cells.Count > 1 Then CleanMAX = "index(" & CleanMAX & ",)"
    CleanMAX = Evaluate(CleanMAX)
End Function
        
Public Function CustomSplit(strParent As String, _
                            Optional ByVal lngPosition As Long = 1, _
                            Optional ByVal strDelimiter As String = "-") As String
                            
    'You may pass a different value as lngPosition, else the function will return the left part of the string
    'You may pass a different delimieter to the function else, the function will use "-" as the delimiter

    

    Dim str As String
    Dim x As Long
    
    x = InStrRev(StringCheck:=strParent, _
                 Stringmatch:=strDelimiter)
                 
    Debug.Print "strParent: ", strParent
    Debug.Print "Delimeter: ", strDelimiter
    Debug.Print "lngFirst: ", lngPosition
                 
    Select Case lngPosition
        Case 1
            str = Trim(Mid(strParent, 1, x - 1))
        Case Else
            str = Trim(Mid(strParent, x + 1, Len(strParent) - (x + 1)))
    End Select
        
    CustomSplit = str
                   
End Function
        
'==================================================================================================
Public Function ScrollToCell(wb As Workbook, _
                             Optional ByVal lngRow As Long = 1, _
                             Optional ByVal lngColumn As Long = 1) As String
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'ScrollToCell
    'Scroll to cell(lngRow,lngColumn) on each worksheet in thew workbook
    
    'Parameters :
    'Workbook   :   Required, a workbook object
    'lngRow     :   Row to scroll to, optional, Default = Row 1
    'lngColumn  :   Column to scroll to, optional, Default = column 1
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Declare variables
        Dim ws As Worksheet
        Dim strAddress As String
    
    'Scroll to cell(lngRow,lngColumn) on each worksheet in the workbook
        With wb
            For Each ws In .Worksheets
                Application.Goto Reference:=ws.Cells(lngRow, lngColumn), _
                                 Scroll:=True
            Next ws
        End With
        
    'Get address of activecell
        strAddress = ActiveCell.Address
        
    'Pass value to function
        ScrollToCell = strAddress
    
End Function
      
'=============================================================================================
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
            
'=================================================================================
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




