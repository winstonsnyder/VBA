Attribute VB_Name = "M_Test"
Option Explicit



Public Sub AutoFilterStuff(rng As range, _
                           varCriteria As Variant, _
                           Optional ByVal lngField As Long = 1)
    
    rng.AutoFilter _
        Field:=lngField, _
        Criteria1:=varCriteria, _
        VisibleDropDown:=True
 End Sub

Sub TestRemoveAutoFilter()

    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("dataFinal")
    
    With ws
    End With
    
    Set ws = Nothing
    Set wb = Nothing
End Sub

Sub TestFilter()

    'Declare objects
        Dim wb As Workbook
        Dim wsRtrv As Worksheet
        Dim wsInputs As Worksheet
        Dim wsFinal As Worksheet
        Dim rngReplace As range
        Dim rngFilter As range
        Dim rngForCopy As range
        Dim lngRowsRetrieve As Long
        Dim lngRowsFinal As Long
        Dim blnFlag As Boolean
        
    'Initialize objects
        Set wb = ThisWorkbook
        Set wsRtrv = wb.Worksheets("Rtrv")
        Set wsInputs = wb.Worksheets("infInputs")
        Set wsFinal = wb.Worksheets("dataFinal")
        blnFlag = True
        
    'Clear previous use
        Call RemoveAutoFilter(ws:=wsRtrv)
        wsFinal.UsedRange.Clear
        
        
    'Create Range object
        lngRowsRetrieve = GetLast(ws:=wsRtrv, _
                                  RC:="r") - 1
                                  
        wsRtrv.Cells(6, 1).Value = "Accounts"
        
        With wsRtrv
            Set rngReplace = .range(.Cells(6, 2), .Cells(lngRowsRetrieve, 2))
            Set rngFilter = .range(.Cells(6, 1), .Cells(lngRowsRetrieve, 2))
        End With
        
    'Filter the range
        Call AutoFilterStuff(rng:=rngFilter, _
                             varCriteria:="<>0", _
                             lngField:=2)
                             
    'Copy visible range to final sheet
        If blnFlag = True Then
            lngRowsFinal = 1
        Else
            lngRowsFinal = GetLast(ws:=wsFinal, _
                                   RC:="r") + 1
            blnFlag = False
        End If
                                   
        'Reshape the range to remove the header row
            Set rngForCopy = rngFilter.Offset(1, 0).Resize(rngFilter.Rows.Count - 1).SpecialCells(xlCellTypeVisible)

        'Copy the range object
            rngForCopy.Copy
        
        'paste to final destination
            wsFinal.Cells(lngRowsFinal, 1).PasteSpecial (xlValues)
           
    'Detroy objects
        Set rngFilter = Nothing
        Set rngReplace = Nothing
        Set wsRtrv = Nothing
        Set wb = Nothing
        
End Sub


