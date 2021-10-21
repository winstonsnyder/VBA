Attribute VB_Name = "M_Fx_ReshapeData"
Option Explicit

Public Function ReshapeData(wsDataIn As Worksheet, _
                            wsDataOut As Worksheet, _
                            wsMeta As Worksheet, _
                            lngIndex As Long) As Long
 
    'Reshape Essbase Retrieve sheet into columnar layout for Pivot
        
    'Declare variables
        Dim RowsInput As Long
        Dim RowsOutput As Long
        Dim i As Long
        Dim k As Long
            
    'Get dimensions for reshaping operations
        RowsInput = GetRows(ws:=wsDataIn)
        If lngIndex = 5 Then
            RowsOutput = 5
        Else
            RowsOutput = GetRows(ws:=wsDataOut) + 1
        End If
        
            
    'Begin output row for reshaped data
        k = RowsOutput
 
    'Convert cross-tab report to normalized data table structure
        With wsDataOut
            For i = 7 To RowsInput 'Begin with first row of measures
                .Cells(k, 1).Value = wsDataIn.Cells(1, 2).Value  'Document Type
                .Cells(k, 2).Value = wsDataIn.Cells(2, 2).Value  'Functional Area
                .Cells(k, 3).Value = wsDataIn.Cells(3, 2).Value  'Currency
                .Cells(k, 4).Value = wsDataIn.Cells(4, 2).Value  'Scenario
                .Cells(k, 5).Value = wsDataIn.Cells(5, 2).Value  'Time
                .Cells(k, 6).Value = wsDataIn.Cells(6, 2).Value  'Cost Center
                .Cells(k, 7).Value = wsDataIn.Cells(i, 1).Value  'Account
                .Cells(k, 8).Value = wsMeta.Cells(1, 1).Value  'Retrieve Date
                .Cells(k, 9).Value = wsMeta.Cells(2, 1).Value  'Retrieve Time
                .Cells(k, 10).Value = wsDataIn.Cells(i, 2).Value  'Amount
                .Cells(k, 11).Value = "SDX Hyperion"
                k = k + 1
            Next i
        End With
        
    'Process is complete
        ReshapeData = 0

End Function



