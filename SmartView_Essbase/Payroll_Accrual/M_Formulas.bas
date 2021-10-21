Attribute VB_Name = "M_Formulas"
Option Explicit

Sub ApplyADPHierarchyFormulas()
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim x As Long
    
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("Flat")
    
    x = GetApplyADPHierarchyFormulas(ws:=ws)
    
    Set ws = Nothing
    Set wb = Nothing
    
End Sub

Sub ApplyADPGLFormulas()
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim x As Long
    
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("Flat")
    
    x = GetApplyADPGLFormulas(ws:=ws)
    
    Set ws = Nothing
    Set wb = Nothing
    
End Sub

Public Function GetApplyADPGLFormulas(ws As Worksheet) As Long
    
    'Objects
        Dim rng As Range
    
    'Variables
        Dim r As Long
        Dim x As Long
        
    'Formula                                                                                Cell    Col Nmbr
    '=======================================================================================================
    Const EssGLNumber8Char As String = "=SUBSTITUTE(RIGHT($G5,9),""."","""")"                'M         13
    Const ADPAccount As String = "=VLOOKUP($M5,Map_GL!$B:$N,11,FALSE)"                       'Q         17
    Const ADPSubAccount As String = "=VLOOKUP($M5,Map_GL!$B:$N,12,FALSE)"                    'R         18
    Const ADPProduct As String = "=VLOOKUP($M5,Map_GL!$B:$N,13,FALSE)"                       'S         19

    'Get total number of rows on sheet
        r = GetRows(ws:=ws)
    
    'Clear previous use
        Set rng = Nothing
        With ws
            Set rng = Union(.Range(.Cells(5, 13), .Cells(r, 13)), _
                            .Range(.Cells(5, 17), .Cells(r, 17)), _
                            .Range(.Cells(5, 18), .Cells(r, 18)), _
                            .Range(.Cells(5, 19), .Cells(r, 19)))
        End With
            rng.Clear
        Set rng = Nothing

    'Create a range as a pivot point
        With ws
            Set rng = .Range(.Cells(5, 11), .Cells(r, 11))
        End With
        
    'Add formulas
        rng.Offset(0, 2).Formula = EssGLNumber8Char
        rng.Offset(0, 6).Formula = ADPAccount
        rng.Offset(0, 7).Formula = ADPSubAccount
        rng.Offset(0, 8).Formula = ADPProduct

    'Update Headers
        With ws
            .Cells(4, 13).Value = "GL Nmbr Essbase 8Char"
            .Cells(4, 17).Value = "ADP-Account"
            .Cells(4, 18).Value = "ADP-Sub Account"
            .Cells(4, 19).Value = "ADP-Product"
        End With
        
    'Clear header row
        x = ClearHeaderRow(ws:=ws)
        
    'Format Header row
        x = FormatHeaderRow(ws:=ws)
        
    'Autofit columnwidths
        ws.Range("A4").CurrentRegion.Columns.AutoFit
        
    'Tidy up
        Set rng = Nothing
        
    'Return function
        GetApplyADPGLFormulas = 0

End Function


Public Function GetApplyADPHierarchyFormulas(ws As Worksheet) As Long
    
    'Objects
        Dim rng As Range
    
    'Variables
        Dim r As Long
        Dim x As Long
        
    'Formula                                                                                Cell    Col Nmbr
    '=======================================================================================================
    Const EssCostCenter8Char As String = "=SUBSTITUTE(RIGHT($F5,9),""-"","""")"              'L         12
    Const ADPSatellite As String = "=VLOOKUP($L5,Map_Organization!$F:$I,2,FALSE)"            'N         14
    Const ADPRegion As String = "=VLOOKUP($L5,Map_Organization!$F:$I,3,FALSE)"               'O         15
    Const ADPDepartment As String = "=VLOOKUP($L5,Map_Organization!$F:$I,4,FALSE)"           'P         16

    'Get total number of rows on sheet
        r = GetRows(ws:=ws)
    
    'Clear previous use
        Set rng = Nothing
        With ws
            Set rng = Union(.Range(.Cells(5, 12), .Cells(r, 12)), _
                            .Range(.Cells(5, 14), .Cells(r, 14)), _
                            .Range(.Cells(5, 15), .Cells(r, 15)), _
                            .Range(.Cells(5, 16), .Cells(r, 16)))
        End With
            rng.Clear
        Set rng = Nothing

    'Create a range as a pivot point
        With ws
            Set rng = .Range(.Cells(5, 11), .Cells(r, 11))
        End With
        
    'Add formulas
        rng.Offset(0, 1).Formula = EssCostCenter8Char
        rng.Offset(0, 3).Formula = ADPSatellite
        rng.Offset(0, 4).Formula = ADPRegion
        rng.Offset(0, 5).Formula = ADPDepartment

    'Update Headers
        With ws
            .Cells(4, 12).Value = "Cost Cntr Essbase8Char"
            .Cells(4, 14).Value = "Satellite"
            .Cells(4, 15).Value = "ADP-Region"
            .Cells(4, 16).Value = "Department"
        End With
        
    'Clear header row
        x = ClearHeaderRow(ws:=ws)
        
    'Format Header row
        x = FormatHeaderRow(ws:=ws)
        
    'Autofit columnwidths
        ws.Range("A4").CurrentRegion.Columns.AutoFit
        
    'Tidy up
        Set rng = Nothing
        
    'Return function
        GetApplyADPHierarchyFormulas = 0

End Function
