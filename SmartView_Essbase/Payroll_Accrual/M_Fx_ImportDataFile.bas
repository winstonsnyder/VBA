Attribute VB_Name = "M_Fx_ImportDataFile"
Option Explicit

Sub GetDataFile()
    
    Dim wb As Workbook
    Dim x As Long
    
    Const MyOption As Long = 2
    
    Set wb = ThisWorkbook
    
    
    x = ImportDataFile(wb:=wb, _
                       CallSub:=MyOption)
    
    Set wb = Nothing
    
End Sub
Public Function ImportDataFile(wb As Workbook, _
                               CallSub As Long) As Long
                               
    'Call sub:
        'Main sub = 1
        'Other sub = anything else
                               
    'Objects
        Dim DataFolder As Object
        Dim DataFile As Object
        Dim wbData As Workbook
        Dim wsData As Worksheet
        Dim rngData As Range
        Dim rng As Range
        Dim ws As Worksheet
        Dim r As Long
        Dim c As Long
        Dim opt As Boolean

    'Inititialize objects
        Set DataFolder = CreateObject("Scripting.FileSystemObject")
        
    'Disable clipboard warning
        Application.CutCopyMode = False
        
        For Each DataFile In DataFolder.GetFolder(gDataPath).Files
            Select Case DataFile.Name
                Case "Accounts_Data.xlsx"
                    Set ws = wb.Worksheets("Accounts")
                    Set rng = ws.Range("A1")
                    If CallSub = 1 Then
                        opt = True
                    Else
                        opt = False
                    End If
                Case "Organization_Data.xlsx"
                    Set ws = wb.Worksheets("Organization")
                    Set rng = ws.Range("A1")
                    If CallSub = 1 Then
                        opt = True
                    Else
                        opt = False
                    End If
                Case "GL_Map_Data.xlsx"
                    Set ws = wb.Worksheets("Map_GL")
                    Set rng = ws.Range("A1")
                    If CallSub = 1 Then
                        opt = True
                    Else
                        opt = False
                    End If
                Case "Hierarchy_Map_Data.xlsx"
                    Set ws = wb.Worksheets("Map_Organization")
                    Set rng = ws.Range("A1")
                    If CallSub = 1 Then
                        opt = True
                    Else
                        opt = False
                    End If
                Case "Payroll_Accrual_Data.xlsx"
                    Set ws = wb.Worksheets("Flat")
                    r = GetRows(ws:=ws) + 1
                    With ws
                        Set rng = .Range(.Cells(r, 1), .Cells(r, 1))
                    End With
                    If CallSub <> 1 Then
                        opt = True
                    Else
                        opt = False
                    End If
                Case Else
                    opt = False
            End Select
            
            If opt Then
                'Data objects
                    Set wbData = Workbooks.Open(DataFile)
                    Set wsData = wbData.Worksheets(1)
                    Select Case DataFile.Name
                        Case "Payroll_Accrual_Data.xlsx"
                            'Remove header row
                            Set rngData = wsData.Range("A1").CurrentRegion
                            Set rngData = rngData.Offset(1, 0).Resize(rngData.Rows.Count - 1, _
                                                                      rngData.Columns.Count)
                        Case Else
                            Set rngData = wsData.Range("A1").CurrentRegion
                    End Select
                    
                
                'Transfer the data range to the new range
                    rngData.Copy
                    rng.PasteSpecial xlPasteValuesAndNumberFormats
                    
                'Autofit columnwidth
                    ws.Range("A1").CurrentRegion.Columns.AutoFit
                    
                'Goto A1
                    Application.Goto Reference:=ws.Range("A1"), _
                                     Scroll:=True
                    
                'Destroy objects
                    wbData.Close
                    Set rng = Nothing
                    Set ws = Nothing
                    Set rngData = Nothing
                    Set wsData = Nothing
                    Set wbData = Nothing
            End If
        Next DataFile
    
    'Destroy objects
        Set DataFolder = Nothing
        
    'Restore clipboard warning
        Application.CutCopyMode = True
        
    'Return value to function
        ImportDataFile = 0
                                
End Function

