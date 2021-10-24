Attribute VB_Name = "M_RenameLastWorksheet"
Option Explicit

Sub TestPivotTableRange1()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim rng As Range
    Dim cell As Range
    Dim my_var As Variant
    Dim pf As PivotField
    Dim i As Long
    
    Set wb = Workbooks("Test_Pivot_TableRange1.xlsm")
    Set ws = wb.Worksheets("Sheet1")
    i = 1
   
    For Each pt In ws.PivotTables
        Debug.Print i, pt.Name
        i = i + 1
    Next pt
   
    'Tidy up
        Set rng = Nothing
        Set pt = Nothing
        Set ws = Nothing
        Set wb = Nothing
    
    
End Sub


Sub Test_PT_Count()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim rng As Range
    Dim cell As Range
    Dim my_var As Variant
    Dim pf As PivotField
    Dim i As Long
    
    Set wb = Workbooks("CPM CO Log.xlsx")
    Set ws = wb.Worksheets("Pivot")
    Set pt = ws.PivotTables(1)
   
        Debug.Print pt.Name
   
    'Tidy up
        Set rng = Nothing
        Set pt = Nothing
        Set ws = Nothing
        Set wb = Nothing
    
    
End Sub


Sub TestCells()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim rng As Range
    Dim cell As Range
    Dim wbs As String

    
    Set wb = Workbooks("CPM CO Log.xlsx")
    Set ws = wb.Worksheets("Pivot")
    Set pt = ws.PivotTables(1)
    Set rng = pt.TableRange1
    Set rng = rng.Offset(2).Resize(rng.Cells.Count - 2)
    wbs = rng.Cells(1, 1).Value
    
    Debug.Print "wbs: "; wbs
   
    'Tidy up
        Set rng = Nothing
        Set pt = Nothing
        Set ws = Nothing
        Set wb = Nothing
    
    
End Sub

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
