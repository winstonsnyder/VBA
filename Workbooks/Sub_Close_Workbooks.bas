Attribute VB_Name = "M_WB"
Option Explicit

Sub CloseAllWbs()

    Dim wb As Workbook
    
    For Each wb In Workbooks
        If wb.Name <> ThisWorkbook.Name Then
            wb.Close SaveChanges:=True
        End If
    Next wb
    
    MsgBox "Current workbook was not closed"
    
End Sub
