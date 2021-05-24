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
