Attribute VB_Name = "M_Worksheets"
Option Explicit

Sub AddWorksheetAtEnd(wb As Workbook, _
                      Optional ByVal wsName As String = "temp")

    'Purpose    :   Add a new worksheet to the end of a workbook
    'Parameters :
    '--------------------------------------------------------------------------------------------------------------
    '
    '1.) wb     :   Required parameter. A workbook object.
    '2.) wsName :   Optional parameter. A string literal for the name of the worksheet. The default value is "temp"
    '
    '==============================================================================================================

    Dim ws As Worksheet

    With wb
        Set ws = .Worksheets.Add(after:=.Worksheets(.Worksheets.Count))
        ws.Name = wsName
    End With
    
    Set ws = Nothing

End Sub

Sub DeleteLastWorksheet(wb As Workbook)

    'Purpose    :   Delete the last worksheet at the end of a workbook
    'Parameters :
    '-----------------------------------------------------------
    '1.)    wb  :   Required parameter. A workbook object.
    '
    '===========================================================

    With wb
         .Worksheets(.Worksheets.Count).Delete
    End With

End Sub
