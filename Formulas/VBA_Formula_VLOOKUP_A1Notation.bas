Attribute VB_Name = "M_Formula_A1_Notation"
Option Explicit
Sub foo()

'=VLOOKUP("cat",A2:B4,2,FALSE)

Dim wb As Workbook
Dim ws As Worksheet
Dim rng As Range
Dim rngFrmla As Range
Dim colBegin As Long
Dim colEnd As Long
Dim rowEnd As Long
Dim i As Long           'Begin Col
Dim j As Long           'End Col
Dim strColBegin As String
Dim strColEnd As String
Dim LookupArray As String
Dim LookupFormula As String


Const EndColHeader As String = "Number"
Const BeginColHeader As String = "Animal"
Const LookupTerm As String = "cat"
Const rowBegin As Long = 2


Set wb = ThisWorkbook
Set ws = wb.Worksheets(1)
Set rng = ws.Rows("1:1")
rowEnd = GetRows(ws:=ws)

'Begin Column - number
i = FindColumnHeader(rng:=rng, _
                     SearchTerm:=BeginColHeader)
                     
If i = 0 Then
    MsgBox Prompt:="The search term, " & """" & BeginColHeader & """" & " was not found." & _
                   " Please double check to ensure that the search term exists in the context " & _
                   "and that it is spelled correctly.", _
           Title:="Search Term Missing Warning", _
           Buttons:=vbOKOnly + vbExclamation
    Exit Sub
End If
                     
'End Column - number
j = FindColumnHeader(rng:=rng, _
                     SearchTerm:=EndColHeader)
                     
If j = 0 Then
    MsgBox Prompt:="The search term, " & """" & EndColHeader & """" & " was not found." & _
                   " Please double check to ensure that the search term exists in the context " & _
                   "and that it is spelled correctly.", _
           Title:="Search Term Missing Warning", _
           Buttons:=vbOKOnly + vbExclamation
    Exit Sub
End If
                     
'Begin Column - letter
    strColBegin = Split(Cells(1, i).Address, "$")(1)
    
'End Column - letter
    strColEnd = Split(Cells(1, j).Address, "$")(1)
    
'Lookup Array String
    LookupArray = "$" & strColBegin & "$" & rowBegin & ":$" & strColEnd & "$" & rowEnd
    
'Lookup Formula - Exact match
'j = column to return
    LookupFormula = "=VLOOKUP(" & """" & LookupTerm & """" & "," & LookupArray & "," & j & ",FALSE)"
    
Debug.Print LookupFormula
    
'Range for formula
    With ws
    Set rngFrmla = .Range(.Cells(2, 6), .Cells(rowEnd, 6))
    End With

'Apply formula
    rngFrmla.Formula = LookupFormula
                           
'Tidy up
    Set rng = Nothing
    Set ws = Nothing
    Set wb = Nothing

End Sub
