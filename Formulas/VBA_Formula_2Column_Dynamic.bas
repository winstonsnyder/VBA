
Option Explicit

Sub callit()

Dim wb As Workbook
Dim ws As Worksheet
Dim s1 As String
Dim s2 As String
Dim fx As String
Dim r As Long

Set wb = ThisWorkbook
Set ws = wb.Worksheets(1)
r = 4
s1 = "Cat"
s2 = "Dog"
fx = "-"

fx = xy(ws:=ws, _
        row_header:=r, _
        c1:=s1, _
        c2:=s2, _
        fx:=fx)
        
Debug.Print fx

Set ws = Nothing
Set wb = Nothing

End Sub
Public Function xy(ws As Worksheet, _
                    row_header As Long, _
                    c1 As String, _
                    c2 As String, _
                    fx As String) As String
                    

Dim r As Long   'First formula row
Dim rng As Range
Dim i As Long
Dim j As Long
Dim s1 As String
Dim s2 As String

Set rng = ws.Rows(row_header)
r = row_header + 1


'Find column number of first word/phrase
    i = FindColumnHeader(rng:=rng, _
                         SearchTerm:=c1)
                         
    If i = 0 Then
        MsgBox Prompt:="The search term, " & """" & c1 & """" & " was not found." & _
                       "Please double check to ensure that the search term exists in the context " & _
                       " and that it is spelled correctly.", _
               Title:="Search Term Missing Warning", _
               Buttons:=vbOKOnly + vbExclamation
        Exit Function
    End If
    
'Find column number of first word/phrase
    j = FindColumnHeader(rng:=rng, _
                         SearchTerm:=c2)
                         
    If j = 0 Then
        MsgBox Prompt:="The search term, " & """" & c2 & """" & " was not found." & _
                       "Please double check to ensure that the search term exists in the context " & _
                       " and that it is spelled correctly.", _
               Title:="Search Term Missing Warning", _
               Buttons:=vbOKOnly + vbExclamation
        Exit Function
    End If
    
'Get column letter from column number
    s1 = Split(Cells(1, i).Address, "$")(1)     'First column
    s2 = Split(Cells(1, j).Address, "$")(1)     'Second column
    
'Return Formula
    xy = "=" & s1 & r & " " & fx & " " & s2 & r

    
'Tidy up
    Set rng = Nothing

End Function
