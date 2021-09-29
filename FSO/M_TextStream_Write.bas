Attribute VB_Name = "M_TextStream_Write"
Sub test_environ()

    Dim CompleteFileName As String
    Dim FileExists As String
    Dim my_Date As Date
    Dim PO_Number As Variant
    
    my_Date = Now()
    PO_Number = 12345698
    
    Const delim As String = "\"
    Const FileName As String = "Missing PO Number.txt"
    
    CompleteFileName = Environ("Homepath") & delim & FileName
    FileExists = Dir(CompleteFileName)
    
    Select Case FileExists
        Case ""
            CreateLogFile FName:=CompleteFileName                   'Create log file - Public FSO Function
            AppendToLogFile FName:=CompleteFileName, _
                            Log_Date:=my_Date, _
                            PO:=PO_Number                           'Append to log file - Public FSO Function
        Case Else
             AppendToLogFile FName:=CompleteFileName, _
                            Log_Date:=my_Date, _
                            PO:=PO_Number                           'Append to log file - Public FSO Function
    End Select

End Sub

Public Sub AppendToLogFile(FName As String, _
                           Log_Date As Date, _
                           PO As Variant)
    Dim fso As Object
    Dim ts As Object
    
    Const ForAppending = 8
    Const TristateFalse = 0
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(FName, ForAppending, True, TristateFalse)
    ts.write (Log_Date)
    ts.write ("          ")
    ts.write (PO)
    ts.WriteBlankLines 1
    ts.Close
                           
    Set ts = Nothing
    Set fso = Nothing
                           
End Sub
                            
                        
Public Sub CreateLogFile(FName As String)

    Dim fso As Object
    Dim ts As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(FName)
    
    ts.writeline ("CPM SOV Process")
    ts.writeline ("PO's from Supplier PO File Not Found In CPM CO Log File")
    ts.writeline ("========================================================")
    ts.writeline ("Date                          PO_Number")
    ts.writeline ("========================================================")
    ts.Close
    
    Set ts = Nothing
    Set fso = Nothing

End Sub

