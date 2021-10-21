Attribute VB_Name = "M_Export_Modules"
Sub x()


' reference to extensibility library

Dim objMyProj As VBProject
Dim objVBComp As VBComponent

Set objMyProj = Application.VBE.ActiveVBProject

For Each objVBComp In objMyProj.VBComponents
If objVBComp.Type = vbext_ct_StdModule Then
objVBComp.Export "C:\temp\" & objVBComp.Name & ".bas"
End If
Next

End Sub

