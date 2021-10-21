Attribute VB_Name = "M_VBE"
Sub VBA_ExportProjectModules()


'Export all Modules In Project
'Set reference to Microsoft VBA Extensibility
'Give access in Trust Center Settings
'File >> Options >> Trust Center >> Trust Center Settings >> Macro Settings >> Trust Access to the VBA Project Object Model


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
