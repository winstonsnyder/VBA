Sub ShowAll()

	Dim wb As Workbook
	Dim ws As Worksheet

	Set wb = ThisWorkbook
	
	For Each ws In wb.Worksheets
		If ws.AutoFilterMode Then ws.AutoFilter.ShowAllData
	Next ws

	Set wb = Nothing

End Sub

Sub ShowAll_Lists()

	Dim wb As Workbook
	Dim ws As Worksheet
	Dim lo As ListObject

	Set wb = ThisWorkbook

	For Each ws In wb.Worksheets
		For Each lo In ws.ListObjects
			If Not lo.AutoFilter Is Nothing Then
				lo.AutoFilter.ShowAllData
			Else
				lo.ShowAutoFilter = True
			End If
		Next lo
	Next ws
	Set wb = Nothing

End Sub