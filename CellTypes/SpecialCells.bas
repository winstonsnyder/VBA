Public Function GetLastUsedRow(ws as worksheet) as long
	
	'Purpose	: return the last used row on a worksheet
	'Parameters	:
	'ws		: Required parameter. A worksheet object
	'=====================================================
	
	GetLastUsedRow = ws.range("A1").SpecialCells(xlCellTypeLastCell).Row

End Function
