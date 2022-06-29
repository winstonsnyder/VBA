Sub DeleteLastWorksheet(wb as workbook)

'Purpose	:	Delete the last worksheet at the end of a workbook
'Parameters	:
'==================================================================
'1.) wb		:	Required parameter. A workbook object.
'
'==================================================================

with wb
	.worksheets(.worksheets.count).delete
end with