Sub AddWorksheetAtEnd(wb as workbook, _
					  wsName as string)
					  
'Objects
	dim ws as worksheet
	dim wsCriteria as worksheet
	
'Variables
	dim flag as boolean
	dim i as long
	
'Initialize
	flag = false
	
'check if sheet already exists
	with wb
		for i = 1 to .worksheets.count
			if .worksheets(i).name = wsname then
				flag = true
		next i
	end with
	
'if the flag is true, the worksheet already exists
	with wb
		if not flag then
			set wsCriteria = .worksheets.add(after:=.worksheets(.worksheets.count))
			wsCriteria.name = wsname
		else
			set wecriteria = .worksheets(wsname)
			wscriteria.usedrange.clear
		end if
	end with
	
	
'Tidy up
	Set wsCriteria = nothing