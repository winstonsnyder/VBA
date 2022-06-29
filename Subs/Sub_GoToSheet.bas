Public Sub GoToSheet(ws as worksheet)
Application.Goto Reference = TOC.Range("A1"), _
				 Scroll:=True
				 
end sub