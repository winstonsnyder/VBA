Private sub UserForm_Initalize()

'Objects
dim wb as workbook
dim ws as worksheet
dim lo as listobject

'Variables
dim arr as variant

'Initialize
set wb = thisworkbook
set ws = wb.worksheets("dataSuppliers")
set lo = ws.listobjects(1)

'Populate combo box
	cboPlant.AddItem "OC"
	cboPlant.AddItem "RR"
	
'Tidy up
	Set lo = nothing
	Set ws = nothing
	Set wb = nothing