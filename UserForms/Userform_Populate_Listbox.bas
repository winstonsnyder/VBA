sub populate_listbox

'Objects
dim wb as workbook
dim ws as worksheet
dim lo as listobject

'Variables
	dim arr as variant
	dim pos as long
	
Initialize objects
Set wb = thisworkbook
Set ws = wb.worksheets("dataSuppliers")

if cboPlant.listindex <> -1 then
selectitem = cboPlant.List(cboPlant.ListIndex)
pos = cboPlant.ListIndex + 1
end if

Select case pos
case 1
set lo = ws.listobjects("tblSuppliers_OC")
case 2
set lo = ws.listobjects("tblSuppliers_RR")
End select

'Populate listbox
if lo.databodyrange.rows.count = 1 then
redim arr(1 to1, 1 to 1)
arr(1,1) = lo.databodyrange.value
listboxsuppliers.list = arr
else
lstBoxSuppliers.list = lo.databodyrange.value
end if
end sub