Private sub cmdOK_Click()

'Get user selected values from listbox

'objects
dim wb as workbook
set wb = thisworkbook

'variables
dim i as long
dim r as long

'Initialize
r = 1
TOC.range("K10:N10").clear

'Record selected values to a worksheet to be used as filter criteria
'If user did not select anything, respond with message alerting that nothing was selected before exiting
'Otherwise, record user selections to be used as filter criteria later

if lstBoxSuppliers.listIndex = -1 then
toc.range("K10") = "Nothing was selected"
toc.range("K10:N10").Interior.color = vbyellow
unload me
exit sub
else
dim ws as worksheet
AddWorksheetAtEnd wb:=wb, _
wsname:="UCriteria"
set ws = wb.worksheets("UCriteria")
with lstBoxSuppliers
for i = 0 to .listcount -1
if.slected(i) = true then
ws.cells(r,1) = cstr(.list(i))
r=r+1
end if
next i
end with
toc.range("K10") = "Selected suppliers have been recorded"toc.r("K10:N10")/interior.color = vbyellow
end if

'Tidy up

'Destroy objects
Set ws = nothing
Set wb = nothing

'Unload the supplier form
Unload me

Return to Table of Contents
GotoTOC

end sub