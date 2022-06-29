Private Sub UserForm_QueryClose(Cancel as integer, CloseMode as integer)

if closemode = vbFormControlMenu then

	cancel = true
	msgbox = "Please use the cancel button on the form to close the form", vbOKOnly
	
end if

end sub