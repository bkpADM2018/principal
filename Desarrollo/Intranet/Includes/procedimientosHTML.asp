<%
'*****************************************************************************************
Const SIN_SELECCION = 0

'----------------------------------------------------------------------------------
'El rs debe tener los campos CODIGO y DESC
Function GF_SELECT_RS(pName,pHabilitado,pRs,pSeleccion)
	Dim rtrn
	
	rtrn = "<select id='"&pName&"' name='"&pName&"' "
	if (pHabilitado = "disabled") then rtrn = rtrn & " disabled "
	rtrn = rtrn & " >"
	rtrn = rtrn & GF_OPTIONS(pRs,pSeleccion)
	rtrn = rtrn & "</select>"
	
	GF_SELECT_RS = rtrn
	
End Function
'----------------------------------------------------------------------------------
Function GF_OPTIONS(pRs,pSeleccion)
	Dim rtrn
	if isnull(pSeleccion) then pSeleccion = ""
	rtrn = "<option value='"&SIN_SELECCION&"'>- Seleccione -</option>"	

	while not pRs.EoF 
		if (trim(cstr(pSeleccion)) = trim(cstr(pRs("CODIGO")))) then
			rtrn = rtrn & "<option value="&pRs("CODIGO")&" selected='selected'>"&UCASE(pRs("DESC"))&"</option>"
		else
	    	rtrn = rtrn & "<option value="&pRs("CODIGO")&">"&UCASE(pRs("DESC"))&"</option>"
		end if
		pRs.MoveNext
	wend 
	
	GF_OPTIONS = rtrn
End Function
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------


%>