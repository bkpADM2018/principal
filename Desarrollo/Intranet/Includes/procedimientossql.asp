<%
Const SQL_WHERE_DATE    = 2
Const SQL_WHERE_INTEGER = 1
Const SQL_WHERE_STRING  = 0

dim rs3, cn3, codigo
dim rs4, cn4, url, sqldel, importe, unidades, sqlupd1
'--------------------------------------------------------------------------------------------
Function GF_LIKE (p_campo, p_text)
' Esta funcion genera la sentencia Like de SQL para las busquedas por campo.
'p_campo = nombre del campo en la tabla donde se hace la busqueda.
'p_text = Texto de la busqueda.
'Si p_text tiene espacios en blanco procesa

dim i, v, vword, vbuscar, my_like
vword = ""
vbuscar = " " + p_text
while i < len(vbuscar)
    i = i + 1
    v = mid(vbuscar,i,1)
	if v <> " " then 
	   vword = vword & v
	end if   
	if v = " " or i = len(vbuscar) then
	   if len(vword) > 0 then 
	   my_like = my_like & " AND " & p_campo & " LIKE '%" & vword & "%' "
	   vword = ""
	   end if
	end if   
wend
GF_LIKE = my_like
end function	
'--------------------------------------------------------------------------------------------
function mkWhere(byref pstrWhere, pstrCampo, pstrValor,pstrSigno, pintTipo)
  dim strWhere
  strWhere = pstrWhere
  if len(strWhere) > 0 then 
    strWhere = strWhere & " AND "
  else
    strWhere = strWhere & " WHERE "
  end if
  strWhere = strWhere & pstrCampo & " "
  select case pintTipo
    case 2 'fecha
		strWhere = strWhere & pstrSigno & " " & GF_Fecha_HORA(pstrValor)
    case 1 'entero
		strWhere = strWhere & pstrSigno & " " & pstrValor
    case else 'string		
		if (UCase(pstrSigno) = "LIKE") then
			strWhere = strWhere & "LIKE '%" & pstrValor & "%'"
		else
			strWhere = strWhere & pstrSigno & " '" & pstrValor & "'"
		end if
  end select
  pstrWhere = strWhere 
  mkWhere = strWhere 
end function
%>