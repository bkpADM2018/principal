<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<%
Const TABLA_PM = "PM"
Const TABLA_VALES = "VALES"
Const TABLA_PCT = "PCT"
Const TABLA_REGISTROFIRMAS = "REGISTROFIRMAS"
Const TABLA_OBRAS = "OBRAS"
'*****************************************************************************
Function saveCDSolicitantePM()
	dim strSQL, conn, rs, strUPD, conUPD, rsUPD, idSolicitante, cdSolicitante, aux
	aux = false
	strSQL = "Select distinct IDSOLICITANTE from TBLPMCABECERA"
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then
		While (not rs.eof)
			idSolicitante = rs("IDSOLICITANTE")
			Call GF_MGKR(idSolicitante, "SG", cdSolicitante, "")
			if (cdSolicitante <> "") then
				strUPD = "update TBLPMCABECERA set CDSOLICITANTE = '" & cdSolicitante & "' " & _
						 " where IDSOLICITANTE = " & idSolicitante
				Call executeQueryDB(DBSITE_SQL_INTRA, rsUPD, "EXEC", strUPD)		 
			end if
			rs.MoveNext
		Wend
		aux = true
	end if
	if (aux) then
		%><table><tr><td>Se realizo correctamente. PEDIDOS DE MATERIALES</td></tr></table><%
		Call showButtons()
	else
		%><table><tr><td>No se pudo realizar. PEDIDOS DE MATERIALES</td></tr></table><%
		Call showButtons()
	end if
End Function
'*****************************************************************************
Function saveCDSolicitanteVales()
	dim strSQL, conn, rs, strUPD, conUPD, rsUPD, idSolicitante, cdSolicitante, aux
	aux = false
	strSQL = "Select distinct IDSOLICITANTE from TBLVALESCABECERA"
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then
		While (not rs.eof)
			idSolicitante = rs("IDSOLICITANTE")
			Call GF_MGKR(idSolicitante, "SG", cdSolicitante, "")
			if (cdSolicitante <> "") then
				strUPD = "update TBLVALESCABECERA set CDSOLICITANTE = '" & cdSolicitante & "' " & _
						 " where IDSOLICITANTE = " & idSolicitante
				Call executeQueryDB(DBSITE_SQL_INTRA, rsUPD, "EXEC", strUPD)
			end if
			rs.MoveNext
		Wend
		aux = true
	end if
	if (aux) then
		%><table><tr><td>Se realizo correctamente. VALES</td></tr></table><%
		Call showButtons()
	else
		%><table><tr><td>No se pudo realizar. VALES</td></tr></table><%
		Call showButtons()
	end if
End Function
'*****************************************************************************
Function saveCDSolicitantePCT()
	dim strSQL, conn, rs, strUPD, conUPD, rsUPD, idSolicitante, cdSolicitante, aux
	aux = false
	strSQL = "Select distinct IDSOLICITANTE from TBLPCTCABECERA"
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then
		While (not rs.eof)
			idSolicitante = rs("IDSOLICITANTE")
			Call GF_MGKR(idSolicitante, "SG", cdSolicitante, "")
			if (cdSolicitante <> "") then
				strUPD = "update TBLPCTCABECERA set CDSOLICITANTE = '" & cdSolicitante & "' " & _
						 " where IDSOLICITANTE = " & idSolicitante
				Call executeQueryDB(DBSITE_SQL_INTRA, rsUPD, "EXEC", strUPD)
			end if
			rs.MoveNext
		Wend
		aux = true
	end if
	if (aux) then
		%><table><tr><td>Se realizo correctamente. PEDIDOS DE COTIZACION</td></tr></table><%
		Call showButtons()
	else
		%><table><tr><td>No se pudo realizar. PEDIDOS DE COTIZACION</td></tr></table><%
		Call showButtons()
	end if
End Function
'*****************************************************************************
Function saveCDSolicitanteRegistroFirmas()
	dim strSQL, conn, rs, strUPD, conUPD, rsUPD, idUser, cdUser, aux
	aux = false
	strSQL = "Select distinct IDUSUARIO from TBLREGISTROFIRMAS"
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then
		While (not rs.eof)
			idUser = rs("IDUSUARIO")
			Call GF_MGKR(idUser, "SG", cdUser, "")
			if (cdUser <> "") then
				strUPD = "update TBLREGISTROFIRMAS set CDUSUARIO = '" & cdUser & "' " & _
						 " where IDUSUARIO = " & idUser
				Call executeQueryDB(DBSITE_SQL_INTRA, rsUPD, "EXEC", strUPD)
			end if
			rs.MoveNext
		Wend
		aux = true
	end if
	if (aux) then
		%><table><tr><td>Se realizo correctamente. REGISTRO FIRMAS</td></tr></table><%
		Call showButtons()
	else
		%><table><tr><td>No se pudo realizar. REGISTRO FIRMAS</td></tr></table><%
		Call showButtons()
	end if
End Function
'*****************************************************************************
Function saveCDSolicitanteObra()
	dim strSQL, conn, rs, strUPD, conUPD, rsUPD, idResponsable, cdResponsable, aux
	aux = false
	strSQL = "Select distinct IDRESPONSABLE from TBLDATOSOBRAS"
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then
		While (not rs.eof)
			idResponsable = rs("IDRESPONSABLE")
			Call GF_MGKR(idResponsable, "SG", cdResponsable, "")
			if (cdResponsable <> "") then
				strUPD = "update TBLDATOSOBRAS set CDRESPONSABLE = '" & cdResponsable & "' " & _
						 " where IDRESPONSABLE = " & idResponsable
			    Call executeQueryDB(DBSITE_SQL_INTRA, rsUPD, "EXEC", strUPD)
			end if
			rs.MoveNext
		Wend
		aux = true
	end if
	if (aux) then
		%><table><tr><td>Se realizo correctamente. DATOS OBRAS</td></tr></table><%
		Call showButtons()
	else
		%><table><tr><td>No se pudo realizar. DATOS OBRAS</td></tr></table><%
		Call showButtons()
	end if
End Function
'*****************************************************************************
Function showButtons()
	%>
	<br><br><br>
	<table width="80%" align="center">
		<tr align="center">
			<td><input type="button" value="EDITAR VALES" onClick="location.href='almacenCambioSolicitanteBD.asp?tabla=<% =TABLA_VALES %>'" id=button2 name=button2></td>
			<td width="15%">&nbsp;</td>
			<td><input type="button" value="EDITAR PM" onClick="location.href='almacenCambioSolicitanteBD.asp?tabla=<% =TABLA_PM %>'" id=button1 name=button1></td>
			<td width="15%">&nbsp;</td>
			<td><input type="button" value="EDITAR PCT" onClick="location.href='almacenCambioSolicitanteBD.asp?tabla=<% =TABLA_PCT %>'" id=button1 name=button1></td>
			<td width="15%">&nbsp;</td>
			<td><input type="button" value="EDITAR REG. FIRMAS" onClick="location.href='almacenCambioSolicitanteBD.asp?tabla=<% =TABLA_REGISTROFIRMAS %>'" id=button1 name=button1></td>
			<td width="15%">&nbsp;</td>
			<td><input type="button" value="EDITAR OBRAS" onClick="location.href='almacenCambioSolicitanteBD.asp?tabla=<% =TABLA_OBRAS %>'" id=button1 name=button1></td>
		</tr>
	</table>
	<%
End Function
'*****************************************************************************
'******************************************
'*** COMIENZO DE LA PAGINA
'******************************************
dim tabla
tabla = GF_PARAMETROS7("tabla","",6)

if (tabla = TABLA_PM) then 
	Call saveCDSolicitantePM()
elseif (tabla = TABLA_VALES) then
	Call saveCDSolicitanteVales()
elseif (tabla = TABLA_PCT) then
	Call saveCDSolicitantePCT()
elseif (tabla = TABLA_REGISTROFIRMAS) then
	Call saveCDSolicitanteRegistroFirmas()
elseif (tabla = TABLA_OBRAS) then
	Call saveCDSolicitanteObra()
else
	Call showButtons()
end if

%>