<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosSql.asp"-->
<%
Call comprasControlAccesoCM(RES_AUD)
Function controlar(idNorma, cdNorma, dsNorma, unidad, valor)
	Dim strSQL, rs, conn
	
	controlar = RESPUESTA_OK
	if (idNorma = 0) then
		if (cdNorma = "") then
			controlar = CODIGO_VACIO
		else
			strSQL="Select IDNORMA, CDNORMA, DSNORMA, (VALOR*100) VALOR, UNIDAD from TBLNORMASAUDITORIA where CDNORMA='" & cdNorma & "'"
			Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
			if (not rs.eof) then 
				controlar = CODIGO_EXISTE
			else
				if (not isNumeric(valor)) then controlar = VALOR_NO_VALIDO
			end if
		end if
	end if	
End Function

Function accionGrabar(idNorma, cdNorma, dsNorma, unidad, valor)
	Dim strSQL, rs, conn
	
	if (idNorma = 0) then
		'Es una norma nueva
		strSQL="Insert into TBLNORMASAUDITORIA(CDNORMA, DSNORMA, UNIDAD, VALOR) values('" & cdNorma & "', '" & dsNorma & "', '" & unidad & "', " & valor & ")"
	else
		'Es una modificacion
		strSQL="Update TBLNORMASAUDITORIA Set CDNORMA='" & cdNorma & "', DSNORMA='" & dsNorma & "', UNIDAD='" & unidad & "', VALOR=" & valor & " where IDNORMA=" & idNorma
	end if
	'response.write strSQL	
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
	accionGrabar = true
	
End Function

Function accionConsulta(idNorma, ByRef cdNorma, ByRef dsNorma, ByRef unidad, ByRef valor)
	
	Dim strSQL, rs, conn
	
	strSQL="Select IDNORMA, CDNORMA, DSNORMA, (VALOR*100) VALOR, UNIDAD from TBLNORMASAUDITORIA where IDNORMA=" & idNorma
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then		
		cdNorma = Trim(rs("CDNORMA"))
		dsNorma = Trim(rs("DSNORMA"))
		unidad = rs("UNIDAD")
		valor = CDbl(rs("VALOR"))/100
	end if
	
End Function
'***************************************************
'******   COMIENZO DE LA PAGINA
'***************************************************
Dim accion, errMsg, idNorma, cdNorma, dsNorma, valor, unidad

idNorma = GF_PARAMETROS7("idNorma",0,6)
cdNorma = GF_PARAMETROS7("codigo","",6)
dsNorma = GF_PARAMETROS7("descripcion","",6)
unidad = GF_PARAMETROS7("unidad","",6)
valor = GF_PARAMETROS7("valor", 2 ,6)
accion = GF_PARAMETROS7("accion","",6)

Call GP_ConfigurarMomentos
if (accion = ACCION_GRABAR) then
	errMsg = controlar(idNorma, cdNorma, dsNorma, unidad, valor)
	if (errMsg = RESPUESTA_OK) then
		Call accionGrabar(idNorma, cdNorma, dsNorma, unidad, valor)
		accion = ACCION_CERRAR
	else
		setError(errMsg)
	end if
else
	Call accionConsulta(idNorma, cdNorma, dsNorma, unidad, valor)
end if
if (accion = "") then accion = ACCION_GRABAR
%>
<html>
<head>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/formato.js"></script>
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript">
var refPopUpNorma;
var ch = new channel();
 	
function codigoOnBlur(ref) {
	if (ref.value == "") {
		document.getElementById("aceptar").disabled = true;
	} else {
		document.getElementById("aceptar").disabled = false;
	}			
}

function normaOnLoad() {		
	refPopUpNorma = getObjPopUp('popupNorma');
	<% if (accion = ACCION_CERRAR) then %>
		refPopUpNorma.hide();
	<% end if %>
	//Se enfoca en el primer campo.
	var elem = document.getElementById("codigo");
	if (elem.type != "hidden")
		elem.focus();
	else
		document.getElementById("descripcion").focus();	
}
</script>
</head>
<body onLoad="normaOnLoad()">
<form name="frmSel" method="post" action="comprasPropNorma.asp">
<table align="center">
	<tr>
		<td class="title_sec_section" colspan="2"><img align="absMiddle" src="images/compras/audit-32x32.png"> <% =GF_TRADUCIR("Propiedades de Norma") %></td>
	</tr>	
	<tr>
		<td colspan="3"><% call showErrors() %></td>
	</tr>
	<tr>
		<td></td>
		<td>
			<table>				
				<tr>
					<td width="10%" class="reg_header"><% =GF_TRADUCIR("Norma") %></td>
					<td colspan="2"><% =idNorma %></td>
				</tr>				
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("Codigo") %></td>
					<td colspan="2"> <% if (idNorma = 0) then %>
						<input type="text" id="codigo" name="codigo" maxlength="10" size="10" value="<% =cdNorma %>" onblur="codigoOnBlur(this)" onkeypress="return controlSalto(this, event)">
						<% else %> 
							<b><% =cdNorma %></b> 
							<input type="hidden" id="codigo" name="codigo" value="<% =cdNorma %>">
						<% end if %>
					</td>					
				</tr>				
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("Descripcion") %></td>
						<td colspan="2"> 
						<input type="text" id="descripcion" name="descripcion" maxlength="50" size="50" value="<% =dsNorma %>" onkeypress="return controlSalto(this, event)">												
						</td>
					</td>
				</tr>
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("valor") %></td>
					<td>
						<select name="unidad">
							<option value="C"><% =GF_TRADUCIR("Cantidad") %>
							<option value="<% =MONEDA_PESO %>" <% if (unidad = MONEDA_PESO) then response.write "selected='true'" %>><% =getSimboloMoneda(MONEDA_PESO) %>
							<option value="<% =MONEDA_DOLAR %>" <% if (unidad = MONEDA_DOLAR) then response.write "selected='true'" %>><% =getSimboloMoneda(MONEDA_DOLAR) %>
						</select>						
					</td>
					<td>
						<input type="text" id="valor" name="valor" maxlength="17" size="15" value="<% =valor %>" onKeyPress="return controlDatos(this, event, 'N')">
					</td>
				</tr>
			</table>
		</td>
	</tr>		
	<tr><td>&nbsp;</td><tr>
	<tr>
		<td></td>
		<td align="right">
			<table>	
				<tr><td>
					<input type="submit" id="aceptar" name="aceptar" value="<% =GF_TRADUCIR("Aceptar") %>" <% if (idNorma = 0) then response.write "disabled=true" %>>					
				</td></tr>
			</table>
		</td>		
	</tr>	
</table>
<input type="hidden" name="idNorma" value="<% =idNorma %>">
<input type="hidden" name="accion" value="<% =accion %>">
</form>
</body>
</html>