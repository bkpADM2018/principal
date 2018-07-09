<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosSql.asp"-->
<%
Call comprasControlAccesoCM(RES_ADM)
'-----------------------------------------------------------------------------------------------------------------------
Function controlar(idTipoUnidad, cdTipoUnidad, dsTipoUnidad)
	Dim strSQL, rs, conn
	
	controlar = RESPUESTA_OK
	if (idTipoUnidad = 0) then
		if (cdTipoUnidad = "") then
			controlar = CODIGO_VACIO
		else
			strSQL="Select * from TBLTIPOSUNIDAD where CDTIPOUNIDAD='" & cdTipoUnidad & "'"
			Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
			if (not rs.eof) then controlar = CODIGO_EXISTE
		end if
	end if
End Function

Function accionGrabar(idTipoUnidad, cdTipoUnidad, dsTipoUnidad)
	Dim strSQL, rs, conn
	
	if (idTipoUnidad = 0) then
		'Es una unidad nueva
		strSQL="Insert into TBLTIPOSUNIDAD(CDTIPOUNIDAD, DSTIPOUNIDAD, CDUSUARIO, MOMENTO) values('" & cdTipoUnidad & "', '" & dsTipoUnidad & "', '" & session("Usuario") & "', " & session("MmtoSistema") & ")"
	else
		'Es una modificacion
		strSQL="Update TBLTIPOSUNIDAD Set CDTIPOUNIDAD='" & cdTipoUnidad & "', DSTIPOUNIDAD='" & dsTipoUnidad & "', CDUSUARIO='" & session("Usuario") & "', MOMENTO=" & session("MmtoSistema") & " where IDTIPOUNIDAD=" & idTipoUnidad
	end if
	'response.write strSQL
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
	accionGrabar = true
	
End Function

Function accionConsulta(idTipoUnidad, ByRef cdTipoUnidad, ByRef dsTipoUnidad)
	
	Dim strSQL, rs, conn
	
	strSQL="Select * from TBLTIPOSUNIDAD where IDTIPOUNIDAD=" & idTipoUnidad
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then		
		cdTipoUnidad = Trim(rs("CDTIPOUNIDAD"))
		dsTipoUnidad = Trim(rs("DSTIPOUNIDAD"))
	end if
	
	
End Function
'***************************************************
'******   COMIENZO DE LA PAGINA
'***************************************************
Dim accion, errMsg, idTipoUnidad, cdTipoUnidad, dsTipoUnidad

idUnidad = GF_PARAMETROS7("idUnidad",0,6)
idTipoUnidad = GF_PARAMETROS7("idTipoUnidad",0,6)
cdTipoUnidad = GF_PARAMETROS7("codigo","",6)
dsTipoUnidad = GF_PARAMETROS7("descripcion","",6)
accion = GF_PARAMETROS7("accion","",6)

Call GP_ConfigurarMomentos
if (accion = ACCION_GRABAR) then
	errMsg = controlar(idTipoUnidad, cdTipoUnidad, dsTipoUnidad)	
	if (errMsg = RESPUESTA_OK) then
		Call accionGrabar(idTipoUnidad, cdTipoUnidad, dsTipoUnidad)
		response.redirect "comprasPropUnidad.asp?idUnidad=" & idUnidad		
	else
		setError(errMsg)
	end if
else
	Call accionConsulta(idTipoUnidad, cdTipoUnidad, dsTipoUnidad)
end if
if (accion = "") then accion = ACCION_GRABAR
%>
<html>
<head>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript" src="scripts/formato.js"></script>
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" >

function codigoOnBlur(ref) {
	if (ref.value == "") {
		document.getElementById("aceptar").disabled = true;
	} else {
		document.getElementById("aceptar").disabled = false;
	}			
}

function tipoUnidadOnLoad() {
	var elem = document.getElementById("codigo");
	if (elem.type != "hidden")
		elem.focus();
	else
		document.getElementById("descripcion").focus();
}

function closeWin() {
	var refPopupNewTypeOfUnit = getObjPopUp('popupNewTypeOfUnit');
	if (refPopupNewTypeOfUnit) {
		refPopupNewTypeOfUnit.hide();
	}
	else
	{
		location.href = "comprasPropUnidad.asp?idUnidad=<% =idUnidad %>";
	}	
}
</script>
</head>
<body onLoad="tipoUnidadOnLoad()">
<form name="frmSel" method="post" action="comprasPropTipoUnidad.asp">
<table  width="100">
	<tr>
		<td class="title_sec_section" colspan="2"><img align="absMiddle" src="images/compras/units-32x32.png"> <% =GF_TRADUCIR("Crear Tipo de Unidad") %></td>
	</tr>
	<tr>
		<td colspan="2"><% call showErrors() %></td>
	</tr>
	<tr>
		<td></td>
		<td>
			<table>				
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("Nombre") %></td>
					<td><input type="text" id="codigo" name="codigo" maxlength="10" size="10" value="<% =cdTipoUnidad  %>" onblur="codigoOnBlur(this)" onkeypress="return controlSalto(this, event)"></td>					
				</tr>				
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("Descripcion") %></td>
					<td colspan="3"><input type="text" id="descripcion" name="descripcion" maxlength="50" size="45" value="<% =dsTipoUnidad %>" onkeypress="return controlSalto(this, event)"></td>
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
					<%  if (not isAuditor(SIN_DIVISION)) then %>
					<input type="submit" id="aceptar" name="aceptar" value="<% =GF_TRADUCIR("Aceptar") %>">
					<%	end if	%>
				</td></tr>
			</table>
		</td>		
	</tr>
<%
}
%>
</table>
<input type="hidden" name="accion" value="<% =accion %>">
<input type="hidden" name="idTipoUnidad" value="<% =idTipoUnidad %>">
<input type="hidden" name="idUnidad" value="<% =idUnidad %>">
</form>
</body>
</html>