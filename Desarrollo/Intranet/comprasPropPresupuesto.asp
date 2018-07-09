<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosSql.asp"-->
<%
Call comprasControlAccesoCM(RES_ADM)
'-------------------------------------------------------------------------------------
Function accionGrabar(idItem, descripcion, pKeyWord)
	Dim strSQL, rs, conn
	
	if (idItem = 0) then
		'Es una nueva area
		strSQL="Insert into TBLBUDGET" & pKeyWord & "S(DS" & pKeyWord & ", IDESTADO, CDUSUARIO, MOMENTO)"
		strSQL= strSQL & " values('" & ucase(descripcion) & "', " & ESTADO_ACTIVO & ", '" & session("Usuario") & "', " & session("MmtoSistema") & ")"
	else
		'Es una modificacion
		strSQL="Update TBLBUDGET" & pKeyWord & "S Set DS" & pKeyWord & "='" & ucase(descripcion) & "', CDUSUARIO='" & session("Usuario") & "', MOMENTO=" & session("MmtoSistema")		
		strSQL = strSQL & " where ID" & pKeyWord & "=" & idItem
	end if
	'response.write strSQL
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
	accionGrabar = true
	
End Function
'-------------------------------------------------------------------------------------
Function accionConsulta(idItem, ByRef descripcion, pKeyWord)
	Dim strSQL, rs, conn
	strSQL="Select * from TBLBUDGET" & pKeyWord & "S where ID" & pKeyWord & "=" & idItem
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then descripcion = Trim(rs("DS" & pKeyWord & ""))
End Function
'***************************************************
'******   COMIENZO DE LA PAGINA
'***************************************************
Dim accion, errMsg, idArea, dsArea, tipo, myTitle, myOnLoad
idItem = GF_PARAMETROS7("idItem",0,6)
dsItem = UCase(GF_PARAMETROS7("descripcion","",6))
tipo = UCase(GF_PARAMETROS7("tipo","",6))
accion = GF_PARAMETROS7("accion","",6)
'Response.Write "aca" & tipo
if tipo = "A" then
	myTable = "AREA"
	myTitle = " Area de Presupuesto"
	myImage = "Budget_Area-32x32.png"
else
	myTable = "DETALLE"
	myTitle = " Item de Presupuesto"
	myImage = "Budget_Item-32x32.png"	
end if	
Call GP_ConfigurarMomentos
if (accion = ACCION_GRABAR) then
	Call accionGrabar(idItem, dsItem, myTable)
	myOnLoad = "refPopUpPresupuesto.hide()"
else
	Call accionConsulta(idItem, dsItem, myTable)
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
<script type="text/javascript">
var refPopUpPresupuesto;

function codigoOnBlur(ref) {
	if (ref.value == "") {
		document.getElementById("aceptar").disabled = true;
	} else {
		document.getElementById("aceptar").disabled = false;
	}			
}

function presupuestoOnLoad() {
	document.getElementById("descripcion").focus();
	
	refPopUpPresupuesto = getObjPopUp('popupPresupuesto');
	<% if (accion = ACCION_CERRAR) then %>
		refPopUpPresupuesto.hide();
	<% end if %>
}
	
</script>
</head>
<body onLoad="presupuestoOnLoad();<% =myOnLoad %>">
<form name="frmSel" method="post" action="comprasPropPresupuesto.asp">
<table  width="100">
	<tr>
		<td class="title_sec_section" colspan="2"><img align="absMiddle" src="images/compras/<%=myImage%>"> <% =GF_TRADUCIR("Propiedades del " & myTitle) %></td>
	</tr>
	<tr>
		<td colspan="2"><% call showErrors() %></td>
	</tr>
	<tr>
		<td></td>
		<td>
			<table>				
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("Descripcion") %></td>
					<td><input type="text" id="descripcion" name="descripcion" maxlength="50" size="30" value="<% =dsItem %>" onkeypress="return controlSalto(this, event)"></td>
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
</table>
<input type="hidden" name="accion" value="<% =ACCION_GRABAR %>">
<input type="hidden" name="tipo" value="<% =tipo %>">
<input type="hidden" name="idItem" value="<% =idItem %>">
</form>
</body>
</html>