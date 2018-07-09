<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosSql.asp"-->
<!--#include file="Includes/procedimientosMantenimiento.asp"-->
<%
Call initAccessInfo(RES_OT_SM)

idOt = GF_PARAMETROS7("idOt",0,6)
nroOt = GF_PARAMETROS7("nroOt","",6)
motivo = GF_PARAMETROS7("motivo","",6)
	
if motivo <> "" then	
	SM_idOrder = idOt
	SM_cdState = STATE_CANCELED
	SM_date = GF_DTE2FN(day(date()) & "/" & month(date()) & "/" & year(date()))
	SM_observations = motivo
	SM_cdExecutedBy = ""
	call updateOtStatus()
	accion = ACCION_CERRAR
end if	
%>
<html>
<head>
<link rel="stylesheet" href="css/main.css" type="text/css">
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript">
var refPopUpEquipo;
function equipoOnLoad() {
	refPopUpEquipo = getObjPopUp('popupEquipo');
	<% if (accion = ACCION_CERRAR) then %>
		refPopUpEquipo.hide();
	<% end if %>
}
function cerrarPopUp(){
	refPopUpEquipo.hide();
}
function HabilitarBoton(pObj){
	if (pObj.value != '') 
		document.getElementById("aceptar").disabled = false;
	else
		document.getElementById("aceptar").disabled = true;
	
	
}
</script>
</head>
<body onLoad="equipoOnLoad()">
<form name="frmSel" method="post">
	<h3> <% =GF_TRADUCIR("Cancelación de Orden de Trabajo") %> </h3>
    
    <div class="tableasidecontent">
		<div class="col16 reg_header_navdos"> <% =GF_TRADUCIR("Nro") %> </div>
        <div class="col46"> <%=nroOt%> </div>
        
        <div class="col16 reg_header_navdos"> <% =GF_TRADUCIR("Motivo") %> </div>
        <div class="col46" style="overflow:visible; height:auto;"> 
			<textarea onKeyUp="HabilitarBoton(this)" rows="4" cols="60" id="motivo" name="motivo" maxlength="1000"></textarea>
        </div>
		<%  if isAdminInAny then %>
			<span class="botonera">
				<input type="submit" disabled id="aceptar" name="aceptar" value="<% =GF_TRADUCIR("Aceptar") %>">
				<input type="button" id="cancelar" name="cancelar" value="<% =GF_TRADUCIR("Cancelar") %>" onclick="cerrarPopUp();">
			</span>
		<%	end if	%>
    </div>
<% call showMessages() %>
<input type="hidden" name="accion" value="<% =accion %>">
<input type="hidden" name="accionAnterior" value="<% =accion %>">
</form>
</body>
</html>