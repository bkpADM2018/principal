<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosSql.asp"-->
<!--#include file="Includes/procedimientosMantenimiento.asp"-->
<%
Call initAccessInfo(RES_OT_SM)

idOt = GF_PARAMETROS7("idOt",0,6)
nroOt = GF_PARAMETROS7("nroOt","",6)
cdStatus = GF_PARAMETROS7("pStatus","",6)
if cint(cdStatus) = cint(ORDER_FREQ_ENABLED) then 
	myTitle = "Activar planificación de Orden de Trabajo"
else
	myTitle = "Desactivar planificación de Orden de Trabajo"
end if
accion = GF_PARAMETROS7("accion","",6)
if accion = ACCION_GRABAR then
	SM_idOrder = idOt
	SM_cdState = cdStatus
	call updateSceduledOtStatus()
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
	function submitInfo(pAcc) {		
		document.getElementById("accion").value = pAcc;
		document.getElementById("frmSel").submit();
	}
</script>
</head>
<body onLoad="equipoOnLoad()">
<form id="frmSel" name="frmSel" method="post">
	<h3> <% =GF_TRADUCIR(myTitle) %> </h3>
    
    <div class="tableasidecontent">
		<div class="col16 reg_header_navdos"> <% =GF_TRADUCIR("Nro") %> </div>
        <div class="col46"> <%=nroOt%> </div>
        
		<%  if isAdminInAny then %>
			<span class="botonera">
				<input type="button" id="aceptar" name="aceptar" onclick="submitInfo('<%=ACCION_GRABAR%>');" value="<% =GF_TRADUCIR("Aceptar") %>">
				<input type="button" id="cancelar" name="cancelar" value="<% =GF_TRADUCIR("Cancelar") %>" onclick="cerrarPopUp();">
			</span>
		<%	end if	%>
    </div>
<input type="hidden" id="accion" name="accion" value="<% =accion %>">
</form>
</body>
</html>