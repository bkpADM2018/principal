<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosSql.asp"-->
<!--#include file="Includes/procedimientosMantenimiento.asp"-->
<%
Call initAccessInfo(RES_INV_SM)
'Call comprasControlAccesoCM(RES_ADM)
Call GP_ConfigurarMomentos
idEquipo = GF_PARAMETROS7("idEquipo",0,6)
accion = GF_PARAMETROS7("accion","",6)
accionAnterior = GF_PARAMETROS7("accionAnterior","",6)
if idEquipo <> 0 and accion <> ACCION_GRABAR then
	'Leer desde base
	call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMEQUIPMENT_GET_BY_PARAMETERS", idEquipo & "|| || ||0||")
	if not rs.eof then
		txtCdEquipo = rs("CDEQUIPMENT")
		txtDsEquipo = rs("DSEQUIPMENT")
	end if
else
	txtCdEquipo = GF_PARAMETROS7("txtCdEquipo","",6)
	txtDsEquipo = GF_PARAMETROS7("txtDsEquipo","",6)
end if
myPregunta = "Propiedades del Master"
if (accion = ACCION_GRABAR) then
	rtrn = controlarEquipo(idEquipo, txtCdEquipo, txtDsEquipo)
	if (rtrn = RESPUESTA_OK) then
		Call grabarEquipo(idEquipo, txtCdEquipo, txtDsEquipo, ESTADO_ACTIVO)
		accion = ACCION_CERRAR
	end if
elseif (accion = "H") then
	myPregunta = "Esta seguro que desea habilitar el siguiente Master?"
	if accionAnterior = accion then	
		Call grabarEquipo(idEquipo, txtCdEquipo, txtDsEquipo, ESTADO_ACTIVO)
		accion = ACCION_CERRAR
	end if	
elseif (accion = "D") then
	myPregunta = "Esta seguro que desea deshabilitar el siguiente Master?"
	if accionAnterior = accion then	
		rtrn = tieneActivaciones(idEquipo)
		if (rtrn <> RESPUESTA_OK) then
			Call grabarEquipo(idEquipo, txtCdEquipo, txtDsEquipo, ESTADO_BAJA)
			accion = ACCION_CERRAR
		else
			call setError(SM_TEMPLATE_TIENE_EQUIPO_ACTIVO)
		end if	
	end if	
end if
if (accion = "") then accion = ACCION_GRABAR

%>
<html>
<head>
<!--<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">-->
<link rel="stylesheet" href="css/main.css" type="text/css">
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
</script>
</head>
<body onLoad="equipoOnLoad()">
<form name="frmSel" method="post">
<div class="tableaside size100">
	<h3><% =GF_TRADUCIR(myPregunta) %></h3>
    
    <div class="tableasidecontent">
		<div class="col36 reg_header_navdos"> <% =GF_TRADUCIR("ID") %> </div>
        <div class="col36"> 
        	<% 
			if idEquipo = 0 then
				Response.Write "<i>Sin Asignar</i>"
			else
				Response.Write "<b>" & idEquipo & "</b>"
			end if	
			%>
		</div>
        
        <div class="col36 reg_header_navdos"> <% =GF_TRADUCIR("Código") %> </div>
        <div class="col36">
   			<%if accion = ACCION_GRABAR then%>
				<input type="text" id="txtCdEquipo" name="txtCdEquipo" maxlength="10" size="10" value="<% =txtCdEquipo %>"></td>
			<%else%>
				<% =txtCdEquipo %>
			<%end if%>
		</div>
        
        <div class="col36 reg_header_navdos"> <% =GF_TRADUCIR("Descripción") %></div>
        <div class="col36">
			<%if accion = ACCION_GRABAR then%>	
				<textarea cols="50" id="txtDsEquipo" name="txtDsEquipo" maxlength="100"><% =txtDsEquipo %></textarea></td>
			<%else%>
				<% =txtDsEquipo %>
			<%end if%>	
		</div>
        
        <span class="botonera">
   			<%  if isAdminInAny then %>
				<span class="botonera">
					<input type="submit" id="aceptar" name="aceptar" value="<% =GF_TRADUCIR("Aceptar") %>">
					<input type="button" id="cancelar" name="cancelar" value="<% =GF_TRADUCIR("Cancelar") %>" onclick="cerrarPopUp();">
				</span>
			<%	end if	%>
		</span>
        
    </div>
</div>
<%= showMessages%>

<input type="hidden" name="accion" value="<% =accion %>">
<input type="hidden" name="accionAnterior" value="<% =accion %>">
</form>
</body>
</html>