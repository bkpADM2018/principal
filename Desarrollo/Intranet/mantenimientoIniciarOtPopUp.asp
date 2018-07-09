<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosSql.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosPM.asp"-->
<!--#include file="Includes/procedimientosMantenimiento.asp"-->
<%
Call initAccessInfo(RES_INV_SM)
'TAREA 1748
idOt = GF_PARAMETROS7("idOt",0,6)
nroOt = GF_PARAMETROS7("nroOt","",6)
SM_cdExecutedBy_New = GF_PARAMETROS7("SM_cdExecutedBy","",6)
SM_dsExecutedBy = GF_PARAMETROS7("SM_dsExecutedBy","",6)
SM_idOrder = idOt
call readHeaderOT(SM_idOrder)
if SM_cdExecutedBy_New <> "" then	
	
	SM_date = GF_DTE2FN(day(date()) & "/" & month(date()) & "/" & year(date()))
	
		'Generar el PM
		'call readHeaderOT(SM_idOrder)
		PM_idAlmacen = getMaxFromDivision(SM_idDivision)
		PM_idObra = SM_idObra
		PM_FechaSolicitud = day(date()) & "/" & month(date()) & "/" & year(date())
		PM_FechaRequerido = PM_FechaSolicitud
		PM_idAlmacenDest = 0
		PM_idBudgetArea = SM_idBudgetArea
		PM_idBudgetDetalle = SM_idBudgetDetalle
		PM_comentario = "Generado automaticamente por la Orden de Trabajo Nro: " & SM_nroOrder
		PM_idSector = 0
		PM_cdSolicitante = SM_cdApplicant
		if (initItemsOT()) then
		    idPM = grabarHeaderPMInsert()			
		    while readNextItemOt()
			    call grabarPMDetalle(idPM, SM_idItem, CDbl(SM_programQuantityItem), 0)
		    wend		
		    'Actualizar el nro de PM de los items
	        SM_idPMItem = idPM
	        call udpateOtItemsPM
        end if		    	
	SM_date = GF_DTE2FN(day(date()) & "/" & month(date()) & "/" & year(date()))
	SM_observations = GF_PARAMETROS7("SM_observations", "",6)
	SM_cdState = STATE_STARTED
	SM_cdExecutedBy	= SM_cdExecutedBy_New
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
	autoCompleteExecutedBy();
}
function cerrarPopUp(){
	refPopUpEquipo.hide();
}
	
	function autoCompleteExecutedBy()
		{
			$( "#SM_dsExecutedBy" ).autocomplete({
					minLength: 2,
					source: "comprasStreamElementos.asp?tipo=JQPersonas",
					focus: function( event, ui ) {
						$( "#SM_dsExecutedBy").val(ui.item.nombre);
						return false;
					},
					select: function( event, ui ) {
						$( "#SM_dsExecutedBy"    ).val (ui.item.nombre);
						$( "#SM_cdExecutedBy"    ).val (ui.item.cdusuario);
						document.getElementById("aceptar").disabled = false;
						return false;
					},
					change: function( event, ui ) {
						if (!ui.item) {
							$( "#SM_dsExecutedBy").val ("");
							$( "#SM_cdExecutedBy").val ("");
							document.getElementById("aceptar").disabled = true;
						}
					}
				})
				.data( "autocomplete" )._renderItem = function( ul, item ) {
					return $( "<li></li>" )
						.data( "item.autocomplete", item )
						.append( "<a>" + item.cdusuario + " - <font style='font-size:10;'>" + item.nombre + "</font></a>" )
						.appendTo( ul );
				};
		}

</script>
</head>
<body onLoad="equipoOnLoad()">
<form name="frmSel" method="post">
	<h3> <% =GF_TRADUCIR("Inicio de Orden de Trabajo") %> </h3>
    
    <div class="tableasidecontent">
		<div class="col16 reg_header_navdos"> <% =GF_TRADUCIR("Nro") %> </div>
        <div class="col46"> <%=nroOt%> </div>
        
        <div class="col16 reg_header_navdos"> <%=GF_Traducir("Ejecutante")%> </div>
        <div class="col46">
			<input name="SM_dsExecutedBy" type="text" id="SM_dsExecutedBy" value="<%=SM_dsExecutedBy%>" style="width:200px">
			<input type="hidden" name="SM_cdExecutedBy" id="SM_cdExecutedBy" value="<%=SM_cdExecutedBy%>">
        </div>
		<%  
		proximaEjecucion = getNextExecution(GF_STANDARIZAR_FECHA_RTRN(date()),SM_OTFrequencyUnit, SM_OTFrequencyQuantity)
		if proximaEjecucion <> "" then %>    
			<div class="col16 reg_header_navdos"> <%=GF_Traducir("Próxima Ocurrencia")%> </div>
			<div class="col46">
				<%=proximaEjecucion%>
			</div>
			
        <%  end if %>    
		<%  if isAdminInAny then %>        
	        <span class="botonera">
				<input type="submit" id="aceptar" name="aceptar" value="<% =GF_TRADUCIR("Aceptar") %>">
				<input type="button" id="cancelar" name="cancelar" value="<% =GF_TRADUCIR("Cancelar") %>" onclick="cerrarPopUp();">
			</span>
	    <%	end if	%>   
	</div>
<input type="hidden" name="accion" value="<% =accion %>">
<input type="hidden" name="accionAnterior" value="<% =accion %>">
</form>
</body>
</html>