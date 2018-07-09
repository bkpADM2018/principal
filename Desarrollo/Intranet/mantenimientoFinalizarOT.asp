<!--#include file="Includes/procedimientostraducir.asp"-->
<!--include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/ExternalFunctions.asp"-->
<!--#include file="Includes/procedimientosAS400.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosMantenimiento.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<% 
Call initAccessInfo(RES_OT_SM)
dim params, myIndice, myOnLoad
myIndice = 0
myOnLoad = "bodyOnLoad()"

SM_idOrder = GF_PARAMETROS7("idOT",0,6)
call addParam("idOT", SM_idOrder, params)
accion = GF_PARAMETROS7("accion","",6)
call addParam("accion", accion, params)
call readHeaderOT(SM_idOrder)
myTextAux = "Finalizar"
if accion = ACCION_VISUALIZAR then myTextAux = "Visualizar"
if accion = ACCION_GRABAR then
	if trim(SM_observations) = "" then
		'setError 
		setError(SM_OBS_REQUERIDAS)
	else
		SM_cdState = STATE_FINISHED
		SM_date = GF_DTE2FN(day(date()) & "/" & month(date()) & "/" & year(date()))
		'SM_observations = GF_PARAMETROS7("SM_observations", "",6)
		call updateOtStatus()	
		'Guardar Detalles
		call initTasksOT()
		while readNextTaskOt()
			saveOtTasks()
		wend
		'Guardar Repuestos
		call initItemsOT()
		while readNextItemOt()
			saveOtItems()
		wend	
		myOnLoad = "irA('mantenimientoAdministrarOts.asp');"
	end if
elseif accion = ACCION_CONTROLAR then
	if trim(SM_observations) = "" then setError(SM_OBS_REQUERIDAS)
end if
'---------------------------------------------------------------------------------------------------------
%>
<html>
<head>
<title>Sistema de Mantenimiento - <%=myTextAux%> Orden de Trabajo</title>
<link rel="stylesheet" href="css/main.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript" src="scripts/calendar.js"></script>
<script type="text/javascript" src="scripts/calendar-1.js"></script>
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>

<script type="text/javascript">
	function bodyOnLoad() {
		var tb = new Toolbar('toolbar', 7, "images/");	
		tb.addButtonRETURN("Volver", "irA('mantenimientoAdministrarOTs.asp')");
		<% if accion <> ACCION_VISUALIZAR then %>
		tb.addButtonSAVE("Guardar", "submitInfo('<% =ACCION_GRABAR %>')");
		tb.addButtonCONFIRM("Controlar",  "submitInfo('<% =ACCION_CONTROLAR %>')");			
		tb.addButtonREFRESH("Refrescar", "submitInfo()");		
		<% end if %>
		tb.draw();		
	}
	function submitInfo(pAccion) {		
		document.getElementById("accion").value = pAccion;
		document.getElementById("frmSel").submit();
	}	
	function irA(pLink) {
		location.href = pLink;
	}	
	function abrirPedido(id) {
		window.open("almacenValePedidoPrint.asp?idPedido=" + id, "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes,height=500,width=700",false);		
	}	
</script>
</head>
<%'Response.Buffer = False%>
<body onLoad="<%=myOnLoad%>">
<form method="post" id="frmSel">

<div id="toolbar"></div>

<div class="tableaside size100"> <!-- BUSCAR -->
<%=showMessages()%>
	<h3> <%=GF_Traducir(myTextAux & " Orden de Trabajo")%> </h3>
  
	<div class="tableasidecontent">
        <div class="col26 reg_header_navdos"> <%=GF_Traducir("ID")%> </div>
        <div class="col26"> 
        	<%=SM_idOrder%>
			<input type="hidden" name="idOT" id="idOT" value="<%=SM_idOrder%>">
		</div>
        
        <div class="col26 reg_header_navdos"> <%=GF_Traducir("Número")%> </div>
        <div class="col26">
        	<%=SM_nroOrder%>
			<input type="hidden" name="SM_nroOrder" id="SM_nroOrder" value="<%=SM_nroOrder%>">
		</div>
       
        <div class="col26 reg_header_navdos"> <%=GF_Traducir("Título")%> </div>
        <div class="col26"> 
			<%=SM_dsOrder%>
			<input type="hidden" name="SM_dsOrder" id="SM_dsOrder" value="<%=SM_dsOrder%>" size="50">
        </div>
        
        <div class="col26 reg_header_navdos"> <%=GF_Traducir("División")%> </div>
        <div class="col26"> 
			<%=getDivisionDS(SM_idDivision)%>
			<input type="hidden" name="SM_idDivision" id="SM_idDivision" value="<%=SM_idDivision%>">
        </div>
        
        <div class="col26 reg_header_navdos"> <%=GF_Traducir("Equipo")%> </div>
        <div class="col26"> 
			<input type="hidden" name="SM_idActiveEquipment" id="SM_idActiveEquipment" value="<%=SM_idActiveEquipment%>">
			<%
			call executeProcedureDb(DBSITE_SQL_INTRA, rsList, "TBLSMACTIVEEQUIPMENT_GET_FULL_BY_ID", SM_idActiveEquipment & "||0|| ||0|| || || ||1|| ")
			if not rsList.eof then
				Response.Write trim(rsList("CDACTIVATION")) & "/" & trim(rsList("DSACTIVATION"))
			end if	
			%>
        </div>
        
        <div class="col26 reg_header_navdos"> <%=GF_Traducir("Tipo Mant.")%> </div>
        <div class="col26"> 
			<input type="hidden" name="SM_maintenanceType" id="SM_maintenanceType" value="<%=SM_maintenanceType%>">
			<%=getDsMaintenanceType(SM_maintenanceType) %>
        </div>
        
        <div class="col26 reg_header_navdos"> <%=GF_Traducir("Tipo OT")%> </div>
        <div class="col26"> 
			<input type="hidden" name="SM_orderType" id="SM_orderType" value="<%=SM_orderType%>">
			<%=getDsOrderType(SM_orderType) %>
        </div>
        
        <div class="col26 reg_header_navdos"> <%=GF_Traducir("Fecha Prog.")%> </div>
        <div class="col26"> 
			<% 
			if not isdate(SM_scheduledDate) then SM_scheduledDate = GF_FN2DTE(SM_scheduledDate)
			Response.write SM_scheduledDate 
			%>
			<input type="hidden" id="SM_scheduledDate" name="SM_scheduledDate" value="<%=SM_scheduledDate%>">
			<input type="hidden" id="SM_startDate" name="SM_startDate" value="<%=SM_startDate%>">
			<input type="hidden" id="SM_finishedDate" name="SM_finishedDate" value="<%=SM_finishedDate%>">
        </div>
        
        <div class="col26 reg_header_navdos"> <%=GF_Traducir("Solicitante")%> </div>
        <div class="col26"> 
			<%=SM_dsApplicant%>
			<input name="SM_dsApplicant" type="hidden" id="SM_dsApplicant" value="<%=SM_dsApplicant%>" style="width:150px">
			<input type="hidden" name="SM_cdApplicant" id="SM_cdApplicant" value="<%=SM_cdApplicant%>">
        </div>
        
        <div class="col26 reg_header_navdos"> <%=GF_Traducir("Mano Obra")%> </div>
        <div class="col26"> 
			<%=SM_dsResponsableCompany%>
			<input name="SM_dsResponsableCompany" type="hidden" id="SM_dsResponsableCompany" value="<%=SM_dsResponsableCompany%>" style="width:150px">
			<input type="hidden" name="SM_idResponsableCompany" id="SM_idResponsableCompany" value="<%=SM_idResponsableCompany%>">
        </div>
        <% 
        classCol = "col36"
        if SM_OTFrequencyUnit <> ORDER_FREQ_UNIQUE then classCol = "col26"
		%>
        <div class="col26 reg_header_navdos"> <%=GF_Traducir("Part. Pres.")%> </div>
        <div class="<%=classCol%>"> 
				<input type="hidden" id="idObra" name="idObra" value="<%=SM_idObra%>">
				<input type="hidden" id="idBudgetArea" name="idBudgetArea" value="<%=SM_idBudgetArea%>">
				<input type="hidden" id="idBudgetDetalle" name="idBudgetDetalle" value="<%=SM_idBudgetDetalle%>">
				<%
				set rsObra = obtenerDescripcionCompletaDetalle(SM_idObra, SM_idBudgetArea, SM_idBudgetDetalle)
				if not rsObra.eof then
					Response.write rsObra("DSOBRA") & ": " & rsObra("DSAREA") & "-" & rsObra("DSDETALLE")
				end if
				%>
        </div>
        <% if SM_OTFrequencyUnit <> ORDER_FREQ_UNIQUE then %>
        <div class="col26 reg_header_navdos"> <%=GF_Traducir("Frecuencia")%> </div>
        <div class="col26"> 
			<%=getFrequency(SM_OTFrequencyUnit,SM_OTFrequencyQuantity)%>
        </div>        
        <% end if %>
		<input type="hidden" name="SM_cdState" id="SM_cdState" value="<%=SM_cdState%>">
		<input type="hidden" name="SM_cdUser" id="SM_cdUser" value="<%=SM_cdUser%>">
		<input type="hidden" name="SM_moment" id="SM_moment" value="<%=SM_moment%>">
	</div>
</div>
<%if CLNG(SM_idOrder) <> 0 then
	set files = getOTFiles(SM_idOrder)
	if not (files.EoF) then%>
		<div class="tableaside size100">
			<h3> <%=GF_Traducir("Archivos asociados a esta Orden")%> </h3>
				<table class="datagrid" width="90%" align="center">
					<thead>
						<tr>
							<th class="thicon"><%=GF_Traducir("Tipo")%> </th>
							<th><%=GF_Traducir("Nombre del archivo")%></th>
							<th class="thiconac" width="80"> - </th>
					    </tr>
					</thead>
					<tbody>	
					<%while not files.EoF %>
						<tr>
							<td class="thicon">
								<%=getImageByExt(files("EXT"))%>
							</td>	
							<td>
								<%=files("NAME") & "." & files("EXT")%>
							</td>
							<td class="thiconac">
								<a target='_blank' href='comprasOpenArchivo.asp?id=<%=files("ID")%>&secuencia=<%=files("FILENO")%>&type=SM-OT-OPEN'>
									<img width="16" height="16" src="images/download-16.png" title="Descargar Archivo">
								</a>
								<%if isAdminInAny  then%>
									<img src="images/cross-16.png" style="cursor:pointer;" onclick="deleteFile('<%=files("ID")%>','<%=files("FILENO")%>')" title="Eliminar Archivo">
								<%else
									Response.Write "."
								end if%>
							</td>
						</tr>
					<%files.MoveNext
					wend%>
					</tbody>
				</table>		
			</div>	
	<%end if
end if	%>
<div class="tableaside size100"> 
	<h3> <%=GF_Traducir("Descripción del trabajo a realizar")%> </h3> 
	<div class="tableaside size100"> 
	  <h3><%=GF_Traducir("Tareas")%></h3> 
		<table id="TASKS_TABLE" class="datagrid" width="90%" align="center">
			<%if CLNG(SM_idOrder) = 0 then%>
				<thead>
					<tr>
						<th colspan="3"><%=GF_Traducir("Disponible una vez guardada la cabecera de la OT.")%></td>
					</tr>	
				</thead>	
			<%else%>        
				<thead>
					<tr>
					    <th class="thicon"> <%=GF_Traducir("Nro")%> </th>
					    <th> <%=GF_Traducir("Descripción")%> </th>
					    <th class="thicon"> <%=GF_Traducir("Hecho")%> </th>
					</tr>
				</thead>
				<tbody> 
				<%
				call initTasksOT()
				while readNextTaskOt()
					%>
					<tr>
						<td class="thicon" align="center">
							<%=SM_nroTask%>
							<input type="hidden" id="SM_nroTask<%=SM_ActualTask%>" name="SM_nroTask<%=SM_ActualTask%>" size=4 value="<%=SM_nroTask%>">
						</td>
						<td>
							<%=SM_dsTask%>
							<input type="hidden" id="SM_dsTask<%=SM_ActualTask%>" name="SM_dsTask<%=SM_ActualTask%>" size=120 value="<%=SM_dsTask%>">
							<input type="hidden" id="SM_typeTask<%=SM_ActualTask%>" name="SM_typeTask<%=SM_ActualTask%>" size=4 value="<%=SM_typeTask%>">
							<input type="hidden" id="SM_result<%=SM_ActualTask%>" name="SM_result<%=SM_ActualTask%>" size=4 value="<%=SM_result%>">
						</td>
						<td align="center">
							<% if accion <> ACCION_VISUALIZAR then %>
							<input title="<%=GF_Traducir("Tarea Realizada")%>" type="checkbox" style="cursor:pointer;" name="SM_doneTask<%=SM_ActualTask%>" id="SM_doneTask<%=SM_ActualTask%>" value="<%=SM_TASK_DONE_YES%>" <% if SM_doneTask = SM_TASK_DONE_YES then Response.Write "checked"%>>
							<% end if %>
						</td>
					</tr>	
					<%
				wend
				%>
				</tbody>
			<%end if%>                    
        </table>
</div>	
    
<div class="tableaside size100"> 
	<h3> <%=GF_Traducir("Piezas / Repuestos")%> </h3>
		<table id="ITEMS_TABLE" class="datagrid" width="90%" align="center">           
			<%if CLNG(SM_idOrder) = 0 then%>
				<thead>
					<tr>
						<td colspan="6"><%=GF_Traducir("Disponible una vez guardada la cabecera de la OT.")%></td>
					</tr>	
				</thead>	
			<%else%>
				<thead>
					<tr>
						<th width="3%" rowspan="2" align="center" class="thicon"><%=GF_Traducir("Nro")%></th>
						<th align="center" colspan="2" class="thiconac"><%=GF_Traducir("Repuesto")%></th>
						<th align="center" colspan="2" class="thiconac"><%=GF_Traducir("Cantidad")%></th>
						<th width="3%" align="center" rowspan="2" class="thiconac"><%=GF_Traducir("PM")%></th>
					</tr>	
					<tr>
						<td align="center" width="3%"><%=GF_Traducir("Id")%></th>
						<td align="center"><%=GF_Traducir("Descripción")%></th>
						<td width="10%" align="center" class="thicon"><%=GF_Traducir("Progr.")%></th>
						<td width="10%" align="center" class="thicon"><%=GF_Traducir("Real")%></th>
					</tr>	
				</thead>
				<%
				call initItemsOT()
				while readNextItemOt()
					SM_typeItem = TASK_TYPE_REAL
						%>
						<tr>

							<td class="thicon" align="center">
								<%=SM_nroItem%>
								<input type="hidden" id="SM_nroItem<%=SM_ActualItem%>" name="SM_nroItem<%=SM_ActualItem%>" size=4 value="<%=SM_nroItem%>">
							</td>
							<td class="thicon">
								<%=SM_idItem%>
								<input type="hidden" id="SM_idItem<%=SM_ActualItem%>" name="SM_idItem<%=SM_ActualItem%>" size=8 value="<%=SM_idItem%>">
							</td>
							<td>
								<%=SM_dsItem%>
								<input type="hidden" id="SM_dsItem<%=SM_ActualItem%>" name="SM_dsItem<%=SM_ActualItem%>" size=120 value="<%=SM_dsItem%>">
							</td>
							<td align="right" nowrap>	
								<%=GF_EDIT_DECIMALS(cdbl(SM_programQuantityItem)*100,2)%>
								<input type="hidden" id="SM_programQuantityItem<%=SM_ActualItem%>" name="SM_programQuantityItem<%=SM_ActualItem%>" size=10 value="<%=SM_programQuantityItem%>">
							</td>
							<td align="right" nowrap>	
								<%=GF_EDIT_DECIMALS(cdbl(SM_realQuantityItem)*100,2)%>
								<input style="text-align:right;" type="hidden" id="SM_realQuantityItem<%=SM_ActualItem%>" name="SM_realQuantityItem<%=SM_ActualItem%>" size=10 value="<%=SM_realQuantityItem%>">
							</td>					
							<td class="thicon" align="center" style="cursor:pointer;">	
								<img title="<%=SM_idPmItem%>" src="images/pm-16.png" onClick="javascript:abrirPedido(<% =SM_idPmItem %>)">
								<input type="hidden" id="SM_idPmItem<%=SM_ActualItem%>" name="SM_idPmItem<%=SM_ActualItem%>" size=10 value="<%=SM_idPmItem%>">
								<input type="hidden" id="SM_typeItem<%=SM_ActualItem%>" name="SM_typeItem<%=SM_ActualItem%>" size=4 value="<%=SM_typeItem%>">
							</td>
						</tr>	
						<%
				wend
				%>
			<%end if%>
		</table>
</div>	

<%if accion <> ACCION_VISUALIZAR then %>

<div class="tableaside size100"> 
	<h3> <%=GF_Traducir("Observaciones del trabajo realizado:")%> </h3>
    
    <div class="tableasidecontent">
		<div class="col56 coment"> 
			<textarea cols="150" maxlength="998" id="SM_observations" name="SM_observations"><%=SM_observations%></textarea>
		</div>
    </div>	
</div>
<% end if %>
<input type="HIDDEN" name="accion" id="accion" value="<%=accion%>">

<input type="HIDDEN" name="SM_ActualTask" id="SM_ActualTask" value="<%=SM_ActualTask%>">
<input type="HIDDEN" name="SM_ActualItem" id="SM_ActualItem" value="<%=SM_ActualItem%>">
</form>
</body>
</html>