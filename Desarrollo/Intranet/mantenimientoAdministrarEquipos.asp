<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosMantenimiento.asp"-->

<%
Call initAccessInfo(RES_INV_SM)
Dim txtIdEquipo, txtDsEquipo, myWhere, strSQL, rs, conn

txtIdEquipoActivo = GF_PARAMETROS7("txtIdEquipoActivo", 0, 6)
call addParam("txtIdEquipoActivo", txtIdEquipoActivo, params)

txtCdActivacion = UCASE(GF_PARAMETROS7("txtCdActivacion", "", 6))
call addParam("txtCdActivacion", txtCdActivacion, params)

txtDsActivacion = GF_PARAMETROS7("txtDsActivacion", "", 6)
call addParam("txtDsActivacion", txtDsActivacion, params)

txtIdEquipo = GF_PARAMETROS7("txtIdEquipo", 0, 6)
call addParam("txtIdEquipo", txtIdEquipo, params)

chkEstado = GF_PARAMETROS7("chkEstado", "" ,6)
if chkEstado = "" then chkEstado = 1
call addParam("chkEstado", chkEstado, params)

txtIdDivision = GF_PARAMETROS7("txtIdDivision", 0, 6)
if txtIdDivision = 0 then txtIdDivision = getListaCargosAdmin()
call addParam("txtIdDivision", txtIdDivision, params)

txtIdSector = GF_PARAMETROS7("txtIdSector", 0, 6)
call addParam("txtIdSector", txtIdSector, params)

txtIdUbicacion = GF_PARAMETROS7("txtIdUbicacion", 0, 6)
call addParam("txtIdUbicacion", txtIdUbicacion, params)

txtCdActivoFijo = UCASE(GF_PARAMETROS7("txtCdActivoFijo", "", 6))
call addParam("txtCdActivoFijo", txtCdActivoFijo, params)

myOrder = GF_PARAMETROS7("myOrder", "" ,6)
call addParam("myOrder", myOrder, params)
if myOrder = "" then myOrder = " AEQ.IDACTIVEEQUIPMENT DESC "
paginaActual = GF_PARAMETROS7("numeroPagina",0,6)
if (paginaActual = 0) then paginaActual=1
call addParam("paginaActual", paginaActual, params)
mostrar = GF_PARAMETROS7("registrosPorPagina",0,6)
if (mostrar = 0) then mostrar = 10
call addParam("mostrar", mostrar, params)

call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMACTIVEEQUIPMENT_GET_FULL_BY_PARAMETERS", txtIdEquipoActivo & "||" & txtIdEquipo & "||" & txtIdDivision & "||" & txtIdSector & "||" & txtCdActivacion & "||" & txtDsActivacion & "||" & txtCdActivoFijo & "||" & chkEstado & "||" & myOrder)

ref = GF_PARAMETROS7("ref", "" ,6)  
if ref = "" then ref = "mantenimientoInventarioIndex.asp"
Call setupPaginacion(rs, paginaActual, mostrar)
lineasTotales = rs.recordcount

'Variables de ordenamiento
orderById = GF_PARAMETROS7("orderById","",6)
orderByCodigo = GF_PARAMETROS7("orderByCodigo","",6)
orderByDesc = GF_PARAMETROS7("orderByDs","",6)
orderBySec = GF_PARAMETROS7("orderById","",6)
orderByAF = GF_PARAMETROS7("orderByCodigo","",6)
orderByDiv = GF_PARAMETROS7("orderByDs","",6)
%>

<html>
<head>
<title>Sistema de Mantenimiento - Administrar Equipos</title>
<link rel="stylesheet" href="css/main.css" type="text/css">
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/paginar.css" type="text/css">
<style type="text/css">
.divOculto {
	display: none;
}
.titleStyle {
	font-weight: bold;
	font-size: 20px;
}
</style>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/formato.js"></script>
<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript" src="scripts/paginar.js"></script>
<script type="text/javascript" src="scripts/Toolbar.js"></script>

<script type="text/javascript">
	/* Barra de herramientas de almacenes */
	function bodyOnLoad(){
		var toolBarEquipos = new Toolbar("toolBarEquipos", 8, 'images/');
		toolBarEquipos.addButtonRETURN("Volver", "irA('<%=ref%>')");
		toolBarEquipos.addButtonREFRESH("Refrescar", "submitInfo()");
		toolBarEquipos.addButton("master-16.png", "Ver Masters", "irA('mantenimientoAdministrarMasters.asp')");
 		toolBarEquipos.draw();
		<%if (not rs.eof) then	%>
			var pgn = new Paginacion("paginar");
			pgn.paginar(<% =paginaActual %>, <% =lineasTotales %>, <% =mostrar %>, 50, "mantenimientoAdministrarEquipos.asp<% =params %>");
		<%end if %>
	}
	function loadPopUpEquipo(id) {				
		var puw = new winPopUp('popupEquipo','mantenimientoNuevoMasterPopUp.asp?idEquipo=' + id,'780','300','Propiedades Equipo', "submitInfo()");
	}	
	function irA(pLink) {
		location.href = pLink;
	}	
	function submitInfo() {		
		document.getElementById("frmSel").submit();
	}
	function AbrirElemento(pOpcion, pId){
		var titulo;
		if (pOpcion=='B'){
			titulo = "Desactivar equipo";
			accion = "B";
		}else if (pOpcion=='A'){
			titulo = "Activar equipo";
			accion = "A";
		}else if (pOpcion=='H'){
			titulo = "Habilitar equipo";
			accion = "H";
		}else{
			titulo = "Modificar equipo activado";
			accion = 'M';
		}
		var puw = new winPopUp('popupEquipo','mantenimientoActivacionMaster.asp?idEquipoActivo=' + pId + '&tipoOperacion=' + accion,'780','380',titulo, "submitInfo()");
 	}
	function setOrder(pInput, pCol, pOrder){
		document.getElementById("myOrder").value = pCol + ' ' + pOrder;
		document.getElementById(pInput).value = pOrder;
		submitInfo();
	}
	function AbrirOT(pId){
		document.location.href = "mantenimientoAdministrarOTs.asp?txtIdEquipoActivo=" + pId + "&ref=mantenimientoAdministrarEquipos.asp";	
		//alert("Proximamente")
	}
	function AbrirPlanificaciones(pId){
		alert("Proximamente")
	}
	function AbrirUploader(pId, pCd){
		var puw = new winPopUp('popupEquipo','mantenimientoEquipoFiles.asp?idEquipoActivado=' + pId,'780','500','Archivos del Master: ' + pCd, "");
	}	
	function AbrirDespiece(pIdActivo){
		document.location.href = "mantenimientoAdministrarComponentes.asp?idEquipoActivo=" + pIdActivo + "&ref=mantenimientoAdministrarEquipos.asp";
	}	
</script>
</head>
<body onload="bodyOnLoad()">
<div id="toolBarEquipos"></div>
<form id="frmSel" name="frmSel">
	<input type="hidden" name="accion" id="accion" value="">
		<div class="tableaside size100"> <!-- BUSCAR -->

			<h3> Filtros </h3>
		  
			<div id="searchfilter" class="tableasidecontent">
		        <div class="col16 reg_header_navdos"> <%=GF_Traducir("ID:")%> </div>
		        <div class="col16"> 
					<input SIZE="2" type="text" value="<%=txtIdEquipoActivo%>" id="txtIdEquipoActivo" name="txtIdEquipoActivo">
		        </div>
		        <div class="col16 reg_header_navdos"> <%=GF_Traducir("Código:")%> </div>
		        <div class="col16"> 
			        <input type="text" size="9" value="<%=txtCdActivacion%>" id="txtCdActivacion" name="txtCdActivacion">
		        </div>
		        <div class="col16 reg_header_navdos"> <%=GF_Traducir("Descripción:")%> </div>
		        <div class="col16"> 
					<input type="text" size="25" value="<%=txtDsActivacion%>" id="txtDsActivacion" name="txtDsActivacion">	        
		        </div>
		        <div class="col16 reg_header_navdos"> <%=GF_Traducir("Master:")%> </div>
		        <div class="col16"> 
					<select name="txtIdEquipo" id="txtIdEquipo">
						<option value=0><%=GF_Traducir("Todos...")%></option>
						<%
						call executeProcedureDb(DBSITE_SQL_INTRA, rsList, "TBLSMEQUIPMENT_GET_BY_PARAMETERS", "0|| || ||1||ORDER BY CDEQUIPMENT")
						while not rsList.eof
							%>	
								<option value="<%=rsList("IDEQUIPMENT")%>" <%if cint(txtIdEquipo)=cint(rsList("IDEQUIPMENT")) then Response.Write "Selected"%>><%=rsList("CDEQUIPMENT") & " - " & rsList("DSEQUIPMENT")%></option>
							<%	
							rsList.movenext
						wend	
						%>
					</select>	        
		        </div>
		        <div class="col16 reg_header_navdos"> <%=GF_Traducir("División:")%> </div>
		        <div class="col16"> 
					<select name="txtIdDivision" id="txtIdDivision">
						<option value=0><%=GF_Traducir("Todas...")%></option>
						<%
						call executeProcedureDb(DBSITE_SQL_INTRA, rsList, "TBLDIVISIONES_GET_BY_LIST", getListaCargosAdmin())
						while not rsList.eof
							%>	
								<option value="<%=rsList("IDDIVISION")%>" <%if cint(txtIdDivision)=cint(rsList("IDDIVISION")) then Response.Write "Selected"%>><%=rsList("DSDIVISION")%></option>
							<%	
							rsList.movenext
						wend	
						%>
					</select>	        
		        </div>
		        <div class="col16 reg_header_navdos"> <%=GF_Traducir("Sector:")%> </div>
		        <div class="col16"> 
			        <select name="txtIdSector" id="txtIdSector">
						<option value=0><%=GF_Traducir("Todos...")%></option>
						<%
						call executeProcedureDb(DBSITE_SQL_INTRA, rsList, "TBLSMSECTOR_GET", "")
						while not rsList.eof
							%>	
								<option value="<%=rsList("IDSECTOR")%>" <%if cint(txtIdSector)=cint(rsList("IDSECTOR")) then Response.Write "Selected"%>><%=rsList("DSSECTOR")%></option>
							<%	
							rsList.movenext
						wend	
						%>
					</select>	
		        </div>
		        <div class="col16 reg_header_navdos"> <%=GF_Traducir("Activo Fijo:")%> </div>
		        <div class="col16"> 
					<input type="text" size="9" value="<%=txtCdActivoFijo%>" id="txtCdActivoFijo" name="txtCdActivoFijo">
		        </div>
		        <div class="col16 reg_header_navdos"> <%=GF_Traducir("Incluir Deshabilitados")%> </div>
		        <div class="col16"> 
					<input style="cursor:pointer;" type="checkbox" value="<%=ESTADO_BAJA%>" id="chkEstado" name="chkEstado" <%if chkEstado = 2 then Response.Write "Checked"%>>
		        </div>
		    	<span style="text-align:center; clear:both; float:left; width:100%"><input type="submit" value="<% =GF_TRADUCIR("Buscar") %>" id=button1 name=button1></span>
			</div>
		</div>
	<div class="col66"></div>

<table class="datagrid" width="90%" align="center">
    <thead>
        <tr>
            <th align="center" class="thicon" nowrap> <% =GF_TRADUCIR("Id") %> 
            	<%	
            	if orderById = "ASC" then 
					myTitle="Descendiente"
					orderById="DESC"
				else
					myTitle="Ascendiente"
					orderById="ASC"
				end if 
				%>
				<img style="cursor:pointer" title="<%=myTitle%>" onclick="setOrder('orderById', 'IDACTIVEEQUIPMENT','<%=orderById%>')" src="images\orderlist.png"> 
				<input type="hidden" id="orderById" name="orderById" value="<%=orderById%>">
			</th>
            <th align="center" class="thiconac" nowrap> <% =GF_TRADUCIR("Código") %> 
               	<%	
            	if orderByCodigo = "ASC" then 
					myTitle="Descendiente"
					orderByCodigo="DESC"
				else
					myTitle="Ascendiente"
					orderByCodigo="ASC"
				end if 
				%>
				<img style="cursor:pointer" title="<%=myTitle%>" onclick="setOrder('orderByCodigo', 'CDACTIVATION','<%=orderByCodigo%>')" src="images\orderlist.png"> 
				<input type="hidden" id="orderByCodigo" name="orderByCodigo" value="<%=orderByCodigo%>">
			</th>
            <th align="center" nowrap> <% =GF_TRADUCIR("Descripción") %> 
               	<%	
            	if orderByDesc = "ASC" then 
					myTitle="Descendiente"
					orderByDesc="DESC"
				else
					myTitle="Ascendiente"
					orderByDesc="ASC"
				end if 
				%>
				<img style="cursor:pointer" title="<%=myTitle%>" onclick="setOrder('orderByDesc', 'DSACTIVATION','<%=orderByDesc%>')" src="images\orderlist.png"> 
				<input type="hidden" id="orderByDesc" name="orderByDesc" value="<%=orderByDesc%>">
			</th>
            <th align="center" nowrap> <% =GF_TRADUCIR("División") %> 
               	<%	
            	if orderByDiv = "ASC" then 
					myTitle="Descendiente"
					orderByDiv="DESC"
				else
					myTitle="Ascendiente"
					orderByDiv="ASC"
				end if 
				%>
				<img style="cursor:pointer" title="<%=myTitle%>" onclick="setOrder('orderByDiv', 'DSDIVISION','<%=orderByDiv%>')" src="images\orderlist.png"> 
				<input type="hidden" id="orderByDiv" name="orderByDiv" value="<%=orderByDiv%>">
			</th>            
            <th align="center" nowrap> <% =GF_TRADUCIR("Sector") %> 
               	<%	
            	if orderBySec = "ASC" then 
					myTitle="Descendiente"
					orderBySec="DESC"
				else
					myTitle="Ascendiente"
					orderBySec="ASC"
				end if 
				%>
				<img style="cursor:pointer" title="<%=myTitle%>" onclick="setOrder('orderBySec', 'DSSECTOR','<%=orderBySec%>')" src="images\orderlist.png"> 
				<input type="hidden" id="orderBySec" name="orderBySec" value="<%=orderBySec%>">
			</th>            
            <th align="center" nowrap> <% =GF_TRADUCIR("A. Fijo") %> 
               	<%	
            	if orderByAF = "ASC" then 
					myTitle="Descendiente"
					orderByAF="DESC"
				else
					myTitle="Ascendiente"
					orderByAF="ASC"
				end if 
				%>
				<img style="cursor:pointer" title="<%=myTitle%>" onclick="setOrder('orderByAF', 'CDACTIVECODE','<%=orderByAF%>')" src="images\orderlist.png"> 
				<input type="hidden" id="orderByAF" name="orderByAF" value="<%=orderByAF%>">
			</th>            
            <th align="center" class="thicon"> <% =GF_TRADUCIR("Editar") %> </th>
            <th align="center" class="thicon"> <% =GF_TRADUCIR("Despiece") %> </th>
            <th align="center" class="thicon"> <% =GF_TRADUCIR("Adjuntos") %> </th>
            <th align="center" class="thicon"> <% =GF_TRADUCIR("OT") %> </th>
            <th align="center" class="thicon"> <% =GF_TRADUCIR("Planif.") %> </th>
            <th align="center" class="thicon"> - </th>
        </tr>
    </thead>
	<tbody>
	<%
	while ((not rs.eof)	and (reg < mostrar))
		reg = reg + 1			
		%>
		<tr>
			<td class="thicon" align="center"> <b><% =rs("IDACTIVEEQUIPMENT") %></b></td>
			<td align="center"> <% =rs("CDACTIVATION")%></td>
			<td align="center"> <% =rs("DSACTIVATION")%></td>
			<td align="center"><% =rs("DSDIVISION") %></td>				
			<td align="center"><% =rs("DSSECTOR") %></td>				
			<td align="center"><% =rs("CDACTIVECODE") %></td>				
				<%	if (rs("CDSTATE") = ESTADO_BAJA) then	%>
						<td align="center">	. </td>
						<td align="center">	. </td>
						<td align="center">	. </td>
						<td align="center">	. </td>
						<td align="center">	. </td>
						<td align="center">	
						<% if not isAuditor(rs("IDDIVISION")) then%>
							<img src="images/checkmark-16.png" style="cursor: pointer" title="Habilitar Equipo" onclick="AbrirElemento('H','<% =rs("IDACTIVEEQUIPMENT") %>')"> 
						<% else %>	
							. 
						<% end if %>		
						</td>
				<%	else  %>	
					<% if not isAuditor(rs("IDDIVISION")) then%>
						<td align="center"> <img src="images/edit-16.png" style="cursor: pointer" title="Editar Equipo Activo" onclick="AbrirElemento('M','<% =rs("IDACTIVEEQUIPMENT") %>')"> </td>
						<td align="center">	<img src="images/despiece-16.png" style="cursor: pointer" title="Despiece Template" onclick="AbrirDespiece('<% =rs("IDACTIVEEQUIPMENT") %>')"> </td>
						<td align="center"> <img src="images/adjunto-16.png" style="cursor: pointer" title="Adjuntar Archivos" onclick="AbrirUploader('<% =rs("IDACTIVEEQUIPMENT") %>','<% =rs("CDEQUIPMENT") %>')"> </td>
						<td align="center">	<img src="images/ot-16.png" style="cursor: pointer" title="Ver las Ordenes de Trabajo" onclick="AbrirOT('<% =rs("IDACTIVEEQUIPMENT") %>')"> </td>
						<td align="center">	<img src="images/calendar-16.png" style="cursor: pointer" title="Ver Planificaciones de Trabajo" onclick="AbrirPlanificaciones('<% =rs("IDACTIVEEQUIPMENT") %>')"> </td>	
						<td align="center"> <img src="images/cross-16.png" style="cursor: pointer" title="Desactivar Equipo" onclick="AbrirElemento('B','<% =rs("IDACTIVEEQUIPMENT") %>')"> </td>
					<% else %>	
						<td align="center">.</td>
						<td align="center">.</td>
						<td align="center">.</td>
						<td align="center">	<img src="images/ot-16.png" style="cursor: pointer" title="Ver las Ordenes de Trabajo" onclick="AbrirOT('<% =rs("IDACTIVEEQUIPMENT") %>')"> </td>
						<td align="center">	<img src="images/calendar-16.png" style="cursor: pointer" title="Ver Planificaciones de Trabajo" onclick="AbrirPlanificaciones('<% =rs("IDACTIVEEQUIPMENT") %>')"> </td>	
						<td align="center">.</td>
					<% end if %>	
				<%	end if	%>					
		</tr>
		<%
		rs.MoveNext()
	wend
	if (reg = 0) then		
		%>			
		<tr>
			<td colspan="12"><% =GF_TRADUCIR("No existen equipos activados.") %></td>
		</tr>
		<%
	end if 
	%>	
  	</tbody>	
    <tfoot>
  		<td colspan="12"><div id="paginar"></div></td>
  	</tfoot>	
</table>
<input type="hidden" name="myOrder" id="myOrder" value="<%=myOrder%>">	
</form>	
</body>
</html>