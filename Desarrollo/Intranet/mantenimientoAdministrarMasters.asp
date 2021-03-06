<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosMantenimiento.asp"-->
<%
Call initAccessInfo(RES_INV_SM)

Dim txtCdEquipo, txtDsEquipo, myWhere, strSQL, rs, conn

txtCdEquipo = UCase(GF_PARAMETROS7("txtCdEquipo", "" ,6))
call addParam("txtCdEquipo", txtCdEquipo, params)
txtDsEquipo = GF_PARAMETROS7("txtDsEquipo", "" ,6)
call addParam("txtDsEquipo", txtDsEquipo, params)

chkEstado = GF_PARAMETROS7("chkEstado", "" ,6)
if chkEstado = "" then chkEstado = 1
call addParam("chkEstado", chkEstado, params)
myOrder = GF_PARAMETROS7("myOrder", "" ,6)
if myOrder = "" then myOrder = " ORDER BY IDEQUIPMENT desc"
call addParam("myOrder", myOrder, params)
paginaActual = GF_PARAMETROS7("numeroPagina",0,6)
if (paginaActual = 0) then paginaActual=1
call addParam("paginaActual", paginaActual, params)
mostrar = GF_PARAMETROS7("registrosPorPagina",0,6)
if (mostrar = 0) then mostrar = 10
call addParam("mostrar", mostrar, params)
call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMEQUIPMENT_GET_BY_PARAMETERS", "0||" & txtCdEquipo & "||" & txtDsEquipo & "||" & chkEstado & "||" & myOrder)
lineasTotales = rs.recordcount
Call setupPaginacion(rs, paginaActual, mostrar)

'Variables de ordenamiento
orderById = GF_PARAMETROS7("orderById","",6)
orderByCodigo = GF_PARAMETROS7("orderByCodigo","",6)
orderByDs = GF_PARAMETROS7("orderByDs","",6)
'---------------------------------------------------------------------------------------------------------------------
%>
<html>
<head>
<title>Sistema de Mantenimiento - Administrar Masters</title>
<style type="text/css">
.divOculto {
	display: none;
}
.titleStyle {
	font-weight: bold;
	font-size: 20px;
}
</style>
<link rel="stylesheet" href="css/main.css" type="text/css">
<link rel="stylesheet" href="css/paginar.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css">

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
		var toolBarEquipos = new Toolbar("toolBarEquipos", 7, 'images/');
		toolBarEquipos.addButtonRETURN("Volver", "irA('mantenimientoInventarioIndex.asp')");
		<% if isAdminInAny then %>
		toolBarEquipos.addButton("master-add-16.png", "Nuevo Master", "loadPopUpEquipo()");
		<% end if %>
		toolBarEquipos.addButtonREFRESH("Refrescar", "submitInfo()");
		<%if (not rs.eof) then	%>
			var pgn = new Paginacion("paginar");
			pgn.paginar(<% =paginaActual %>, <% =lineasTotales %>, <% =mostrar %>, 50, "mantenimientoAdministrarMasters.asp<% =params %>");
		<%end if %>
		toolBarEquipos.addButton("active-16.png", "Ver Equipos Activados", "irA('mantenimientoAdministrarEquipos.asp')");		
 		toolBarEquipos.draw();
	}
	function loadPopUpEquipo(id) {				
		var puw = new winPopUp('popupEquipo','mantenimientoNuevoMasterPopUp.asp?idEquipo=' + id,'780','320','Agregar Nuevo Master', "submitInfo()");
	}	
	function irA(pLink) {
		location.href = pLink;
	}	
	function submitInfo() {		
		document.getElementById("frmSel").submit();
	}
	function AbrirElemento(pOpcion, pId){
		var titulo;
		if (pOpcion=='D'){
			titulo = "Deshabilitar ";
			accion = "D";
		}else if (pOpcion=='H'){
			titulo = "Habilitar ";
			accion = "H";
		}else{
			titulo = "Modificar ";
			accion = '';
		}
		var puw = new winPopUp('popupEquipo','mantenimientoNuevoMasterPopUp.asp?idEquipo=' + pId + '&accion=' + accion,'780','320',titulo + 'Master', "submitInfo()");
	}
	function setOrder(pInput, pCol, pOrder){
		document.getElementById("myOrder").value = " ORDER BY " + pCol + ' ' + pOrder;
		document.getElementById(pInput).value = pOrder;
		submitInfo();
	}
	function AbrirDespiece(pId){
		document.location.href = "mantenimientoAdministrarComponentes.asp?idEquipo=" + pId + "&ref=mantenimientoAdministrarMasters.asp";
	}
	function AbrirUploader(pId, pCd){
		var puw = new winPopUp('popupEquipo','mantenimientoEquipoFiles.asp?idEquipo=' + pId,'780','400','Archivos del Master: ' + pCd, "");
	}
	function AbrirActivacion(pId){
		var puw = new winPopUp('popupEquipo','mantenimientoActivacionMaster.asp?idEquipo=' + pId + "&tipoOperacion=A",'780','370','Activaci�n', "submitInfo()");
	}
	function AbrirActivaciones(pId){
		document.location.href = "mantenimientoAdministrarEquipos.asp?txtIdEquipo=" + pId + "&ref=mantenimientoAdministrarMasters.asp";
	}

</script>
</head>
<body onload="bodyOnLoad()">
<div id="toolBarEquipos"></div>
<form id="frmSel" name="frmSel">
	<input type="hidden" name="accion" id="accion" value="">
<div class="tableaside size100"> <!-- BUSCAR -->

	<h3> <% =GF_TRADUCIR("FILTROS") %> </h3>
  
	<div id="searchfilter" class="tableasidecontent">
        <div class="col16 reg_header_navdos"> <%=GF_Traducir("C�digo:")%> </div>
        <div class="col16"> <input type="text" size="10" value="<%=txtCdEquipo%>" id="txtCdEquipo" name="txtCdEquipo"> </div>
       
        <div class="col16 reg_header_navdos"> <%=GF_Traducir("Descripci�n:")%> </div>
        <div class="col16"> <input type="text" size="30" value="<%=txtDsEquipo%>" id="txtDsEquipo" name="txtDsEquipo"> </div>
        
        <div class="col16 reg_header_navdos"> <%=GF_Traducir("Incluir Deshabilitados:")%> </div>
        <div class="col16"> <input style="cursor:pointer;" type="checkbox" value="<%=ESTADO_BAJA%>" id="chkEstado" name="chkEstado" <%if chkEstado = 2 then Response.Write "Checked"%>> </div>
    
    	<span style="text-align:center; clear:both; float:left; width:100%"><input type="submit" value="Buscar"></span>
	</div>
</div><!-- END BUSCAR -->

<div class="col66"></div>

<table class="datagrid" width="90%" align="center">
	<thead>
	<tr>
		<th class="thiconac" width="10%" align="center">
			<% =GF_TRADUCIR("Id") %>
			<%	if orderById = "ASC" then 
					myTitle="Descendiente"
					myOrderId="DESC"
				else
					myTitle="Ascendiente"
					myOrderId="ASC"
				end if 
			%>
			<img style="cursor:pointer" title="<%=myTitle%>" onclick="setOrder('orderById', 'IDEQUIPMENT', '<%=myOrderId%>')" src="images\orderlist.png">
			<input type="hidden" id="orderById" name="orderById" value="<%=orderById%>">
		</th>
		<th class="thiconac" align="center">
			<% =GF_TRADUCIR("C�digo") %>
			<%	if orderByCodigo = "ASC" then 
					myTitle="Descendiente"
					myOrderCod="DESC"
				else
					myTitle="Ascendiente"
					myOrderCod="ASC"
				end if 
			%>
			<img style="cursor:pointer" title="<%=myTitle%>" onclick="setOrder('orderByCodigo', 'CDEQUIPMENT','<%=myOrderCod%>')" src="images\orderlist.png">
			<input type="hidden" id="orderByCodigo" name="orderByCodigo" value="<%=orderByCodigo%>">
		</th>
		<th align="center"><% =GF_TRADUCIR("Descripcion") %>
			<%	if orderByDs = "ASC" then 
					myTitle="Descendiente"
					myOrderDs="DESC"
				else
					myTitle="Ascendiente"
					myOrderDs="ASC"
				end if 
			%>
			<img style="cursor:pointer" title="<%=myTitle%>" onclick="setOrder('orderByDs', 'DSEQUIPMENT','<%=myOrderDs%>')" src="images\orderlist.png">
			<input type="hidden" id="orderByDs" name="orderByDs" value="<%=orderByDs%>">

		
		</th>
		<th class="thicon" width="5%" align="center"><%=GF_Traducir("Editar")%></th>
		<th class="thicon" width="5%" align="center"><%=GF_Traducir("Despiece")%></th>
		<th class="thicon" width="5%" align="center"><%=GF_Traducir("Adjuntos")%></th>
		<th class="thicon" width="5%" align="center"><%=GF_Traducir("Activar")%></th>
		<th class="thicon" width="5%" align="center"><%=GF_Traducir("En linea")%></th>
		<th class="thicon" width="2%" align="center">.</th>
	</tr>
	</thead>
    <tbody>	
	<%
	while ((not rs.eof)	and (reg < mostrar))
		reg = reg + 1			
		%>
		<tr>
			<td class="thiconac" align="center"> <b><% =rs("IDEQUIPMENT") %></b></td>
			<td align="center"> <b><% =rs("CDEQUIPMENT") %></b></td>
			<td><% =rs("DSEQUIPMENT") %></td>				
				<%	if (rs("CDSTATE") = ESTADO_BAJA) then	%>
						<td align="center">	. </td>
						<td align="center">	. </td>
						<td align="center">	. </td>
						<td align="center">	. </td>
						<td align="center">	. </td>
						<td align="center">
						<% if isAdminInAny then %>
								<img src="images/checkmark-16.png" style="cursor: pointer" title="Habilitar Master" onclick="AbrirElemento('H','<% =rs("IDEQUIPMENT") %>')"> 
						<% else %>
							.
						<% end if %>
						</td>

				<%	else  %>	
						<td align="center">
						<% if isAdminInAny then %>
							<img src="images/edit-16.png" style="cursor: pointer" title="Editar Master" onclick="AbrirElemento('E','<% =rs("IDEQUIPMENT") %>')"> 
						<% else %>
							.
						<% end if %>
						</td>
						<td align="center">	<img src="images/despiece-16.png" style="cursor: pointer" title="Despiece Master" onclick="AbrirDespiece('<% =rs("IDEQUIPMENT") %>')"> </td>
						<td align="center">	<img src="images/adjunto-16.png" style="cursor: pointer" title="Adjuntar Archivos" onclick="AbrirUploader('<% =rs("IDEQUIPMENT") %>','<% =rs("CDEQUIPMENT") %>')"> </td>	
						<td align="center"> <img src="images/disconnect-16.png" style="cursor: pointer" title="Activar Master" onclick="AbrirActivacion('<% =rs("IDEQUIPMENT") %>')"> </td>
						<td align="center"> <img src="images/connect-16.png" style="cursor: pointer" title="Ver activaciones de este Master" onclick="AbrirActivaciones('<% =rs("IDEQUIPMENT") %>')"> </td>
						<td align="center">
						<% if isAdminInAny then %>
							<img src="images/cross-16.png" style="cursor: pointer" title="Deshabilitar Master" onclick="AbrirElemento('D','<% =rs("IDEQUIPMENT") %>')">
						<% else %>
							.
						<% end if %>
						</td>


				<%	end if	%>					
		</tr>
		<%
		rs.MoveNext()
	wend
	%>
    </tbody>
    <tfoot>
  		<td colspan="9"><div id="paginar"></div></td>
  	</tfoot>	
  	<%
	if (reg = 0) then		
		%>
		<tr>
			<td colspan="9"><% =GF_TRADUCIR("No existen Templates registrados.") %></td>
		</tr>
		<%
	end if 
	%>	
</table>
<input type="hidden" name="myOrder" id="myOrder" value="<%=myOrder%>">	
</form>	
</body>
</html>