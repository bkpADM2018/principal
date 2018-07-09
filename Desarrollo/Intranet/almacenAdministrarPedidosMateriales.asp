<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientossql.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosPM.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<%
Call controlAccesoAL("")

'-----------------------------------------------------------------------------------------------
Function filtrarPedidosMateriales(ByRef myWhere, idPedido, idAlmacen, idObra, cdSolicitante, fechaSolicitud, fechaRequerido)
		
	'Filtro	
	if (cdSolicitante <> "") then 
		Call mkWhere(myWhere, "A.CDSOLICITANTE", cdSolicitante, "=", 3)
	else
		myWhere = "where (B.NIVEL<>'" & ALMACEN_SOLICITANTE & "' or A.CDSOLICITANTE = '" & session("Usuario") & "')"
	end if
	if ((idPedido <> 0) and (idPedido <> "")) then Call mkWhere(myWhere, "A.IDPEDIDO", idPedido, "=", 1)
	if ((idAlmacen <> 0) and (idAlmacen <> "")) then Call mkWhere(myWhere, "A.IDALMACEN", idAlmacen, "=", 1)
	if ((idObra <> 0) and (idObra <> "")) then Call mkWhere(myWhere, "A.IDOBRA", idObra, "=", 1)	
	if (fechaSolicitud <> "") then  Call mkWhere(myWhere, "A.FECHASOLICITUD", fechaSolicitud, "LIKE", 3)
	if (fechaRequerido <> "") then  Call mkWhere(myWhere, "A.FECHAREQUERIDO", fechaRequerido, "LIKE", 3)	
	filtrarPedidosMateriales = myWhere
End Function
'-----------------------------------------------------------------------------------------------
Function obtenerListaPedidosMateriales(idPedido, idAlmacen, idObra, cdSolicitante, fechaSolicitud, fechaRequerido, pagina, regXpag) 
	Dim strSQL, rs, myWhere, firstRecord, conn
	
	'Ajusto Paginacion
	Call filtrarPedidosMateriales(myWhere, idPedido, idAlmacen, idObra, cdSolicitante, fechaSolicitud, fechaRequerido)
	strSQL = "Select * from TBLPMCABECERA A inner join TBLALMACENESUSUARIO B on A.IDALMACEN=B.IDALMACEN and B.CDUSUARIO='" & session("Usuario") & "' " & myWhere & " order by IDPEDIDO desc"
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set obtenerListaPedidosMateriales = rs
End Function

'**********************************************************
'***	COMIENZO DE PAGINA
'**********************************************************
Dim pedidos, rsSectores, conn, strSQL, cdSolicitante, descripcion, dsSolicitante
Dim params, idEstado, idSector, fechaEmision, fechaCierre, hayBusqueda, pIdObra, flagAdmin, flagSolicitante
Dim txtDE, txtME, txtAE, txtDC, txtMC, txtAC, kr, reg, rsObra, cdObra, dsObra, lineasTotales
Dim rsComentarios, cdUsuario, paginaActual, flagAuditor, myIdObra, idPedido, cdPedido
dim myTitle, tipoCompra

idPedido = GF_PARAMETROS7("idPedido","",6)
call addParam("idPedido", idPedido, params)
idObra = GF_PARAMETROS7("idObra",0,6)
call addParam("idObra", idObra, params)
cdSolicitante = GF_PARAMETROS7("cdSolicitante","",6)
dsSolicitante = getUserDescription(cdSolicitante)
call addParam("cdSolicitante", cdSolicitante, params)
idAlmacen = GF_PARAMETROS7("idAlmacen",0,6)
call addParam("idAlmacen", idAlmacen, params)
txtAS = GF_PARAMETROS7("txtAnioSolicitud","",6)
call addParam("txtAnioSolicitud", txtAS, params)
txtMS = GF_PARAMETROS7("txtMesSolicitud","",6)
call addParam("txtMesSolicitud", txtMS, params)
txtDS = GF_PARAMETROS7("txtDiaSolicitud","",6)
call addParam("txtDiaSolicitud", txtDS, params)
txtAR = GF_PARAMETROS7("txtAnioRequerido","",6)
call addParam("txtAnioRequerido", txtAR, params)
txtMR = GF_PARAMETROS7("txtMesRequerido","",6)
call addParam("txtMesRequerido", txtMR, params)
txtDR = GF_PARAMETROS7("txtDiaRequerido","",6)
call addParam("txtDiaRequerido", txtDR, params)
paginaActual = GF_PARAMETROS7("numeroPagina",0,6)
if (paginaActual = 0) then paginaActual=1
mostrar = GF_PARAMETROS7("registrosPorPagina",0,6)
if (mostrar = 0) then mostrar = 10
cdUsuario = ""
hayBusqueda = false
busquedaActiva = GF_PARAMETROS7("busquedaActiva",0,6)
call addParam("busquedaActiva", busquedaActiva, params)
if busquedaActiva = 1 then hayBusqueda = true


if (txtAS = "") then 
	fechaSolicitud = "____"
else
	fechaSolicitud = txtAE
end if
if (txtMS = "") then 
	fechaSolicitud = fechaSolicitud & "__"
else
	fechaSolicitud = fechaSolicitud & txtMS
end if
if (txtDS = "") then 
	fechaSolicitud = fechaSolicitud & "__"
else
	fechaSolicitud = fechaSolicitud & txtDS
end if

if (txtAR = "") then 
	fechaRequerido = "____"
else
	fechaRequerido = txtAR
end if
if (txtMR = "") then 
	fechaRequerido = fechaRequerido & "__"
else
	fechaRequerido = fechaRequerido & txtMR
end if
if (txtDR = "") then 
	fechaRequerido = fechaRequerido & "__"
else
	fechaRequerido = fechaRequerido & txtDR
end if

GP_ConfigurarMomentos

Set pedidos = obtenerListaPedidosMateriales(idPedido, idAlmacen, idObra, cdSolicitante, fechaSolicitud, fechaRequerido, paginaActual, mostrar)
Call setupPaginacion(pedidos, paginaActual, mostrar)
lineasTotales = pedidos.recordcount
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title>Sistema de Compras - Modulo de Almacenes</title>

<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/iwin.css" type="text/css">
<link rel="stylesheet" href="css/MagicSearch.css" type="text/css">

<style type="text/css">
.titleStyle {
	font-weight: bold;
	font-size: 20px;
}

.divOculto {
	display: none;
}
</style>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/MagicSearchObj.js"></script>
<script type="text/javascript" src="scripts/paginar.js"></script>
<script type="text/javascript" src="scripts/script_fechas.js"></script>
<script type="text/javascript" src="scripts/iwin.js"></script>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript">
	function lightOn(tr) {
		tr.className = "reg_Header_navdosHL";
	}
	
	function lightOff(tr) {
		tr.className = "reg_Header_navdos";
	}
	
	function abrirPedido(id) {
		window.open("almacenValePedidoPrint.asp?idPedido=" + id, "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes,height=500,width=700",false);		
	}
	function abrirTableroPM(id) {
		window.open("almacenTableroPM.asp?idPedido=" + id, "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes,height=500,width=700",false);		
	}
	
	function seleccionarSolicitante(ms) {				
		var desc = ms.getSelectedItem();
		if (desc.indexOf('-') != -1) {
			var arr = desc.split('-');
			document.getElementById("cdSolicitante").value = arr[0];
			ms.setValue(arr[1]);
		} else {
			if (desc == "") document.getElementById("cdSolicitante").value = "";							
		}		
	}
		
	function buscarOn() {
		document.getElementById("busqueda").className = "";
		document.getElementById("busquedaActiva").value = "1";
		startMagicSearch();
	}
	
	function buscarOff() {
		document.getElementById("busqueda").className = "divOculto";
		document.getElementById("busquedaActiva").value = "0";
	}

	function submitInfo(acc) {		
		document.getElementById("frmSel").submit();
	}
	
	function volver() {
		location.href = "almacenIndex.asp";
	}
	
	function irHome() {
		location.href = "almacenIndex.asp";
	}
		
	function irAdministracion() {
		location.href = "almacenAdministracion.asp";
	}
	
	function irObras() {
		location.href = "almacenObras.asp";
	}

	function irPMNuevo() {
		location.href = 'almacenValesTitulo.asp?TC=0&cdVale=PM';
	}

	function irTDC() {
		location.href = "almacenTableroDeControl.asp";
	}
		
	function startMagicSearch() {		
		
		var msSolicitante = new MagicSearch("", "divSolicitante", 30, 2, "comprasStreamElementos.asp?tipo=personas");
		msSolicitante.setToken(";");
		msSolicitante.onBlur = seleccionarSolicitante;
		msSolicitante.setValue('<%=dsSolicitante%>');
	}
	function irPM(){
		location.href = "almacenPedidoMaterial.asp?TC=0&cdVale=PM";		
	}
	function bodyOnLoad() {	
		var tb = new Toolbar('toolbar', 6, "images/almacenes/");
		tb.addButton("Home-16x16.png", "Home", "irHome()");		
		//tb.addButton("PM_new-16x16.png", "Nuevo", "irPMNuevo()");
		tb.addButton("PM_new-16x16.png", "Nuevo", "irPM()");		
		tb.addButtonREFRESH("Recargar", "submitInfo()");		
		var swt = tb.addSwitcher("Search-16x16.png", "Buscar", "buscarOn()", "buscarOff()");				
		tb.addButton("Control_panel_folder-16x16.png", "Tablero", "irTDC()");
		tb.draw();
		<%	if (hayBusqueda) then %>
				tb.changeState(swt);			
		<%	End if 
			if (not pedidos.eof) then		%>								
				var pgn = new Paginacion("paginacion");							
				pgn.paginar(<% =paginaActual %>, <% =lineasTotales %>, <% =mostrar %>, 50, "almacenAdministrarPedidosMateriales.asp<% =params %>");
		<%	end if %>
		pngfix();
	}
</script>
</head>
<body onLoad="bodyOnLoad()">
	<% call GF_TITULO2("kogge64.gif","Pedidos de Materiales") %>		
	<div id="toolbar"></div>
	<br>
	<form name="frmSel" id="frmSel">
	<div id="busqueda" class="divOculto">
	<table width="90%" cellspacing="0" cellpadding="0" align="center" border="0">
       <input type="hidden" name="accion" id="accion" value="">
       <tr>
           <td width="8"><img src="images/marco_r1_c1.gif"></td>
           <td width="25%"><img src="images/marco_r1_c2.gif" width="100%" height="8"></td>
           <td width="8"><img src="images/marco_r1_c3.gif"></td>
           <td width="75%"><td>
           <td></td>
       </tr>
       <tr>
           <td width="8"><img src="images/marco_r2_c1.gif"></td>
           <td align="center" valign="center"><font class="big" color="#517b4a"><% =GF_TRADUCIR("Búsqueda") %></font></td>
           <td width="8"><img src="images/marco_r2_c3.gif"></td>
           <td></td>
           <td></td>
       </tr>
       <tr>
           <td><img src="images/marco_r2_c1.gif" height="8"  width="8"></td>
           <td></td>
           <td><img src="images/marco_c_s_d.gif" height="8" width="8"></td>
           <td><img src="images/marco_r1_c2.gif" width="100%" height="8"></td>
           <td width="8"><img src="images/marco_r1_c3.gif"></td>
       </tr>
       <tr>
           <td height="100%"><img src="images/marco_r2_c1.gif" height="100%" width="8"></td>
           <td colspan="3">
                     <table width="100%" align="center" border="0">
							<tr>
								<td align="right"><% = GF_TRADUCIR("Cód. Pedido") %>:</td>
								<td><input type="text" id="idPedido" name="idPedido" value="<% =idPedido %>"></td>									
							</tr>

                            <tr>
								<td align="right"><% = GF_TRADUCIR("Obra") %>:</td>
									<%
									Set rsObras = obtenerListaObras("", "", "","", "")
									%>						
									<td>
									<select id="idObra" name="idObra">
									<option value="0">- <% =GF_TRADUCIR("Seleccione") %> -
									<%	while (not rsObras.eof)	%>							
										<option value="<% =rsObras("IDOBRA") %>" <% if (rsObras("IDOBRA") = idObra) then response.write "selected='true'" %>><% =GF_TRADUCIR(rsObras("CDOBRA")) %> - <% =GF_TRADUCIR(rsObras("DSOBRA")) %>
									<%		
										rsObras.MoveNext()
										wend 	%>		
									</select>																				
								</td>
								<td align="right"><% =GF_TRADUCIR("Almacen") %>:</td>
								<%Set rsAlmacenes = obtenerListaAlmacenesUsuario()%>   
                                <td>                                
									<select id="idAlmacen" name="idAlmacen">
										<option value="0">- <% =GF_TRADUCIR("Seleccione") %> -
										<%	
										while (not rsAlmacenes.eof)	%>
											<option value="<% =rsAlmacenes("IDALMACEN") %>" <% if (rsAlmacenes("IDALMACEN") = idAlmacen) then response.write "selected='true'" %>><% =GF_TRADUCIR(rsAlmacenes("CDALMACEN")) %> - <% =GF_TRADUCIR(rsAlmacenes("DSALMACEN")) %>
										<%		
											rsAlmacenes.MoveNext()
										wend 	
										%>		
									</select>		
                                </td>	
                            </tr>									

							<tr>
								<td align="right"><% = GF_TRADUCIR("Solicitante") %>:</td>
								<td>
									<div id="divSolicitante"></div>			
									<!--<input type="text" id="dsSolicitante" value="<% =dsSolicitante %>" size="30" onBlur="seleccionarSolicitante()">-->	 
									<input type="hidden" id="cdSolicitante" name="cdSolicitante" value="<% =cdSolicitante %>">
								</td>

                                <td align="right"><% =GF_TRADUCIR("Solicitado") %>:</td>
								<td colspan="3">
                                    <input type="text" size="1" maxLength="2" value="<% =txtDS %>" name="txtDiaSolicitud" onBlur="javascript:ControlarDia(this);"> /
                                    <input type="text" size="1" maxLength="2" value="<% =txtMS %>" name="txtMesSolicitud" onBlur="javascript:ControlarMes(this);"> /
                                    <input type="text" size="3" maxLength="4" value="<% =txtAS %>" name="txtAnioSolicitud" onBlur="javascript:ControlarAnio(this);">
                                
                               &nbsp;&nbsp;&nbsp;&nbsp;<% =GF_TRADUCIR("Requerido") %>:

                                    <input type="text" size="1" maxLength="2" value="<% =txtDR %>" name="txtDiaRequerido" onBlur="javascript:ControlarDia(this);"> /
                                    <input type="text" size="1" maxLength="2" value="<% =txtMR %>" name="txtMesRequerido" onBlur="javascript:ControlarMes(this);"> /
                                    <input type="text" size="3" maxLength="4" value="<% =txtAR %>" name="txtAnioRequerido" onBlur="javascript:ControlarAnio(this);">
								</td>                                
                            </tr>
                            <tr>
								<td colspan="4" align="center"><input type="button" value="Buscar..." id=submit1 name=submit1 onclick='submitInfo();'></td>						
                            </tr>								
                     </table>
	           </td>
	           <td height="100%"><img src="images/marco_r2_c3.gif" width="8" height="100%"></td>
	       </tr>
	       <tr>
	           <td width="8"><img src="images/marco_r3_c1.gif"></td>
	           <td width="100%" align=center colspan="3"><img src="images/marco_r3_c2.gif" width="100%" height="8"></td>
	           <td width="8"><img src="images/marco_r3_c3.gif"></td>
	       </tr>
	</table>
	  
	</div>
	<input type="hidden" name="busquedaActiva" id="busquedaActiva" value="<%=hayBusqueda%>">
	<input type="hidden" name="cdObra" id="cdObra" value="<% =pCdObra %>">
	<input type="hidden" value="<%=tipoCompra%>" name="tipoCompra" id="tipoCompra">
	</form>
	<br>
	<table align="center" width="90%" class="reg_Header">
			<% 	if (not pedidos.eof) then %>
			<tr><td colspan="10"><div id="paginacion"></div></td></tr>
		<%	end if 	%>
			<tr class="reg_Header_nav">
				<td  style="text-align: center" colspan="2"><% =GF_TRADUCIR("Pedido") %></td>				
				<td  style="text-align: center"><% =GF_TRADUCIR("Ptda.Presup o Sector") %></td>
				<td  style="text-align: center"><% =GF_TRADUCIR("Almacen") %></td>
				<td  style="text-align: center"><% =GF_TRADUCIR("Solicitud") %></td>
				<td  style="text-align: center"><% =GF_TRADUCIR("Requerido") %></td>
				<td  style="text-align: center"><% =GF_TRADUCIR("Solicitante") %></td>
			</tr>		
<%	reg=0

	if (not pedidos.eof) then			
			while ((not pedidos.eof) and (reg < mostrar))
				Call initHeaderPMDB(pedidos("IDPEDIDO"))	
				reg=reg+1
%>
			<tr class="reg_Header_navdos" onMouseOver="javascript:lightOn(this)" onMouseOut="javascript:lightOff(this)">			
				<%	if (pm_idObra > 0) then
						call loadDatosObra(PM_idObra, PM_cdObra, "", 0, "", 0, "", 0, "", "", "", "", "")
						PM_cdObra = PM_cdObra & " (" & PM_idBudgetArea & "-" & PM_idBudgetDetalle & ")"
					else
						PM_cdObra=""
						Set rsSector = obtenerSectores(PM_idSector)
						if (not rsSector.eof) then PM_cdObra = rsSector("DSSECTOR")
					end if		
					cdAlmacen=""
					if (pm_idAlmacen > 0) then
						Set rsAlmacenes = obtenerListaAlmacenes(pm_idAlmacen) 												
						if not rsAlmacenes.eof then 
							cdAlmacen = rsAlmacenes("CDAlmacen")
						else
							cdAlmacen = ""
						end if							
					end if
					
					%>
				<td align="center" onclick="javascript:abrirTableroPM(<% =pm_idPedido %>)"><% =PM_idPedido %></td>
				<td align="center"><img src="images/almacenes/PM-16x16.png" onClick="javascript:abrirPedido(<% =pm_idPedido %>)"></td>				
				<td align="center" onclick="javascript:abrirTableroPM(<% =pm_idPedido %>)"><% =PM_cdObra %></td>				
				<td align="center" onclick="javascript:abrirTableroPM(<% =pm_idPedido %>)"><% =cdAlmacen %></td>
				<td align="center" onclick="javascript:abrirTableroPM(<% =pm_idPedido %>)"><% =PM_fechaSolicitud %></td>
				<td align="center" onclick="javascript:abrirTableroPM(<% =pm_idPedido %>)"><% =PM_fechaRequerido %></td>
				<td align="center" onclick="javascript:abrirTableroPM(<% =pm_idPedido %>)"><% =PM_dsSolicitante %></td>				
			</tr>		
	<%			pedidos.MoveNext()				
			wend 
	end if
	if (reg = 0) then
%>
			<tr class="TDNOHAY"><td colSpan="10"><% =GF_TRADUCIR("No hay informacion disponible en estos momentos") %></td></tr>		
<%  end if %>				
		</table>	
</body>
</html>
<%
'******************************************************************************************
	Function addParam(p_strKey,p_strValue,ByRef p_strParam)
           if (not isEmpty(p_strValue)) then
              if (isEmpty(p_strParam)) then
                 p_strParam = "?"
              else
                 p_strParam = p_strParam & "&"
              end if
              p_strParam = p_strParam & p_strKey & "=" & p_strValue
           end if
	End Function
'******************************************************************************************
	Function createOption(id, text, param)
	
		Dim sel
		
		sel=""
		if (CLng(id) = CLng(param)) then sel ="selected"		
		createOption = "<option value='" & id & "' " & sel & ">" & text
		
	End Function
'******************************************************************************************
	Function setFunc(func)
		if (not flagAuditor) then 
			setFunc = func
		else
			setFunc = ""
		end if
	End Function
	
%>