<!--include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientossql.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosPM.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<%
'-----------------------------------------------------------------------------------------------
Function filtrarVales(ByRef myWhere, nrVale, cdVale, idAlmacen, cdResponsable, idArticulo, fecha, almAutorizadas)
		
	'Filtro	
	Call mkWhere(myWhere, "VC.IDALMACEN", "(" & almAutorizadas & ")", "IN", 1)
	if (nrVale <> "") then Call mkWhere(myWhere, "nrVale", nrVale, "LIKE", 3)
	if (cdVale <> "") then 	Call mkWhere(myWhere, "cdVale", ucase(cdVale), "=", 0)		
	if ((idAlmacen <> 0) and (idAlmacen <> "")) then Call mkWhere(myWhere, "VC.IDALMACEN", idAlmacen, "=", 1)
	if (cdResponsable <> "") then Call mkWhere(myWhere, "VC.CDSOLICITANTE", cdResponsable, "=", 3)
	if (fecha <> "") then  Call mkWhere(myWhere, "VC.FECHA", fecha, "LIKE", 3)
	if (idArticulo <> 0) then Call mkWhere(myWhere, "VD.IDARTICULO", idArticulo, "=", 1)
	filtrarVales = myWhere	
End Function
'-----------------------------------------------------------------------------------------------
Function obtenerListaVales(nrVale, idAlmacen, cdResponsable, idArticulo, fecha, pagina, regXpag, almAutorizadas) 
	Dim strSQL, rs, myWhere, firstRecord, conn
	
	Call filtrarVales(myWhere, nrVale, cdVale, idAlmacen, cdResponsable, idArticulo, fecha, almAutorizadas)	
	strSQL = "Select VC.IDVALE from TBLVALESCABECERA VC "
	if idArticulo <> 0 then 
		strSQL = strSQL & " INNER JOIN TBLVALESDETALLE VD ON VC.IDVALE=VD.IDVALE "
		myWhere = myWhere & " GROUP BY VC.IDVALE"
	end if	
	strSQL = strSQL & myWhere
    strSQL = strSQL & " order by VC.IDVALE desc"    
    call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set obtenerListaVales = rs
End Function
'**********************************************************
'***	COMIENZO DE PAGINA
'**********************************************************
Dim vales, rsSectores, conn, strSQL, cdResponsable, descripcion, dsResponsable
Dim params, idEstado, idSector, fecha, hayBusqueda, pIdObra, flagAdmin, flagSolicitante, nrVale
Dim txtDE, txtME, txtAE, txtDC, txtMC, txtAC, kr, reg, rsObra, cdObra, dsObra, lineasTotales
Dim rsComentarios, cdUsuario, paginaActual, flagAuditor, myIdObra, idPedido, cdPedido
dim myTitle, tipoCompra, idArticulo, dsArticulo

Call controlAccesoAL("")

'Se obtienen las almacenes a las que el usuario es pañolero o administrador o auditor.
Set rsAlmacenes = obtenerListaAlmacenesUsuario()
if (rsAlmacenes.eof) then
    response.redirect "comprasAccesoDenegado.asp"
else
    while not rsAlmacenes.eof
	    almacenesAutorizadas = rsAlmacenes("IDALMACEN") & "," & almacenesAutorizadas
	    rsAlmacenes.MoveNext()
    wend	    
    almacenesAutorizadas = left(almacenesAutorizadas, len(almacenesAutorizadas)-1)
    rsAlmacenes.MoveFirst()
end if
almacenesAutorizadas = Replace(almacenesAutorizadas,Chr(9), ", ")

GP_ConfigurarMomentos

cdVale = ucase(GF_PARAMETROS7("cdVale","",6))
call addParam("cdVale", cdVale, params)
nrVale = GF_PARAMETROS7("nrVale","",6)
call addParam("nrVale", nrVale, params)
cdResponsable = GF_PARAMETROS7("cdResponsable","",6)
dsResponsable = getUserDescription(cdResponsable)
call addParam("cdResponsable", cdResponsable, params)
idArticulo = GF_PARAMETROS7("idArticulo",0,6)
call addParam("idArticulo", idArticulo, params)
call getArticuloFull(idArticulo, dsArticulo, "")
idAlmacen = GF_PARAMETROS7("idAlmacen",0,6)
call addParam("idAlmacen", idAlmacen, params)
hayBusqueda = false
busquedaActiva = GF_PARAMETROS7("busquedaActiva",0,6)
call addParam("busquedaActiva", busquedaActiva, params)
if busquedaActiva = 1 then hayBusqueda = true
txtAS = GF_PARAMETROS7("txtAnioSolicitud","",6)
if (not hayBusqueda) then txtAS = left(session("MmtoDato"), 4)
call addParam("txtAnioSolicitud", txtAS, params)
txtMS = GF_PARAMETROS7("txtMesSolicitud","",6)
if (not hayBusqueda) then txtMS = Mid(session("MmtoDato"), 5, 2)
call addParam("txtMesSolicitud", txtMS, params)
txtDS = GF_PARAMETROS7("txtDiaSolicitud","",6)
call addParam("txtDiaSolicitud", txtDS, params)
paginaActual = GF_PARAMETROS7("numeroPagina",0,6)
if (paginaActual = 0) then paginaActual=1
mostrar = GF_PARAMETROS7("registrosPorPagina",0,6)
if (mostrar = 0) then mostrar = 40
cdUsuario = ""

if (txtAS = "") then
	fechaSolicitud = "____"
else
	fechaSolicitud = txtAS
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

Set vales = obtenerListaVales(nrVale, idAlmacen, cdResponsable, idArticulo, fechaSolicitud, paginaActual, mostrar, almacenesAutorizadas)
Call setupPaginacion(vales, paginaActual, mostrar)
lineasTotales = vales.recordcount
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
	var ch = new channel();	
	function lightOn(tr, estado) {
		if (estado == <%=ESTADO_BAJA%>) {
			tr.className = "reg_Header_navdosHL reg_header_rejected";
		}
		else{
			tr.className = "reg_Header_navdosHL";
		}
	}
	
	function lightOff(tr, estado) {
		if (estado == <%=ESTADO_BAJA%>) {
			tr.className = "reg_Header_navdos reg_header_rejected";
		}
		else{
			tr.className = "reg_Header_navdos";
		}
	}	
	
	function abrirVale(id) {
		window.open("almacenValePedidoPrint.asp?idVale=" + id, "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes,height=500,width=700",false);		
	}
	function editarVale(id) {
		window.open("almacenValesTitulo.asp?idVale=" + id, "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes,height=500,width=700",false);		
	}

	function abrirTableroPM(id) {
		window.open("almacenTableroPM.asp?idPedido=" + id, "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes,height=500,width=700",false);		
	}
	
	function seleccionarSolicitante(ms) {				
		var desc = ms.getSelectedItem();
		if (desc.indexOf('-') != -1) {
			var arr = desc.split('-');
			document.getElementById("cdResponsable").value = arr[0];
			ms.setValue(arr[1]);
		} else {
			if (desc == "") document.getElementById("cdResponsable").value = "";							
		}		
	}
	
	function seleccionarArticulo(ms) {
		var desc = ms.getSelectedItem();
		if (desc.indexOf('|') != -1) {
			var arr = desc.split('|');
			document.getElementById("idArticulo").value = arr[0];
			ms.setValue(arr[1]);
		} else {
			if (desc == "") document.getElementById("idArticulo").value = "";							
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

	function irTDC() {
		location.href = "almacenTableroDeControl.asp";
	}
		
	function startMagicSearch() {		
		var msSolicitante = new MagicSearch("", "divResponsable", 30, 2, "comprasStreamElementos.asp?tipo=personas");
		msSolicitante.setToken(";");
		msSolicitante.onBlur = seleccionarSolicitante;
		msSolicitante.setValue('<%=dsResponsable%>');
		var msArticulo = new MagicSearch("", "articuloItem0", 30, 4, "comprasStreamElementos.asp?tipo=articulos&linea=0&all=1");
		msArticulo.setToken(";");
		msArticulo.onBlur = seleccionarArticulo;
		msArticulo.setValue('<% =dsArticulo %>');			
	}

	function bodyOnLoad() {	
		var tb = new Toolbar('toolbar', 6, "images/almacenes/");
		tb.addButton("Home-16x16.png", "Home", "irHome()");		
		tb.addButtonREFRESH("Recargar", "submitInfo()");		
		var swt = tb.addSwitcher("Search-16x16.png", "Buscar", "buscarOn()", "buscarOff()");				
		tb.addButton("Control_panel_folder-16x16.png", "Tablero", "irTDC()");
		tb.draw();
		<%	if (hayBusqueda) then %>
				tb.changeState(swt);	
		<%	End if 
			if (not VALES.eof) then		%>								
				var pgn = new Paginacion("paginacion");							
				pgn.paginar(<% =paginaActual %>, <% =lineasTotales %>, <% =mostrar %>, 200, "almacenAdministrarVales.asp<% =params %>");
		<%	end if %>
		pngfix();
	}
	function anularVale(idAlmacen, idVale, cdVale, idPM, img){
		if (confirm("Esta seguro que desea anular el vale?")) {
			img.src = "images/loading_small_green.gif"
			ch.bind("almacenAnularValeAjax.asp?idAlmacen=" + idAlmacen + "&idVale=" + idVale + "&cdVale=" + cdVale + "&idPM=" + idPM, "anularValeCallback('" + img.id + "')");
			ch.send();			
		}		
	}
	function anularValeCallback(pId){
		var txt = new String();
		txt = ch.response();
		txt = txt.replace(/<BR>/g,'\n'); 
		if (ch.response() != "") alert(txt);
		document.getElementById(pId).src = "images/almacenes/accept-16x16.png";
		submitInfo();
	}
</script>
</head>
<body onLoad="bodyOnLoad()">
	<% call GF_TITULO2("kogge64.gif","Administración - Consulta de Vales") %>		
	<div id="toolbar"></div>
	<br>
	<form name="frmSel" id="frmSel">
	<div id="busqueda" class="divOculto">
	<table width="70%" cellspacing="0" cellpadding="0" align="center" border="0">
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
								<td align="right"><% = GF_TRADUCIR("Nro. Vale") %>:</td>
                                <td>                                
									<input size="15" type="text" id="nrVale" name="nrVale" value="<%=nrVale%>">
										
                                </td>								
								<td align="right"><% =GF_TRADUCIR("Almacen") %>:</td>
								<td>                                
									<select id="idAlmacen" name="idAlmacen">
											<option value="0" <% if (idAlmacen=0) then response.write "selected='true'" %>><% =GF_TRADUCIR("Todos") %>
											<%	
											while (not rsAlmacenes.eof)	
											%>
												<option value="<% =rsAlmacenes("IDALMACEN") %>" <% if (rsAlmacenes("IDALMACEN") = idAlmacen) then response.write "selected='true'" %>><% =GF_TRADUCIR(rsAlmacenes("CDALMACEN")) %> - <% =GF_TRADUCIR(rsAlmacenes("DSALMACEN")) %>
											<%		
												rsAlmacenes.MoveNext()
											wend 	
											%>		
									</select>		
                                </td>	
                            </tr>									

							<tr>
								<td align="right"><% = GF_TRADUCIR("Cód. Vale") %>:</td>
                                <td>                                
									<input size="3" type="text" id="cdVale" name="cdVale" value="<%=cdVale%>" maxlength=3>
										
                                </td>	
                                <td align="right"><% =GF_TRADUCIR("Fecha") %>:</td>
								<td colspan="3">
                                    <input type="text" size="2" maxLength="2" value="<% =txtDS %>" name="txtDiaSolicitud" onBlur="javascript:ControlarDia(this);"> /
                                    <input type="text" size="2" maxLength="2" value="<% =txtMS %>" name="txtMesSolicitud" onBlur="javascript:ControlarMes(this);"> /
                                    <input type="text" size="4" maxLength="4" value="<% =txtAS %>" name="txtAnioSolicitud" onBlur="javascript:ControlarAnio(this);">
                                
								</td>                                
                            </tr>
                            <tr>
								<td align="right"><% = GF_TRADUCIR("Articulo") %>:</td>
								<td>
									<div id="articuloItem0"></div>																		
									<input type="hidden" id="idArticulo" name="idArticulo" value="<% =idArticulo%>">
								</td>
								<td align="right"><% = GF_TRADUCIR("Solicitante") %>:</td>
								<td>
									<div id="divResponsable"></div>			
									<input type="hidden" id="cdResponsable" name="cdResponsable" value="<% =cdResponsable %>">
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
		<% 	if (not vales.eof) then %>
				<tr><td colspan="10"><div id="paginacion"></div></td></tr>
		<%	end if 	%>
			<tr class="reg_Header_nav">
				<td align="center"><% =GF_TRADUCIR("Nro.") %></td>
				<td align="center"><% =GF_TRADUCIR("Vale") %></td>		
				<td align="center"><% =GF_TRADUCIR("Pedido") %></td>		
				<td align="center"><% =GF_TRADUCIR("Almacen") %></td>
				<td align="center"><% =GF_TRADUCIR("Responsable") %></td>				
				<td align="center">.</td>
				<td align="center">.</td>
			</tr>		
<%	reg=0
		while ((not vales.eof) and (reg < mostrar))
			Call initHeaderValeDB(vales("IDVALE"))	
			reg = reg + 1
			%>
			<tr class="reg_Header_navdos <% if (VS_estado = ESTADO_BAJA) then Response.Write "reg_header_rejected" %>" onMouseOver="javascript:lightOn(this,<%=VS_estado%>)" onMouseOut="javascript:lightOff(this,<%=VS_estado%>)">						
					<td align="center" onClick="javascript:abrirVale(<% =VS_idVALE %>)"><% =VS_nrVale %></td>					
					
					<td align="left" onClick="javascript:abrirVale(<% =VS_idVALE %>)"><% =VS_cdVALE & " - " & getLeyendaCdVale(VS_cdVALE) %></td>
					<td align="center" onClick="javascript:abrirVale(<% =VS_idVALE %>)">
						<%
						if VS_cdVALE = CODIGO_VS_ENTRADA then 
							Response.write "-"
						else
							Response.write VS_PartidaPendiente 
						end if	
						%>		
					</td>
					<td align="center" onClick="javascript:abrirVale(<% =VS_idVALE %>)">
						<% 
						rsAlmacen = obtenerListaAlmacenes(VS_idAlmacen)
						Response.Write rsAlmacen("DSALMACEN")
						%>		
					</td>						
					<td align="center" onClick="javascript:abrirVale(<% =VS_IDVale %>)"><% =VS_DSSOLICITANTE %></td>
					<% if (isAuditorAL(VS_idAlmacen) or isAdminAL(VS_idAlmacen)) then %>
					    <td align="center"><a href="almacenValesFirma.asp?idvale=<% =VS_idVALE %>&tipo=<% =VS_cdVALE %>" target="_blank"><img src="images/almacenes/Invoice-16x16.png" title="Ver valuación del vale""></a></td>
					<% else %>
					    <td align="center">-</td>
					<% end if %>
					<%
					if ((VS_estado = ESTADO_BAJA) or (VS_estado = ESTADO_ANULACION) or (cint(mid(VS_FechaSolicitud,4,2)) <> cint(month(date)))) then 
						%>
						<td align="center">-</td>
						<%
					else 	
						%>
						<td align="center" title="Anular vale"><img id="ID_<%=VS_IDVALE%>" src="images/almacenes/vale_reget-16x16.png" onClick="javascript:anularVale(<%=VS_idAlmacen %>, <%=VS_IDVALE%>, '<%=VS_CDVALE%>', <%=VS_PartidaPendiente%>, this);">	</td>
						<%
					end if
					%>	
			</tr>
			<%			
			vales.MoveNext()				
		wend 
	if (reg = 0) then
%>
			<tr class="TDNOHAY"><td colSpan="11"><% =GF_TRADUCIR("No hay informacion disponible en estos momentos") %></td></tr>		
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