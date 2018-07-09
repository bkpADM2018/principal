<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientossql.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosREM.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<%
'Call controlAccesoCM("CMADMREM")
'-----------------------------------------------------------------------------------------------
Function filtrarRemitos(ByRef myWhere, idRemito, nroRemito, idAlmacen, idProveedor, idArticulo, fecha)
		
	'Filtro
	if ((idRemito <> 0) and (idRemito <> "")) then Call mkWhere(myWhere, "A.IDRemito", idRemito, "=", 1)
	if ((nroRemito <> 0) and (nroRemito <> "")) then Call mkWhere(myWhere, "A.NRORemito", nroRemito, "=", 1)
	if ((idAlmacen <> 0) and (idAlmacen <> "")) then Call mkWhere(myWhere, "A.IDALMACEN", idAlmacen, "=", 1)
	if ((idProveedor <> "0") and (idProveedor <> "")) then Call mkWhere(myWhere, "A.IDProveedor", idProveedor, "=", 1)
	if ((idArticulo <> 0) and (idArticulo <> "")) then Call mkWhere(myWhere, "C.IDARTICULO", idArticulo, "=", 1)
	if (fecha <> "") then  Call mkWhere(myWhere, "A.FECHA", fecha, "LIKE", 3)
	Call mkWhere(myWhere, "B.CDUSUARIO", session("usuario"), "=", 3)	
	filtrarRemitos = myWhere	
End Function
'-----------------------------------------------------------------------------------------------
Function obtenerListaRemitos(idRemito, nroRemito, idAlmacen, idProveedor, idArticulo, fecha, pagina, regXpag) 
	Dim strSQL, rs, myWhere, firstRecord, conn
	Call filtrarRemitos(myWhere, idRemito, nroRemito, idAlmacen, idProveedor, idArticulo, fecha)	
	strSQL = "Select A.IDREMITO, A.CDREMITO, A.FECHA from TBLREMCABECERA A inner join TBLALMACENESUSUARIO B on A.IDALMACEN=B.IDALMACEN "
	strSQL = strSQL & " inner join TBLREMDETALLE C on A.IDREMITO = C.IDREMITO "& myWhere
	strSQL = strSQL & " Group by A.IDREMITO, A.CDREMITO, A.FECHA "
	strSQL = strSQL & " order by fecha desc, A.idRemito desc"
	'Response.Write strSQL
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set obtenerListaRemitos = rs
End Function
'**********************************************************
'***	COMIENZO DE PAGINA
'**********************************************************
Dim Remitos, rsSectores, conn, strSQL, cdProveedor, descripcion, dsProveedor, idProveedor
Dim params, idEstado, idSector, fechaEmision, fechaCierre, hayBusqueda, pIdObra, flagAdmin, flagProveedor
Dim txtDE, txtME, txtAE, txtDC, txtMC, txtAC, kr, reg, rsObra, cdObra, dsObra, lineasTotales
Dim rsComentarios, cdUsuario, paginaActual, flagAuditor, myIdObra, idRemito, nroRemito
dim myTitle, tipoCompra, idArticulo, dsArticulo

idRemito = GF_PARAMETROS7("idRemito","",6)
call addParam("idRemito", idRemito, params)
nroRemito = GF_PARAMETROS7("nroRemito","",6)
call addParam("idRemito", idRemito, params)
idProveedor = GF_PARAMETROS7("idProveedor",0,6)
dsProveedor = GF_PARAMETROS7("dsProveedor","",6)
call addParam("idProveedor", idProveedor, params)
idAlmacen = GF_PARAMETROS7("idAlmacen",0,6)
call addParam("idAlmacen", idAlmacen, params)
idArticulo = GF_PARAMETROS7("idArticulo",0,6)
call addParam("idArticulo", idArticulo, params)
call getArticuloFull(idArticulo, dsArticulo, "")
txtAS = GF_PARAMETROS7("txtAnio","",6)
call addParam("txtAnio", txtAS, params)
txtMS = GF_PARAMETROS7("txtMes","",6)
call addParam("txtMes", txtMS, params)
txtDS = GF_PARAMETROS7("txtDia","",6)
call addParam("txtDia", txtDS, params)
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
	fecha = "____"
else
	fecha = txtAS
end if
if (txtMS = "") then 
	fecha = fecha & "__"
else
	fecha = fecha & txtMS
end if
if (txtDS = "") then 
	fecha = fecha & "__"
else
	fecha = fecha & txtDS
end if

GP_ConfigurarMomentos

Set Remitos = obtenerListaRemitos(idRemito, nroRemito, idAlmacen, idProveedor, idArticulo, fecha, paginaActual, mostrar)
Call setupPaginacion(Remitos, paginaActual, mostrar)
lineasTotales = Remitos.recordcount

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
<script type="text/javascript" src="scripts/controles.js"></script>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript">
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
	function abrirRemito(id) {
		var myPage, w, h;
		w=640;
		h=430;
		window.scrollTo(0,0); 
		myPage = 'almacenREM.asp?idRemito=' + id;
		var puw = new PopUpWindow('popupREM',myPage, w, h,'Remitos');				
	}
	function refreshPage(){
		document.frmSel.submit();
	}
	function seleccionarProveedor(ms) {
		var desc = ms.getSelectedItem();
		if (desc.indexOf('-') != -1) {
			var arr = desc.split('-');
			document.getElementById("idProveedor").value = arr[0];
			document.getElementById("dsProveedor").value = arr[1];
			ms.setValue(arr[1]);
		} else {
			if (desc == "") document.getElementById("idProveedor").value = 0;
			if (desc == "") document.getElementById("dsProveedor").value = "";
			ms.setValue("");
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
	function irHome() {
		location.href = "almacenIndex.asp";
	}
	function irREMNuevo() {
		location.href = "almacenSeleccionarArticulosREM.asp";
	}
	function irTDC() {
		location.href = "almacenTableroDeControl.asp";
	}
	function startMagicSearch() {		
		var msProveedor = new MagicSearch("", "companyName0", 30, 2, "comprasStreamElementos.asp?tipo=empresas");
		msProveedor.setMinChar(3);
		msProveedor.setToken(";");
		msProveedor.onBlur = seleccionarProveedor;
		msProveedor.setValue('<%=dsProveedor%>');
		var msArticulo = new MagicSearch("", "articuloItem0", 30, 4, "comprasStreamElementos.asp?tipo=articulos&linea=0&all=1");
		msArticulo.setToken(";");
		msArticulo.onBlur = seleccionarArticulo;
		msArticulo.setValue('<% =dsArticulo %>');
	}
	function bodyOnLoad() {	
		var tb = new Toolbar('toolbar', 6, 'images/almacenes/');
		tb.addButton("Home-16x16.png", "Home", "irHome()");		
		tb.addButton("REM_new-16x16.png", "Nuevo", "irREMNuevo()");
		tb.addButtonREFRESH("Recargar", "submitInfo()");		
		var swt = tb.addSwitcher("Search-16x16.png", "Buscar", "buscarOn()", "buscarOff()");		
		tb.addButton("Control_panel_folder-16x16.png", "Tablero", "irTDC()");
		tb.draw();
		<%	if (hayBusqueda) then %>
				tb.changeState(swt);			
				startMagicSearch();
		<%	end if
			if (not remitos.eof) then		%>
				var pgn = new Paginacion("paginacion");
				pgn.paginar(<% =paginaActual %>, <% =lineasTotales %>, <% =mostrar %>, 50, "almacenAdministrarREM.asp<% =params %>");
		<%	end if %>
		pngfix();
	}
	function anularRemito(idRemito, img){
		if (confirm("Esta seguro que desea devolver la totalidad de este remito?")) {
			window.scrollTo(0,0); 
			//img.src = "images/loading_small_green.gif";
			var puw = new PopUpWindow('popupREMAnulacion','almacenREMAnulacion.asp?idRemito=' + idRemito, 640, 430,'Remitos');	
		}		
	}
</script>
</head>
<body onLoad="bodyOnLoad()">
	<% call GF_TITULO2("kogge64.gif","Administración de Remitos")%>
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
								<td align="right"><% = GF_TRADUCIR("Id Remito") %>:</td>
								<td><input type="text" id="idRemito" name="idRemito" value="<% =idRemito %>" onKeyPress="return controlIngreso (this, event, 'N');"></td>
								<td align="right"><% = GF_TRADUCIR("Nro. Remito") %>:</td>
								<td><input type="text" id="nroRemito" name="nroRemito" value="<% =nroRemito %>" onKeyPress="return controlIngreso (this, event, 'N');"></td>									
							</tr>

                            <tr>
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
								<td align="right"><% = GF_TRADUCIR("Articulo") %>:</td>
								<td>
									<div id="articuloItem0"></div>																		
									<input type="hidden" id="idArticulo" name="idArticulo" value="<% =idArticulo%>">
								</td>
                            </tr>									

							<tr>
								<td align="right"><% = GF_TRADUCIR("Proveedor") %>:</td>
								<td>
									<div id="companyName0"></div>			
									<!--<input type="text" id="dsProveedor" value="<% =dsProveedor %>" size="30" onBlur="seleccionarProveedor()">-->	 
									<input type="hidden" id="idProveedor" name="idProveedor" value="<% =idProveedor %>">
									<input type="hidden" id="dsProveedor" name="dsProveedor" value="<% =dsProveedor %>">
								</td>

                                <td align="right"><% =GF_TRADUCIR("Fecha Remito") %>:</td>
								<td colspan="3">
                                    <input type="text" size="2" maxLength="2" value="<% =txtDS %>" name="txtDia" onBlur="javascript:ControlarDia(this);" onKeyPress="return controlIngreso (this, event, 'N');"> /
                                    <input type="text" size="2" maxLength="2" value="<% =txtMS %>" name="txtMes" onBlur="javascript:ControlarMes(this);" onKeyPress="return controlIngreso (this, event, 'N');"> /
                                    <input type="text" size="4" maxLength="4" value="<% =txtAS %>" name="txtAnio" onBlur="javascript:ControlarAnio(this);" onKeyPress="return controlIngreso (this, event, 'N');">
                                
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
	<input type="hidden" name="busquedaActiva" id="busquedaActiva" value="0">
	<input type="hidden" name="cdObra" id="cdObra" value="<% =pCdObra %>">	
	</form>
	<br>
	
	<table align="center" width="80%" class="reg_Header">
		<% 	if (not Remitos.eof) then %>
				<tr><td colspan="10"><div id="paginacion"></div></td></tr>
		<%	end if 	%>
			<tr class="reg_Header_nav">
				<!--<td width="3%" align="center">.</td>-->
				<td align="center"><% =GF_TRADUCIR("Id") %></td>				
				<td align="center"><% =GF_TRADUCIR("Codigo") %></td>				
				<td align="center"><% =GF_TRADUCIR("Numero") %></td>
				<td align="center"><% =GF_TRADUCIR("Almacen") %></td>
				<td align="center"><% =GF_TRADUCIR("Fecha") %></td>
				<td align="center"><% =GF_TRADUCIR("Proveedor") %></td>		
				<td align="center"><% =GF_TRADUCIR(".") %></td>
			</tr>
<%	reg=0
	if (not Remitos.eof) then
			while ((not Remitos.eof) and (reg < mostrar))
				Call initHeaderREM(Remitos("IDRemito"))
				reg = reg + 1
%>
			<tr title="<% =GF_TRADUCIR("Abrir Remito") %>" class="reg_Header_navdos <% if (REM_estado = ESTADO_BAJA) then Response.Write "reg_header_rejected" %>" onMouseOver="javascript:lightOn(this,<%=REM_estado%>)" onMouseOut="javascript:lightOff(this,<%=REM_estado%>)">						
			<!--<tr class="<%=myClass%>" onMouseOver="javascript:lightOn(this)" onMouseOut="javascript:lightOff(this)">	-->
				<%  cdAlmacen=""
					if (REM_idAlmacen > 0) then
						Set rsAlmacenes = obtenerListaAlmacenes(REM_idAlmacen) 												
						cdAlmacen = rsAlmacenes("CDAlmacen")
					end if
				%>			
				<td onclick="abrirRemito(<% =Remitos("IDRemito") %>)" align="center"><% =Remitos("IDRemito") %></td>				
				<td onclick="abrirRemito(<% =Remitos("IDRemito") %>)" align="center"><% =Remitos("cdRemito") %></td>				
				<td onclick="abrirRemito(<% =Remitos("IDRemito") %>)" align="center"><% =REM_nroRemito %></td>
				<td onclick="abrirRemito(<% =Remitos("IDRemito") %>)" align="center"><% =cdAlmacen %></td>
				<td onclick="abrirRemito(<% =Remitos("IDRemito") %>)" align="center"><% =REM_fecha %></td>
				<td onclick="abrirRemito(<% =Remitos("IDRemito") %>)"><% =REM_idProveedor & "-"& REM_dsProveedor %></td>
				<%
				if ((REM_estado = ESTADO_BAJA) or (REM_estado = ESTADO_ANULACION) or (cint(mid(REM_fecha,4,2)) <> cint(month(date)))) then 
					%>
					<td onclick="abrirRemito(<% =Remitos("IDRemito") %>)" align="center">-</td>
					<%
				else 	
					%>
					<td align="center" title="Anular Remito"><img id="ID_<%=Remitos("IDRemito")%>" src="images/almacenes/vale_reget-16x16.png" onClick="javascript:anularRemito(<%=Remitos("IDRemito")%>, this);">	</td>
					<%
				end if
				%>	
			</tr>		
	<%			Remitos.MoveNext()				
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