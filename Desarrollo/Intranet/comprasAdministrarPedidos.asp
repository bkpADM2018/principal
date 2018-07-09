<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientossql.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->
<!--#include file="Includes/procedimientosTitulos.asp"-->
<%
Const ESTADO_TODOS=-1
'constantes de busqueda
Const BUSQUEDA_CDPEDIDO  = 1
Const BUSQUEDA_PART_PRES = 2
Const BUSQUEDA_TITULO	 = 3
Const BUSQUEDA_F_CIERRE  = 4
Const BUSQUEDA_REF		 = 5

Call comprasControlAccesoCM(RES_CC)
'-----------------------------------------------------------------------------------------------
'funcion que arma el order by para la consulta sql a la bd
'recibe la variable donde se guarda la consulta, el campo en una constante numerica
' y el tipo (asc, desc)
Function getSqlOrder(myOrder, pCampoOrder, pTipoOrder)

	if (myOrder	= "") then
		myOrder = "ORDER BY "
	else
		myOrder = myOrder & ", "
	end if

	Select case (pCampoOrder)
		case BUSQUEDA_CDPEDIDO:
			'en el caso del cdpedido debe realizarse por separado para lograr el orden deseado
			'primero se ordena por la abreviatura de la division
			myOrder = myOrder & " SUBSTRING(pct.cdpedido, 0, Len(pct.cdpedido) - 6 ) " & pTipoOrder
			'en segundo lugar por los ultimos dos digitos que indican el año fiscal
			myOrder = myOrder & " , SUBSTRING(pct.cdpedido, 8, 2 ) " & pTipoOrder
			'por ultimo se ordena por los tres digitos que indican el numero de pedido en el cd
			myOrder = myOrder & " , SUBSTRING(pct.cdpedido, (LEN(pct.cdpedido) - 5), 3) " & pTipoOrder
		case BUSQUEDA_PART_PRES:
			myOrder = myOrder & " obras.cdobra " & pTipoOrder
		case BUSQUEDA_TITULO:
			myOrder = myOrder & " UPPER(pct.titulo) " & pTipoOrder
		case BUSQUEDA_F_CIERRE:
			myOrder = myOrder & " pct.fechacierre " & pTipoOrder
		case BUSQUEDA_REF:
			'este caso tambien es complejo, primero se ordena por numero de contrato, tengo o no y luego por numero de pic
			'de esta manera listara primero los pedidos con contrato y luego los que solo tienen pics
			myOrder = myOrder & " ctc.CDCONTRATO " & pTipoOrder
			myOrder = myOrder & " , coti.idcotizacion " & pTipoOrder
		case else:
			myOrder = myOrder & " pct.estado ASC, pct.cdpedido ASC"
	end Select

	getSqlOrder = true
End Function
'-----------------------------------------------------------------------------------------------
'**********************************************************
'***	COMIENZO DE PAGINA
'**********************************************************
Dim pedidos, rsDivisiones, conn, strSQL, cdSolicitante, titulo, dsSolicitante
Dim params, idEstado, idDivision, fechaEmision, fechaCierre, hayBusqueda, pIdObra
Dim txtDE, txtME, txtAE, txtDC, txtMC, txtAC, kr, reg, rsObra, cdObra, dsObra
Dim rsComentarios, cdUsuario, paginaActual, flagAuditor, myIdObra, idPedido, cdPedido
dim myTitle, rsCtz, pctImage
Dim sql_Order,ds_Solicitante,cd_olicitante,pkm,totalRows,cantProv,sp_parameter

idPedido = GF_PARAMETROS7("idPedido","",6)
call addParam("idPedido", idPedido, params)
cdPedido = UCase(GF_PARAMETROS7("cdPedido","",6))
call addParam("cdPedido", cdPedido, params)
idDivision = GF_PARAMETROS7("idDivision",0,6)
call addParam("idDivision", idDivision, params)
idEstado = GF_PARAMETROS7("idEstado",0,6)
call addParam("idEstado", idEstado, params)
txtAE = GF_PARAMETROS7("txtAnioEmision","",6)
call addParam("txtAnioEmision", txtAE, params)
txtME = GF_PARAMETROS7("txtMesEmision","",6)
call addParam("txtMesEmision", txtME, params)
txtDE = GF_PARAMETROS7("txtDiaEmision","",6)
call addParam("txtDiaEmision", txtDE, params)
txtAC = GF_PARAMETROS7("txtAnioCierre","",6)
call addParam("txtAnioCierre", txtAC, params)
txtMC = GF_PARAMETROS7("txtMesCierre","",6)
call addParam("txtMesCierre", txtMC, params)
txtDC = GF_PARAMETROS7("txtDiaCierre","",6)
call addParam("txtDiaCierre", txtDC, params)
titulo = UCase(GF_PARAMETROS7("titulo","",6))
call addParam("titulo", titulo, params)
cantProv = Trim(GF_PARAMETROS7("cantProv","",6))
call addParam("cantProv", cantProv, params)
hayBusqueda = GF_PARAMETROS7("busquedaActiva",0,6)
call addParam("busquedaActiva", hayBusqueda, params)
paginaActual = GF_PARAMETROS7("numeroPagina",0,6)
if (paginaActual = 0) then paginaActual=1
mostrar = GF_PARAMETROS7("registrosPorPagina",0,6)
if (mostrar = 0) then mostrar = 10
cdUsuario = ""
'EAB
cdSolicitante = GF_PARAMETROS7("cdSolicitante","",6)
call addParam("cdSolicitante", cdSolicitante, params)
sql_Campo_Order = GF_PARAMETROS7("sqlCampoOrder",0,6)
call addParam("sqlCampoOrder", sql_Campo_Order, params)
sql_Tipo_Order = GF_PARAMETROS7("sqlTipoOrder","",6)
call addParam("sqlTipoOrder", sql_Tipo_Order, params)

idObra = GF_PARAMETROS7("idObra","",6)
call addParam("idObra", idObra, params)
cdObra = GF_PARAMETROS7("divObra_text","",6)

if ((txtAE <> "") or (txtME <> "") or (txtDE <> "")) then
	if (txtAE = "") then 
		fechaEmision = "____"
	else
		fechaEmision = txtAE
	end if
	if (txtME = "") then 
		fechaEmision = fechaEmision & "__"
	else
		fechaEmision = fechaEmision & txtME
	end if
	if (txtDE = "") then 
		fechaEmision = fechaEmision & "__"
	else
		fechaEmision = fechaEmision & txtDE
	end if
end if
if ((txtAC <> "") or (txtMC <> "") or (txtDC <> "")) then
	if (txtAC = "") then 
		fechaCierre = "____"
	else
		fechaCierre = txtAC
	end if
	if (txtMC = "") then 
		fechaCierre = fechaCierre & "__"
	else
		fechaCierre = fechaCierre & txtMC
	end if
	if (txtDC = "") then 
		fechaCierre = fechaCierre & "__"
	else
		fechaCierre = fechaCierre & txtDC
	end if
end if
GP_ConfigurarMomentos
myIdObra = 0
if (idObra <> "") then 	
	Set rsObra = obtenerListaObras("", idObra, "", "","")
	if (not rsObra.eof) then
		myIdObra = rsObra("IDOBRA")
	end if
end if
myDivisiones = getListaCargosAdmin()
Call getSqlOrder(myOrder, sql_Campo_Order, sql_Tipo_Order)
sp_parameter = cdPedido &"||"& myIdObra &"||"& idDivision &"||"& idEstado &"||"& fechaEmision &"||"& fechaCierre &"||"& titulo &"||"& cdSolicitante &"||"&_
			   session("Usuario") &"||"& myDivisiones &"||"& cantProv &"||"& myOrder &"||"& paginaActual &"||"& mostrar &"$$totalRegistros"
Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, pedidos, "TBLPCTCABECERA_GET_BY_PARAMETERS", sp_parameter)
totalRegistros = sp_ret("totalRegistros")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title>Sistema de Compras</title>

<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
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
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript">
	function lightOn(tr) {
		tr.className = "reg_Header_navdosHL";
	}
	
	function lightOff(tr) {
		tr.className = "reg_Header_navdos";
	}
	
	function abrirFicha(id, tab) {
		window.open("comprasFichaPedidoCotizacion.asp?idPedido=" + id + "&tab=" + tab, "_blank", "resizable=yes,location=no,scrollbars=yes,menubar=no,statusbar=no,height=500,width=500",false);
	}
	
	function abrirCotizacion(id) {
		window.open("comprasPICPrint.asp?idCotizacionElegida=" + id, "_blank", "resizable=yes,location=no,menubar=no,statusbar=no,height=500,width=700",false);
	}
	
	function abrirPedido(id) {
		window.open("comprasPedidoCotizacion.asp?idPedido=" + id, "_blank", "resizable=yes,location=no,menubar=no,scrollbars=yes,scrolling=yes,height=500,width=700",false);		
	}
	
	function abrirObra(id) {
		window.open("comprasTableroObra.asp?idObra=" + id, "_blank", "resizable=yes,location=no,menubar=no,scrollbars=yes,scrolling=yes",false);		
	}
	function abrirContrato(id) {
		window.open("comprasCTC.asp?idContrato=" + id, "_blank", "resizable=yes,location=no,menubar=no,scrollbars=yes,scrolling=yes",false);
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
		location.href = "comprasIndex.asp";
	}
	
	function irHome() {
		location.href = "comprasIndex.asp";
	}
		
	function irAdministracion() {
		location.href = "comprasAdministracion.asp";
	}
	
	function irObras() {
		location.href = "comprasObras.asp";
	}
	
	function irPCT() {
		location.href = "comprasPedidoCotizacion.asp";
	}
	
	function irPIC() {
		location.href = "comprasAdministrarCotizaciones.asp?fromAP=1";
	}
	
	function enviarMail(idPedido) {
		window.open("comprasEnvioPCTMail.asp?idPedido=" + idPedido, "_blank", "resizable=yes,location=no,menubar=no,statusbar=no,height=240,width=500",false);		
	}
	
	function seleccionarObra(ms) {
		var desc = ms.getSelectedItem();
		if (desc.indexOf('-') != -1) {
			var arr = desc.split('-');
			document.getElementById("idObra").value = arr[0];
			ms.setValue(arr[1]);
		} else {
			if (desc == "") document.getElementById("idObra").value = "";							
		}		
	}
	
	function startMagicSearch() {		
		var ms = new MagicSearch("", "divObra", 30, 2, "comprasStreamElementos.asp?tipo=obras");		
		ms.setToken(";");		
		ms.onBlur = seleccionarObra;	
		//ms.setValue(document.getElementById("idObra").value);
		ms.setValue('<%=cdObra%>');
		
		var msSolicitante = new MagicSearch("", "divSolicitante", 30, 2, "comprasStreamElementos.asp?tipo=personas");
		msSolicitante.setToken(";");
		msSolicitante.onBlur = seleccionarSolicitante;
		//msSolicitante.setValue(document.getElementById('<%=dsSolicitante%>').value);
		msSolicitante.setValue('<%=dsSolicitante%>');
	}
	function irDirecta(){
		location.href = "comprasAdministrarCotizaciones.asp";
	}		
	function bodyOnLoad() {	
		var tb = new Toolbar('toolbar', 7, 'images/compras/');
		tb.addButton("Home-16x16.png", "Home", "irHome()");		
		<% if (puedeCrear()) then %>			
			tb.addButton("PCT_new-16X16.png", "Ped. Precio", "irPCT()");
		<% end if %>
		tb.addButtonREFRESH("Recargar", "submitInfo()");		
		var swt = tb.addSwitcher("Search-16x16.png", "Buscar", "buscarOn()", "buscarOff()");		
		tb.addButton("Direct_purchase_folder-16x16.png", "Compra Directa", "irDirecta()");
		tb.addButton("Direct_purchase_folder-16x16.png", "Ver PICs", "irPIC()");
		tb.addButton("Direct_purchase_folder-16x16.png", "Ver PDC", "irPDC()");
		tb.draw();
		<%	if (cint(hayBusqueda) = 1) then %>
				tb.changeState(swt);
				buscarOn();				
		<%	end if 			%>				
		<% 	if (not pedidos.eof) then %>
			var pgn = new Paginacion("paginacion");				
			pgn.paginar(<% =paginaActual %>, <% =totalRegistros %>, <% =mostrar %>, 50, "comprasAdministrarPedidos.asp<% =params %>");			
			
		<%	end if 	%>
	}
	function irPDC(){
		location.href = "comprasPDCAdministrar.asp";
	}
	function setOrder(p_campo,p_orden){
		document.getElementById("sqlCampoOrder").value = p_campo;
		document.getElementById("sqlTipoOrder").value = p_orden;
		submitInfo();
	}
	
	/*
	function keyPressed : se encarga de controlar los datos que se ingresan en el campo de Cantidad Proveedores
						  Restringe para que se ingrese unsa sola vez el signo + o -, y que no se ingresen letras.
						  Este control es necesario para luego validarlo desde el store	*/
	function keyPressed(pEvent,e) {			
		var ret;
		ret = controlIngreso(e,pEvent,'N');
		if(ret == false){
			var key=(document.all) ? pEvent.keyCode : pEvent.which;						
			if((key == 43)||(key == 45)){
				var auxText = new String();
				var caracter = String.fromCharCode(key);
				auxText = e.value;
				if ((auxText.indexOf("-") < 0)&&(auxText.indexOf("+") < 0))
					ret = true;
				else
					ret = false;					
			}
		}
		return ret;
	}
</script>
</head>
<body onLoad="bodyOnLoad()">
	<% call GF_TITULO2("kogge64.gif","Administrar Pedidos de Compras - Pedidos de Precio") %>		
	<div id="toolbar"></div>
	<br>
	<form name="frmSel" id="frmSel">
	<div id="busqueda" class="divOculto">
	<table width="90%" cellspacing="0" cellpadding="0" align="center" border="0">
       <input type="hidden" name="accion" id="accion" value="">
	   <input type="hidden" name="sqlCampoOrder" id="sqlCampoOrder" value="<%=sql_Campo_Order%>">
	   <input type="hidden" name="sqlTipoOrder" id="sqlTipoOrder" value="<%=sql_Tipo_Order%>">
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
								<td><input type="text" name="cdPedido" value="<% =cdPedido %>"></td>
							
								<%
								strSQL = "Select divi.IDDIVISION, divi.DSDIVISION from TBLDIVISIONES divi"
								Call executeQueryDb(DBSITE_SQL_INTRA, rsDivisiones, "OPEN", strSQL)
								%>                                
                                <td align="right"><% =GF_TRADUCIR("Division") %>:</td>
                                <td>                                
									<select style="z-index:-1;" name="idDivision">
									        <option SELECTED value="<% =SIN_DIVISION %>">- <% =GF_TRADUCIR("Seleccione") %> -
									<%		while (not rsDivisiones.eof)		
												selected = ""										
												if (CLng(rsDivisiones("IDDIVISION")) = CLng(idDivision)) then selected = "selected"
									%>
												<option value="<% =rsDivisiones("IDDIVISION") %>" <% =selected %>><% =rsDivisiones("DSDIVISION") %>                                        
									<%			rsDivisiones.MoveNext()
											wend 	
									%>
									</select>
                                </td>									
							</tr>

                            <tr>
								<td align="right"><% = GF_TRADUCIR("Obra") %>:</td>
								<td><div id="divObra"></div></td>
								<input type="hidden" id="idObra" name="idObra" value="<% =idObra %>">

                                <td align="right"><% =GF_TRADUCIR("Estado") %>:</td>
                                <td>                                
                                <select name="idEstado">
                                        <option SELECTED value="0">- <% =GF_TRADUCIR("Seleccione") %> -
                                        <% =createOption(ESTADO_TODOS, GF_TRADUCIR("Todos"), idEstado) %>
										<% =createOption(ESTADO_PCT_ABIERTO, GF_TRADUCIR("Apertura de Sobres Realizada"), idEstado) %>
										<% =createOption(ESTADO_PCT_ADJUDICADO, GF_TRADUCIR("Adjudicado"), idEstado) %>
										<% =createOption(ESTADO_PCT_APROBADO, GF_TRADUCIR("Aprobados"), idEstado) %>
										<% =createOption(ESTADO_PCT_AUTORIZADO, GF_TRADUCIR("Autorizados por el solicitante"), idEstado) %>
										<% =createOption(ESTADO_PCT_CANCELADO, GF_TRADUCIR("Cancelados"), idEstado) %>
										<% =createOption(ESTADO_PCT_COTIZADO, GF_TRADUCIR("Con Cotizacion Completa"), idEstado) %>
										<% =createOption(ESTADO_PCT_PUBLICADO, GF_TRADUCIR("Enviados a Proveedores"), idEstado) %>
										<% =createOption(ESTADO_PCT_PENDIENTE, GF_TRADUCIR("Pendientes"), idEstado) %>										
                                </select>
                                </td>
                            </tr>									

							<tr>
								<td align="right"><%=GF_TRADUCIR("Solicitante") %>:</td>
								<td>
									<div id="divSolicitante"></div>			

									<input type="hidden" id="cdSolicitante" name="cdSolicitante" value="<% =cdSolicitante %>">
								</td>
								
									<td align="right"><% = GF_TRADUCIR("Titulo") %>:</td>
									<td><input type="text" name="Titulo" value="<% =titulo %>"></td>
								
							</tr>
							<tr>
                                <td align="right"><% =GF_TRADUCIR("Emisión") %>:</td>
								<td>
                                    <input type="text" size="2" maxLength="2" value="<% =txtDE %>" name="txtDiaEmision"> /
                                    <input type="text" size="2" maxLength="2" value="<% =txtME %>" name="txtMesEmision"> /
                                    <input type="text" size="4" maxLength="4" value="<% =txtAE %>" name="txtAnioEmision">
                                </td>
                                <td align="right"><% =GF_TRADUCIR("Cierre") %>:</td>
                                <td>
                                    <input type="text" size="2" maxLength="2" value="<% =txtDC %>" name="txtDiaCierre"> /
                                    <input type="text" size="2" maxLength="2" value="<% =txtMC %>" name="txtMesCierre"> /
                                    <input type="text" size="4" maxLength="4" value="<% =txtAC %>" name="txtAnioCierre">
								</td>                                
                            </tr>                            
							<tr>
								<td align="right"><% = GF_TRADUCIR("Cantidad proveedores") %>:</td>
								<td>
									<input type="text" onkeypress="return keyPressed(event, this)" size="4" maxLength="4" name="cantProv" id="cantProv" value="<% =cantProv %>">									
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
	
	<br>
	<table align="center" width="90%" class="reg_Header">			
			<tr><td colspan="10"><div id="paginacion"></div></td></tr>						
			<tr class="reg_Header_nav">
				<td width="15%" colspan="2" style="text-align: center"><img src="images\compras\arrow_up_12x12.gif" onclick='setOrder(<% =BUSQUEDA_CDPEDIDO  %>	,"asc")' style="cursor:pointer">&nbsp <% =GF_TRADUCIR("Pedido")			%> &nbsp <img src="images\compras\arrow_down_12x12.gif" onclick='setOrder(<% =BUSQUEDA_CDPEDIDO	 %>,"desc")' style="cursor:pointer"></td>				
				<td width="15%" colspan="2" style="text-align: center"><img src="images\compras\arrow_up_12x12.gif" onclick='setOrder(<% =BUSQUEDA_PART_PRES %>	,"asc")' style="cursor:pointer">&nbsp <% =GF_TRADUCIR("Ptda. Presup.")	%> &nbsp <img src="images\compras\arrow_down_12x12.gif" onclick='setOrder(<% =BUSQUEDA_PART_PRES %>,"desc")' style="cursor:pointer"></td>
				<td width="20%" 		    style="text-align: center"><% =GF_TRADUCIR("Solicitante")%></td>				
				<td width="20%"				style="text-align: center"><img src="images\compras\arrow_up_12x12.gif" onclick='setOrder(<% =BUSQUEDA_TITULO    %>	,"asc")' style="cursor:pointer">&nbsp <% =GF_TRADUCIR("Titulo") 		%> &nbsp <img src="images\compras\arrow_down_12x12.gif" onclick='setOrder(<% =BUSQUEDA_TITULO	 %>,"desc")' style="cursor:pointer"></td>
				<td width="10%" 		    style="text-align: center"><img src="images\compras\arrow_up_12x12.gif" onclick='setOrder(<% =BUSQUEDA_F_CIERRE  %>	,"asc")' style="cursor:pointer">&nbsp <% =GF_TRADUCIR("Cierre") 		%> &nbsp <img src="images\compras\arrow_down_12x12.gif" onclick='setOrder(<% =BUSQUEDA_F_CIERRE  %>,"desc")' style="cursor:pointer"></td>
				<td width="3%"				style="text-align: center"><% =GF_TRADUCIR("STS")%></td>
				<td width="10%" colspan="2" style="text-align: center"><img src="images\compras\arrow_up_12x12.gif" onclick='setOrder(<% =BUSQUEDA_REF		 %>	,"asc")' style="cursor:pointer">&nbsp <% =GF_TRADUCIR("REF")			%> &nbsp <img src="images\compras\arrow_down_12x12.gif" onclick='setOrder(<% =BUSQUEDA_REF		 %>,"desc")' style="cursor:pointer"></td>
			</tr>		
<%	
	if (not pedidos.eof) then			
			while (not pedidos.eof)
				
				pctImage = "PCT-16x16.png"
				if (pedidos("ESTADO") = ESTADO_PCT_CANCELADO) then pctImage = "PCTR-16x16.png"
				
				accesoPermitido = false					
				
%>
			<tr class="reg_Header_navdos" onMouseOver="javascript:lightOn(this)" onMouseOut="javascript:lightOff(this)">			

				<td style="text-align: center" onclick="javascript:abrirFicha(<% =pedidos("IDPEDIDO") %>,1)"><% =pedidos("CDPEDIDO") %></td>
				<td style="text-align: center"><img onclick="abrirPedido(<% =pedidos("IDPEDIDO") %>)" src="images/compras/<% =pctImage %>" title="Ver Pedido de Cotizacion"></td>
				<td style="text-align: center" onclick="javascript:abrirFicha(<% =pedidos("IDPEDIDO") %>,1)"><% =pedidos("CDOBRA") %></td>
				<td style="text-align: center" width="3%">
					<% 	if (pedidos("IDOBRA") > 0) then	%>
						<img src="images/compras/obr-16x16.png" onclick="abrirObra(<% =pedidos("IDOBRA") %>)" title="<% =GF_TRADUCIR("Ver Partida") %>">
					<%	end if	%>
				</td>
				<!-- falta-->
				<%
					cd_Solicitante = pedidos("CDSOLICITANTE")
					ds_Solicitante = getUserDescription(cd_Solicitante)
				%>
				<td style="text-align: center" onclick="javascript:abrirFicha(<% =pedidos("IDPEDIDO") %>,1)"><% =ds_Solicitante %></td>				
				<td onclick="javascript:abrirFicha(<% =pedidos("IDPEDIDO") %>, 1)"><%
					if (len(GF_TRADUCIR(pedidos("TITULO")))) > 35 then
						response.write left(GF_TRADUCIR(pedidos("Titulo")), 35) & "..."
					else
						response.write GF_TRADUCIR(pedidos("TITULO"))
					end if
					%>
				</td>				
				<td style="text-align: center" onclick="javascript:abrirFicha(<% =pedidos("IDPEDIDO") %>, 1)"><% =GF_FN2DTE(pedidos("FECHACIERRE")) %></td>
				<td style="text-align: center">
		<%			pct_idEstado = pedidos("ESTADO")
					
					Call actualizarEstado(pedidos)
					
					Select Case cint(pct_idEstado)
						Case ESTADO_PCT_PENDIENTE
								if ((isAdmin(pedidos("IDDIVISION"))) or (cd_Solicitante = session("Usuario"))) then  %>										
									<a href="javascript:abrirPedido(<% =pedidos("IDPEDIDO") %>)"><img src="images/compras/PCT_confirm-16x16.png" title="<% =GF_TRADUCIR("Se requiere que el solicitante apruebe el pedido.") %>"></a>
							<%	else
									'Es el usuario que cargo el pedido.%>
									<img onclick="javascript:abrirFicha(<% =pedidos("IDPEDIDO") %>,1)" src="images/compras/PCT_confirm-16x16.png" title="<% =GF_TRADUCIR("Se requiere que el solicitante apruebe el pedido.") %>">
							<%	end if
						Case ESTADO_PCT_AUTORIZADO	%>
							<a href="javascript:<% =setFunc("enviarMail(" & pedidos("IDPEDIDO") & ")") %>"><img src="images/compras/PCT_publish-16x16.png" title="<% =GF_TRADUCIR("Pedido listo para enviar a proveedores.") %>"></a>
		<%				Case ESTADO_PCT_PUBLICADO	%>
							<a title="<% =GF_TRADUCIR("Los proveedores estan presentando sus cotizaciones")%>" href="javascript:abrirFicha(<% =pedidos("IDPEDIDO") %>,2)"><font style="font-family:courier;" color=#ff0000 size=+2><b><%=getCantidadCotizacionesRecibidas(pedidos("IDPEDIDO"))%></b></font></a>
		<%				Case ESTADO_PCT_COTIZADO	%>
							<span style="cursor:pointer" onclick="javascript:<% =setFunc("abrirFicha(" & pedidos("IDPEDIDO") & ", 3)") %>"><img src="images/compras/bid_purchase-16x16.png" title="<% =GF_TRADUCIR("Pedido listo para apertura de sobres") %>"></span>
		<%				Case ESTADO_PCT_ABIERTO		%>
							<span style="cursor:pointer" onclick="javascript:<% =setFunc("abrirFicha(" & pedidos("IDPEDIDO") & ", 2)") %>"><img src="images/compras/Bid_purchase_open-16x16.png" title="<% =GF_TRADUCIR("Las cotizaciones ya estan disponibles para evaluar") %>"></span>
		<%				Case ESTADO_PCT_EN_ANALISIS	%>
							<span style="cursor:pointer" onclick="javascript:<% =setFunc("abrirFicha(" & pedidos("IDPEDIDO") & ", 1)") %>"><img src="images/compras/PCP-16x16.png" title="<% =GF_TRADUCIR("Se esta realizando el Analisis Comparativo...") %>"></span>
		<%				Case ESTADO_PCT_EN_FIRMA_AC	%>
							<span style="cursor:pointer" onclick="javascript:<% =setFunc("abrirFicha(" & pedidos("IDPEDIDO") & ", 1)") %>"><img src="images/compras/PCP-16x16.png" title="<% =GF_TRADUCIR("Se aguarda la firma Analisis Comparativo...") %>"></span>
		<%				Case ESTADO_PCT_ADJUDICADO	%>
							<span style="cursor:pointer" onclick="javascript:<% =setFunc("abrirFicha(" & pedidos("IDPEDIDO") & ", 1)") %>"><img src="images/compras/PCT_awarded-16x16.png" title="<% =GF_TRADUCIR("Pedido Adjudicado!") %>"></span>
		<%				Case ESTADO_PCT_APROBADO	%>
							<span style="cursor:pointer" onclick="javascript:<% =setFunc("abrirFicha(" & pedidos("IDPEDIDO") & ", 1)") %>"><img src="images/compras/PCT_completed-16x16.png" width="12px" height="12px" title="<% =GF_TRADUCIR("Pedido Completo!") %>"></span>
		<%				Case ESTADO_PCT_CANCELADO	%>
							<span style="cursor:pointer" onclick="javascript:<% =setFunc("abrirFicha(" & pedidos("IDPEDIDO") & ", 1)") %>"><img src="images/compras/PCT_cancelled-16x16.png" width="12px" height="12px" title="<% =GF_TRADUCIR("Pedido Cancelado") %>"></span>
		<%			End Select						%>
				</td>
					<%	if ((pedidos("ESTADO") >= ESTADO_PCT_ADJUDICADO) and (pedidos("ESTADO") <= ESTADO_PCT_APROBADO)) then	
							if (not isnull(pedidos("ID_CONTRATO"))) then	%>
								<td align="right"><% =pedidos("CDCONTRATO") %></td>
								<td align="center" width="3%"><img src="images/compras/CTC-16x16.png" title="Ver Contrato" style="cursor:pointer" onclick="javascript:abrirContrato(<% =pedidos("ID_CONTRATO") %>)"></td>
					<%		elseif (not isnull(pedidos("ID_COTIZACION"))) then
								if cint(pedidos("CTZCANTIDAD"))>1 then
									%>
									<td align="right"><% =GF_TRADUCIR("+ 1") %></td>
									<td align="center" width="3%"><img src="images/compras/PICS-16x16.png" title="Ver Pedidos Internos de Compras" style="cursor:pointer" onclick="javascript:abrirFicha(<% =pedidos("IDPEDIDO") %>,4)"></td>
									<%
								else
									%>
									<td align="right"><% =pedidos("ID_COTIZACION") %></td>
									<td align="center" width="3%"><img src="images/compras/PIC-16x16.png" title="Ver Pedido Interno de Compra" style="cursor:pointer" onclick="javascript:abrirCotizacion(<% =pedidos("ID_COTIZACION") %>)"></td>
									<%
								end if
							else %>
								<td>&nbsp;</td><td width="3%">&nbsp;</td>
		<%					end if
						else %>
							<td>&nbsp;</td><td width="3%">&nbsp;</td>
		<%				end if %>
			</tr>
	<%			pedidos.MoveNext()
			wend
	else
%>
			<tr class="TDNOHAY"><td colSpan="10"><% =GF_TRADUCIR("No hay informacion disponible en estos momentos") %></td></tr>
<%  end if %>
		</table>
</form>
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
		if (isNumeric(id)) then
			if (CLng(id) = CLng(param)) then sel ="selected"
		else
			if (id = param) then sel ="selected"
		end if
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