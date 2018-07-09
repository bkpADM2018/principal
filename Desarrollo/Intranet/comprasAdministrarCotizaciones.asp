<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientossql.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/md5.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosTitulos.asp"-->
<%
Call comprasControlAccesoCM(RES_CD)

'-----------------------------------------------------------------------------------------------
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
'-----------------------------------------------------------------------------------------------
Function filtrarSeguridad(ByRef myWhere)
	dim myDivisionesAdmin
	'Filtro
	cdUsuario = session("Usuario")
	myDivisionesAdmin = getListaCargosAdmin()
	if (myDivisionesAdmin <> "") then  		 
		myWhere = myWhere & " WHERE ((CDUSUARIO='" & cdUsuario & "')"
		myWhere = myWhere & " OR (IDDIVISION IN (0," & myDivisionesAdmin & ")))"		
	else
		myWhere = myWhere & " WHERE (CDUSUARIO='" & cdUsuario & "')"
	end if
	filtrarSeguridad = myWhere
End Function
'-----------------------------------------------------------------------------------------------
Function filtrarCotizaciones(idPIC, idProveedor, fechaEmision, observaciones, momento, idDivision, pEstado)
	'Filtro
	Dim ret, campoImporte
	
	ret = filtrarSeguridad(ret)
	if ((idPIC <> 0) and (idPIC <> "")) then Call mkWhere(ret, "IDCOTIZACION", idPIC, "=", 1)
	if ((idProveedor <> "0") and (idProveedor <> "")) then Call mkWhere(ret, "IDPROVEEDOR", idProveedor, "=", 1)
	if ((idDivision <> 0) and (idDivision <> "")) then Call mkWhere(ret, "IDDIVISION", idDivision, "=", 1)
	if (fechaEmision <> "") then Call mkWhere(ret, "FECHAENTREGA", fechaEmision, "LIKE", 3)	        
	if(picSearch_radioMoneda = TIPO_MONEDA_PESO) then campoImporte = "IMPORTEPESOS"
	if(picSearch_radioMoneda = TIPO_MONEDA_DOLAR) then campoImporte = "IMPORTEDOLARES"
			
	if (picSearch_import <> 0) then 		
		if (isnumeric(picSearch_import)) then
			select case picSearch_radioImport
				case "Mayor"
					Call mkWhere(ret,campoImporte,picSearch_import * 100,">",1)
				case "Menor"
					Call mkWhere(ret,campoImporte,picSearch_import * 100,"<",1)
				case "Igual"
					Call mkWhere(ret,campoImporte,picSearch_import * 100,"=",1)
			end select
		end if
	end if	
	if (pEstado <> "ALL") then
		if (pEstado = "") then		
			Call mkWhere(ret, "ESTADO", CTZ_ANULADA, "<>", 3)
			Call mkWhere(ret, "ESTADO", CTZ_FACTURADA, "<>", 3)
		else
			Call mkWhere(ret, "ESTADO", pEstado, "=", 3)
		end if
	end if
	filtrarCotizaciones = ret	
End Function
'-----------------------------------------------------------------------------------------------
Function filtrarCotizacionesArticulo(pIdPIC, pidArticulo)  
dim ret
if ((idPIC <> 0) and (idPIC <> "")) then Call mkWhere(ret, "IDCOTIZACION", pIdPIC, "=", 1)
if(pidArticulo <> "") then Call mkWhere(ret, "IDARTICULO", pidArticulo, "=", 1)
filtrarCotizacionesArticulo = ret
End Function
'-----------------------------------------------------------------------------------------------
Function obtenerListaCotizaciones(idPIC, cdPedido, idProveedor, fechaEmision, observaciones, momento, idDivision, pidArticulo, pCdContrato, pCdSolicitante, pSuccbt, pNrocbt, pEstado) 
	Dim strSQL, rs, myWhere, firstRecord, conn	
	'Ajusto Paginacion	
	strSQL =" SELECT DISTINCT CTZ.* , PCT.CDPEDIDO, EMP.NROEMP AS IDEMPRESA, EMP.NOMEMP AS DSEMPRESA, EMP.TIPDOC AS TIPODOCUMENTO, RP.IDPIC IDPIC, FAC.FACTURADO,FLS.maximo cantFiles,CON.CDCONTRATO " & _
        " FROM (Select * from tblctzcabecera " & filtrarCotizaciones(idPIC, idProveedor, fechaEmision, observaciones, momento, idDivision, pEstado) & ")  CTZ " & _
        " INNER JOIN (select IDARTICULO, IDCOTIZACION from TBLCTZDETALLE " & filtrarCotizacionesArticulo(idPIC, pidArticulo) & ") ART on ART.IDCOTIZACION = CTZ.IDCOTIZACION " & _ 
        " INNER JOIN (select IDCOTIZACION, sum(FACTURADO) FACTURADO from tblctzdetalle group by IDCOTIZACION) FAC on FAC.IDCOTIZACION=CTZ.IDCOTIZACION " & _
        " INNER JOIN (Select IDCOTIZACION, CDUSUARIO from TBLCTZFIRMAS where SECUENCIA=0) SOL on CTZ.IDCOTIZACION=SOL.IDCOTIZACION "
	if ((pSuccbt > 0) and (pNrocbt > 0)) then
		strSQL = strSQL & " INNER JOIN	(Select Distinct NROINT, IDPIC from MEP001C where anulado = 'N') P on P.IDPIC=CTZ.IDCOTIZACION " &_
						" inner join	 [Database].[dbo].MEP001A M on M.NROINT=P.NROINT "
	end if
    strSQL = strSQL & " LEFT JOIN [Database].[dbo].MET001A EMP ON CTZ.idproveedor = EMP.nroemp " & _ 
        " LEFT JOIN TBLPCTCABECERA PCT on PCT.IDPEDIDO=CTZ.IDPEDIDO "& _ 
        " LEFT JOIN (Select Distinct IDPIC from TBLREMPIC A inner join TBLREMCABECERA B on A.IDREMITO=B.IDREMITO and B.ESTADO = " & ESTADO_ACTIVO & ") RP on RP.IDPIC=CTZ.IDCOTIZACION " & _
        " LEFT JOIN (select max(fileno) maximo,idcotizacion from TBLCTZBINARYFILES group by idcotizacion) FLS on CTZ.IDCOTIZACION = FLS.IDCOTIZACION "&_
        " LEFT JOIN TBLOBRACONTRATOS CON ON CON.IDCONTRATO = CTZ.IDCONTRATO "
	'Filtros Generales
	if ((cdPedido <> "0") and (cdPedido <> "")) then Call mkWhere(myWhere, "PCT.CDPEDIDO", UCASE(cdPedido), "like", 3)
    if (pCdContrato <> "") then Call mkWhere(myWhere, "CON.CDCONTRATO", UCASE(pCdContrato), "like", 3)    
    if (pCdSolicitante <> "") then Call mkWhere(myWhere, "SOL.CDUSUARIO", UCASE(pCdSolicitante), "=", 3)    
	if ((pSuccbt > 0) and (pNrocbt > 0)) then
		Call mkWhere(myWhere, "M.succbt", pSuccbt, "=", 1)    
		Call mkWhere(myWhere, "M.nrocbt", pNrocbt, "=", 1)    
	end if
	strSQL = strSQL & " " & myWhere
	strSQL = strSQL & " ORDER BY CTZ.IDCOTIZACION desc"		
	'response.write strSQL
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set obtenerListaCotizaciones = rs
End Function
'-----------------------------------------------------------------------------------------------
'Funcion responsable de determinar si una nota de aceptación puede o no ser enviada segun el tipo de compra que se esta realizando.
Function mostrarNDA(rsCotizaciones)
	
	mostrarNDA = false
	if (CLng(rsCotizaciones("IDPEDIDO")) = 0) then ' Es compra directa?
		if (rsCotizaciones("ESTADO") = CTZ_FIRMADA) then mostrarNDA=true
	else
		'Es compra con pedido
		'Se utiliza el mismo control que en comprasFichaPCTTab1
		if (rsCotizaciones("ESTADO") <> CTZ_ANULADA) then mostrarNDA=true
	end if
				
End Function
'**********************************************************
'***	COMIENZO DE PAGINA
'**********************************************************
dim idSector, descripcion, paginaActual, mostrar, flagAuditor, cdSolicitante, dsSolicitante, cdUsuario
dim idProveedor, cdProveedor, dsProveedor, fechaEmision, observaciones, momento
dim rsCotizaciones, lineasTotales, txtDE, txtME, txtAE, idPIC, myTitle, cdPedido
Dim estado, trClass, idDivision, picSearch_radioImport,picSearch_radioMoneda, picSearch_import 
Dim cdArticulo, idArticulo, dsArticulo,cdContrato, nrocbt, succbt
index = 0
idSector = GF_PARAMETROS7("idSector","",6)
call addParam("idSector", idSector, params)


descripcion = GF_PARAMETROS7("descripcion","",6)
call addParam("descripcion", descripcion, params)

paginaActual = GF_PARAMETROS7("numeroPagina",0,6)
if (paginaActual = 0) then paginaActual=1
mostrar = GF_PARAMETROS7("registrosPorPagina",0,6)
if (mostrar = 0) then mostrar = 10

idPIC = GF_PARAMETROS7("idPIC", 0, 6)
call addParam("idPedido", idPIC, params)

estado = GF_PARAMETROS7("estado","",6)
call addParam("estado", estado, params)
succbt = GF_PARAMETROS7("succbt", 0, 6)
call addParam("succbt", succbt, params)
nrocbt = GF_PARAMETROS7("nrocbt", 0, 6)
call addParam("nrocbt", nrocbt, params)
if ((succbt > 0) and (nrocbt > 0)) then estado = "ALL"

cdPedido = GF_PARAMETROS7("cdPedido", "", 6)
call addParam("cdPedido", cdPedido, params)
idProveedor = GF_PARAMETROS7("idProveedor", 0, 6)
call addParam("idProveedor", idProveedor, params)
cdProveedor = GF_PARAMETROS7("cdProveedor", "", 6)
call addParam("cdProveedor", cdProveedor, params)
dsProveedor = GF_PARAMETROS7("dsProveedor", "", 6)
call addParam("dsProveedor", dsProveedor, params)
txtAE = GF_PARAMETROS7("txtAnioEmision","",6)
call addParam("txtAnioEmision", txtAE, params)
txtME = GF_PARAMETROS7("txtMesEmision","",6)
call addParam("txtMesEmision", txtME, params)
txtDE = GF_PARAMETROS7("txtDiaEmision","",6)
call addParam("txtDiaEmision", txtDE, params)
idDivision = GF_PARAMETROS7("idDivision", 0, 6)
call addParam("idDivision", idDivision, params)

cdSolicitante = GF_PARAMETROS7("cdSolicitante", "", 6)
call addParam("cdSolicitante", cdSolicitante, params)
dsSolicitante = GF_PARAMETROS7("divSolicitante", "", 6)

'CNA
idArticulo = GF_PARAMETROS7("idArticulo", "", 6)
call addParam("idArticulo", idArticulo, params)
dsArticulo = GF_PARAMETROS7("dsArticulo", "", 6)
call addParam("dsArticulo", dsArticulo, params)
cdArticulo = GF_PARAMETROS7("cdArticulo", "", 6)
call addParam("cdArticulo", cdArticulo, params)
cdContrato = GF_PARAMETROS7("cdContrato", "", 6)
call addParam("cdContrato", cdContrato, params)
'linea agregada
picSearch_import = GF_PARAMETROS7("ImpTotal", 2, 6)
picSearch_radioMoneda  = GF_PARAMETROS7("radio_TipoMoneda" ,"",6)
picSearch_radioImport  = GF_PARAMETROS7("radio_Import" ,"",6)
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
Call GP_ConfigurarMomentos()
Set rsCotizaciones = obtenerListaCotizaciones(idPIC, cdPedido, idProveedor, fechaEmision, observaciones, momento, idDivision, idArticulo, cdContrato, cdSolicitante, succbt, nrocbt, estado)
Call setupPaginacion(rsCotizaciones, paginaActual, mostrar)

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title>Administrar Pedidos Internos de Compra</title>

<link rel="stylesheet" href="css/main.css" type="text/css">
<link rel="stylesheet" href="css/paginar.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/MagicSearch.css" type="text/css">
<link href="css/jqueryUI/custom-theme/jquery-ui-1.8.15.custom.css" rel="stylesheet" type="text/css" />

<style type="text/css">
.titleStyle {
	font-weight: bold;
	font-size: 20px;
}

.divOculto {
	display: none;
}
</style>
<script type="text/javascript" src="scripts/jquery/jquery-1.5.1.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.15.custom.min.js"></script>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/MagicSearchObj.js"></script>
<script type="text/javascript" src="scripts/script_fechas.js"></script>
<script type="text/javascript" src="scripts/paginar.js"></script>
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript">
	var ch = new channel();	
	var msArticulos = new Array();
	var popUpNDA;
		
	function abrirPedido(id) {
		window.open("comprasFichaPedidoCotizacion.asp?idPedido=" + id + "&tab=1", "_blank", "resizable=yes,location=no,scrollbars=yes,menubar=no,statusbar=no,height=500,width=500",false);
	}
	
	function abrirCotizacion(idCTZ) {
		window.open("comprasPICPrint.asp?idCotizacionElegida=" + idCTZ, "_blank", "resizable=yes,location=no,menubar=no,scrollbars=yes,scrolling=yes,height=500,width=700",false);				
	}
	function editarCTZ(idCTZ) {
	    if (confirm("Modificar un PIC implica cambiar su fecha de creación. Esta seguro que desea modificarlo? Si lo que desea es ajustarlo ingrese a la administración del mismo y realice los ajustes directamente en los items.")) {
		    document.location.href = "comprasPIC.asp?idCotizacionElegida=" + idCTZ;
		}
	}	
	function SeleccionarProveedor(ms){
		var desc = ms.getSelectedItem();
		if (desc.indexOf('-') != -1) {
			var arr = desc.split('-');
			document.getElementById("idProveedor").value = arr[0];
			document.getElementById("dsProveedor").value = arr[1];
			ms.setValue(arr[1]);
		} else {
			if (desc == ""){
				document.getElementById("idProveedor").value = 0;
				document.getElementById("dsProveedor").value = "";
				ms.setValue("");
			}	
		}				
	}
	function SeleccionarArticulo(ms)
	{
		var desc = ms.getSelectedItem();
		if (desc.indexOf('|') != -1) 
		{
			var arr = desc.split('|');
			document.getElementById("idArticulo").value = arr[0];
			document.getElementById("divArticulo").innerHTML = arr[0];
			ms.setValue(arr[1]);
			document.getElementById("dsArticulo").value = arr[1];
			var arr2 = arr[1].split('[');				
			document.getElementById("dsArticulo").value = arr2[0];
			ms.setValue(arr2[0]);
		}
		else 
		{
			if (desc == "")
			{
				document.getElementById("idArticulo").value = "";
				document.getElementById("dsArticulo").value = "";
				document.getElementById("divArticulo").innerHTML = "";
			}	
		}				
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
	
	function irNewCotizacion() {
		location.href = "comprasPIC.asp";
	}	
	function irPedidos() {
		location.href = "comprasAdministrarPedidos.asp";
	}	
		
	function bodyOnLoad() {	
		var tb = new Toolbar('toolbar', 6, 'images/compras/');
		tb.addButton("Home-16x16.png", "Home", "irHome()");		
		<% if puedeCrear then %>
			tb.addButton("PCT_new-16X16.png", "Nueva", "irNewCotizacion()");
		<% end if %>
		tb.addButtonREFRESH("Recargar", "submitInfo()");		
		tb.addButton("Quote_purchase-16x16.png", "Ped. Precio", "irPedidos()");		
		tb.draw();				
				
		<% 	if (not rsCotizaciones.eof) then %>
				var pgn = new Paginacion("paginacion");							
				pgn.paginar(<% =paginaActual %>, <% =rsCotizaciones.RecordCount %>, <% =mostrar %>, 50, "comprasAdministrarCotizaciones.asp<% =params %>");					
		<%	end if 	%>
		
		startMagicSearch();
	}
	
	function startMagicSearch() {		
		var msProveedor = new MagicSearch("", "companyName0", 30, 2, "comprasStreamElementos.asp?tipo=empresas");
		msProveedor.setMinChar(3);
		msProveedor.setToken(";");
		msProveedor.onBlur = SeleccionarProveedor;
		msProveedor.setValue('<% =dsProveedor %>');	
		
		var ms = new MagicSearch("", "articuloItem", 30, 2, "comprasStreamElementos.asp?tipo=articulos");		
		ms.setMinChar(3);
		ms.setToken(";");		
		ms.onBlur = SeleccionarArticulo;	
		ms.setValue('<% =dsArticulo%>');
		
		var msSolicitante = new MagicSearch("", "divSolicitante", 30, 4, "comprasStreamElementos.asp?tipo=personas");
		msSolicitante.setToken(";");
		msSolicitante.onBlur = seleccionarSolicitante;		
		msSolicitante.setValue('<% =dsSolicitante %>');
	}	

	function anularCTZ(idCotizacion, idPedido, img){
		if (confirm("Esta seguro que desea anular este Pedido Interno?")) {
			img.src = "images/loading_small_green.gif"
			ch.bind("comprasAnularCTZAjax.asp?idCotizacion=" + idCotizacion + "&idPedido=" + idPedido, "anularCTZCallback('" + img.id + "')");
			ch.send();			
		}		
	}
	function anularCTZCallback(pId){
		<% if (estado="ALL") then %>
			var myImg = document.getElementById(pId);
			myImg.src="images/1p.gif";
			myImg.onClick="";
		<% else %>
			var myTable = document.getElementById("TBL_COTIZACIONES");
			var myImg = document.getElementById(pId);
			myTable.deleteRow(myImg.parentNode.parentNode.rowIndex);
		<% end if %>
	}
	function abrirREMPIC(idpic){
		window.open("comprasPIC.asp?verRemitos=true&idCotizacionElegida=" + idpic);
	}
	
	function openPicFiles(idcotizacion){
		winPopUp('Iframe', 'comprasPicFiles.asp?idcotizacion='+idcotizacion, "500", "300", 'Archivos Cotizacion', '');
	}
	function abrirNDA(idPedido, idcotizacion){
		window.open('comprasListaNDA.asp?idPedido='+idPedido+'&idCotizacion='+idcotizacion,'_blank','resizable=yes,location=no,menubar=no,statusbar=no,height=400,width=500,scrollbars=yes',false)
		
	}		
	function reloadPage(){
		window.location.reload();
	}
	function actualizarNDA(idPedido,idcotizacion,accion){
		document.getElementById("iconNDA_"+idcotizacion).src="images/compras/loading_small_orange.gif";
		ch.bind("comprasNDAAjax.asp?idPedido="+idPedido+"&idCotizacion="+idcotizacion+"&accion="+accion,"envioNDA_callBack("+idPedido+","+idcotizacion+")");
		ch.send();			
	}
	var chn = new channel();	
	function envioNDA_callBack(idPed,idCot) {
		chn.bind("comprasEnvioNDAMail.asp?idPedido=" + idPed +"&idCotizacion="+ idCot,"");
		chn.send();
		document.getElementById("iconNDA_"+idCot).src="images/compras/mail_sent-16x16.png";
	}	
	function cerrarPopUpNDA()
	{
		popUpNDA.hide();
	}

</script>
</head>
<body onLoad="bodyOnLoad()">	
	
	<div id="toolbar"></div>
	<br>
<form name="frmSel" id="frmSel">
	<input type="hidden" name="accion"   id="accion"   value="">
	
	<div class="tableaside size100"> <!-- BUSCAR -->

		<h3> Filtros </h3>
	  
		<div id="searchfilter" class="tableasidecontent">		
		
			<div class="col16 reg_header_navdos"> <%=GF_Traducir("Nro PIC:")%> </div>
			<div class="col16"> <input type="text" size="5" name="idPIC" value="<% =idPIC %>"></div>
			
			<div class="col16 reg_header_navdos"> <%=GF_Traducir("Proveedor:")%> </div>
			<div class="col16"> 
				<div id="companyName0"></div>
				<input type="hidden" id="idProveedor" name="idProveedor" value="<%=idProveedor%>">
				<input type="hidden" id="dsProveedor" name="dsProveedor" value="<%=dsProveedor%>">
			</div>
			
			<div class="col16 reg_header_navdos"> <%=GF_Traducir("Cbte. Proveedor:")%> </div>
			<div class="col16"> 
				<input type="text" size="5" maxlength="4" name="succbt" value="<% =succbt %>">	-
				<input type="text" size="10" maxlength="8" name="nrocbt" value="<% =nrocbt %>">	
			</div>
			
			<div class="col16 reg_header_navdos"> <%=GF_Traducir("Estado:")%> </div>
			<div class="col16"> 
				<select id="estado" name="estado">
					<option value=""					<% if estado = ""			then Response.Write "Selected"%>><%=GF_traducir("Activo")	%></option>																
					<option value="<%=CTZ_FIRMADA	%>" <% if estado = CTZ_FIRMADA		then Response.Write "Selected"%>><%=GF_traducir("Autorizados")	%></option>
					<option value="<%=CTZ_PENDIENTE	%>" <% if estado = CTZ_EN_AJUSTE	then Response.Write "Selected"%>><%=GF_traducir("Ajuste Pend.")	%></option>
					<option value="<%=CTZ_FACTURADA	%>" <% if estado = CTZ_FACTURADA	then Response.Write "Selected"%>><%=GF_traducir("Facturado")	%></option>
					<option value="<%=CTZ_ANULADA	%>" <% if estado = CTZ_ANULADA		then Response.Write "Selected"%>><%=GF_traducir("Anulado")		%></option>
					<option value="ALL"					<% if estado = "ALL"			then Response.Write "Selected"%>><%=GF_traducir("Todos")		%></option>									
				</select>
			</div>

			<div class="col16 reg_header_navdos"> <%=GF_Traducir("C&oacuted. Pedido:")%> </div>
			<div class="col16"><input type="text" name="cdPedido" value="<% =cdPedido %>"></div>
			
			<div class="col16 reg_header_navdos"> <%=GF_Traducir("Fec. Entrega:")%> </div>
			<div class="col16">
				<input type="text" size="1" maxLength="2" value="<% =txtDE %>" name="txtDiaEmision" onBlur="javascript:ControlarDia(this);"> /
				<input type="text" size="1" maxLength="2" value="<% =txtME %>" name="txtMesEmision" onBlur="javascript:ControlarMes(this);"> /
				<input type="text" size="3" maxLength="4" value="<% =txtAE %>" name="txtAnioEmision" onBlur="javascript:ControlarAnio(this);">
			</div>					
			
			<div class="col16 reg_header_navdos"> <%=GF_Traducir("Art&iacuteculo:")%> </div>
			<div class="col16">
				<div id="articuloItem"></div>
				<input type="hidden" id="idArticulo" name="idArticulo" value="<%= idArticulo%>">
				<input type="hidden" id="dsArticulo" name="dsArticulo" value="<%= dsArticulo%>">
			</div>
			
			<div class="col16 reg_header_navdos"> <%=GF_Traducir("Contrato:")%> </div>
			<div class="col16"><input type="text" name="cdContrato" id="cdContrato" value="<% =cdContrato %>" style="text-transform:uppercase"></div>
			
			<div class="col16 reg_header_navdos"> <%=GF_Traducir("Solicitante:")%> </div>
			<div class="col16">
				<div id="divSolicitante"></div>																		
				<input type="hidden" id="cdSolicitante" name="cdSolicitante" value="<% =cdSolicitante %>">
			</div>
			
			<div class="col16 reg_header_navdos"> <%=GF_Traducir("Importe:")%> </div>
			<div class="col16">
				<input type="text" size="13" name="ImpTotal" id="ImpTotal" value="<% if (picSearch_import <> 0) then response.write picSearch_import  %>" onKeyPress="return controlIngreso(this, event, 'I')">
				<input type="radio" name="radio_TipoMoneda" id="radio_TipoMoneda" value="$" <%if (picSearch_radioMoneda = TIPO_MONEDA_PESO) then %>checked="checked"<%end if%> onClick="selectedRadio(this.value)"/><% = GF_TRADUCIR("$")%>
				<input type="radio" name="radio_TipoMoneda" id="radio_TipoMoneda" value="US$" <%if (picSearch_radioMoneda = TIPO_MONEDA_DOLAR) then %>checked="checked"<%end if%> onClick="selectedRadio(this.value)"/><% = GF_TRADUCIR("US$")%>
				<input type="radio" name="radio_Import" id="radio_Import" value="Menor" <%if (picSearch_radioImport = "Menor") then %>checked="checked"<%end if%> onClick="selectedRadio(this.value)"/><% = GF_TRADUCIR("Menor")%>
				<input type="radio" name="radio_Import" id="radio_Import" value="Igual" <%if (picSearch_radioImport = "Igual") then %>checked="checked"<%end if%> onClick="selectedRadio(this.value)"/><% = GF_TRADUCIR("Igual")%>
				<input type="radio" name="radio_Import" id="radio_Import" value="Mayor" <%if (picSearch_radioImport = "Mayor") then %>checked="checked"<%end if%> onClick="selectedRadio(this.value)"/><% = GF_TRADUCIR("Mayor")%>			
			</div>
			
			<div class="col16 reg_header_navdos"> <%=GF_Traducir("Divisi&oacuten:")%> </div>
			<div class="col16"> 
				<select name="idDivision" id="idDivision">
					<option VALUE=""><%=GF_TRADUCIR("Seleccionar...")%></option>
					<%
					strSQL = "SELECT * FROM TBLDIVISIONES"
					call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
					while not rs.eof
						if rs("IDDIVISION") = idDivision then
							mySelected = "SELECTED"
						else
							if idDivision = 0 then	idDivision = rs("IDDIVISION")
							mySelected = ""
						end if
						%>
						<option title="<%=rs("DSDIVISION")%>" VALUE="<%=rs("IDDIVISION")%>" <%=mySelected%>><%=rs("DSDIVISION")%></option>
						<%
						rs.movenext
					wend
					call executeQueryDb(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
					%>
				</select>		
			</div>
			
			<span class="btnaction"><input type="submit"  onclick='submitInfo()' value="<% =GF_TRADUCIR("Buscar") %>" id=submitir name=submitir></span>
		</div>
	<div>
	<input type="hidden" value="<%=tipoCompra%>" name="tipoCompra" id="tipoCompra">
	<input type="hidden" value="<%=fromAP%>" name="fromAP" id="fromAP">
	<input type="hidden" name="valorRadios" id="valorRadios"  value=<%=picSearch_radioImport%> >
</form>		
	<br>
	
	<table class="datagrid" width="100%" align="center">
	    <thead>
			<tr>
				<th align="center">	 			 	<% =GF_TRADUCIR("Numero")	%> </th>
				<th align="center" colspan="2">	 	<% =GF_TRADUCIR("Pedido")	%> </th>
				<th align="center">				 	<% =GF_TRADUCIR("Proveedor")%> </th>			
				<th align="center">					<% =GF_TRADUCIR("Contrato")	%> </th>
				<th align="center">					<% =GF_TRADUCIR("Importe")  %> </th>
				<th align="center" width="20px"> 	. </th>
				<th align="center" width="20px"> 	. </th>
				<th align="center" width="20px"> 	. </th>
				<th align="center" width="20px"> 	. </th>
				<th align="center" colspan="2"> 	. </th>
			</tr>
		</thead>
		<tbody>
		<%
		reg=0	
		if (not rsCotizaciones.eof) then			
			while ((not rsCotizaciones.eof) and (reg < CInt(mostrar)))
				reg=reg+1			
				trClass= "reg_Header_navdos"
				if (CStr(rsCotizaciones("ESTADO")) = CTZ_ANULADA) then trClass = trClass & " reg_Header_rejected"
				%>
				<tr>
				
					<td align="center" onclick="javascript:abrirCotizacion(<% =rsCotizaciones("IDCOTIZACION") %>)">
						<% = rsCotizaciones("IDCOTIZACION") %>
					</td>
					<td align="center" onclick="javascript:abrirCotizacion(<% =rsCotizaciones("IDCOTIZACION") %>)"><% =rsCotizaciones("CDPEDIDO")	%></td>
					<td style="text-align: center" width="1%">
					<%	if (rsCotizaciones("IDPEDIDO") > 0) then	%>
						<span style="cursor:pointer" onclick="abrirPedido(<% =rsCotizaciones("IDPEDIDO") %>)"><img src="images/compras/pct-16x16.png" title="Ver Pedido de Cotizacion" /></span>
					<%	end if	%>
					</td>
					<td align="center" width="30%"  onclick="javascript:abrirCotizacion(<% =rsCotizaciones("IDCOTIZACION") %>)"><% =getDescripcionProveedor(rsCotizaciones("IDPROVEEDOR"))	%></td>			
					<td align="center" onclick="javascript:abrirCotizacion(<% =rsCotizaciones("IDCOTIZACION") %>)"><% =rsCotizaciones("CDCONTRATO")				%></td>
					<%  if (rsCotizaciones("CDMONEDA") = MONEDA_PESO) then %>
						<td align="right"  onclick="javascript:abrirCotizacion(<% =rsCotizaciones("IDCOTIZACION") %>)"><% =getSimboloMoneda(MONEDA_PESO) & " " & GF_EDIT_DECIMALS(CDbl(rsCotizaciones("IMPORTEPESOS")),2)	%></td>
					<%  else %>				    
						<td align="right"  onclick="javascript:abrirCotizacion(<% =rsCotizaciones("IDCOTIZACION") %>)"><% =getSimboloMoneda(MONEDA_DOLAR) & " " & GF_EDIT_DECIMALS(CDbl(rsCotizaciones("IMPORTEDOLARES")),2)	%></td>
					<%  end if %>				    
					<td align="center">
					<% if (not isnull(rsCotizaciones("cantFiles"))) then%>
						<img src="images/compras/clip.png" title="<%=cdbl(rsCotizaciones("cantFiles"))+1%> Archivos Asociados" onclick="openPicFiles('<%=rsCotizaciones("IDCOTIZACION")%>')">
					<%end if%>
					</td>
					<%							
						'Sin no esta anulada y no esta pagada, permite modificar.
						if ((CDbl(rsCotizaciones("IDCONTRATO")) = 0) and (CDbl(rsCotizaciones("FACTURADO")) = 0) and (CStr(rsCotizaciones("ESTADO")) <> CTZ_ANULADA) and (CStr(rsCotizaciones("ESTADO")) <> CTZ_EN_AJUSTE)) then
					%>
						<td align="center"> <img title="<%=GF_TRADUCIR("Editar Pedido Interno")%>" id="ID_<%=rsCotizaciones("IDCOTIZACION")%>" src="images\compras\edit-16x16.png" onclick='editarCTZ(<%=rsCotizaciones("IDCOTIZACION")%>, this)'></td>					
					<%	elseif (CStr(rsCotizaciones("ESTADO")) = CTZ_EN_AJUSTE) then %>
						<td align="center"> <img title="<%=GF_TRADUCIR("En proceso de Ajuste")%>" id="ID_<%=rsCotizaciones("IDCOTIZACION")%>" src="images\compras\ajustes.gif" onclick='abrirREMPIC(<%=rsCotizaciones("IDCOTIZACION")%>)'></td>
					<%	else	%>
						<td align='center'>&nbsp;</td>
					<%	end if	
						'Si no está anulada, ni facturada y no se recibió mercadería, permite anular.
						if ((CStr(rsCotizaciones("ESTADO")) <> CTZ_ANULADA) and (isNull(rsCotizaciones("IDPIC"))) and (CDbl(rsCotizaciones("FACTURADO")) = 0)) then %>
						<td align="center"> <img title="<%=GF_TRADUCIR("Anular Pedido Interno")%>" id="ID_<%=rsCotizaciones("IDCOTIZACION")%>" src="images\compras\CTZ_cancel-16x16.png" onclick="anularCTZ(<%=rsCotizaciones("IDCOTIZACION")%>, '<%=rsCotizaciones("IDPEDIDO")%>', this)"></td>
					<%	else	%>
						<td align='center'>&nbsp;</td>
					<%	end if
						'Si no está anulada y se recibió mercadería, permite ver los remitos que tiene asocido.
						if (CStr(rsCotizaciones("ESTADO")) <> CTZ_ANULADA) then %>
						<td align="center"> <img title="<%=GF_TRADUCIR("Ver Cumplimiento")%>" style="cursor:pointer" id="ID_<%=rsCotizaciones("IDCOTIZACION")%>" src="images\search_b-16.png" onclick='abrirREMPIC(<%=rsCotizaciones("IDCOTIZACION")%>)'></td>
					<%	else	%>
						<td align='center'>&nbsp;</td>
					<%	end if  %>
					
				<%		
				'Controlar si corresponde mostrar nota de aceptación				
				if (mostrarNDA(rsCotizaciones)) then    %>		
					  <td align="center" colspan="2"> <img title="<%=GF_TRADUCIR("Administrar NDA")%>" style="cursor:pointer" id="NDA" src="images\compras\NDA-16x16.png" onclick='abrirNDA(<% =rsCotizaciones("IDPEDIDO") %>, <%=rsCotizaciones("IDCOTIZACION")%>)'></td>							
			 <% else %>
					<td colspan="2"></td> 		
			 <% end if%>
				<tr>
				<%
				rsCotizaciones.MoveNext()				
			wend 
		end if
		if (reg = 0) then
		%>
			<tr class="reg_Header_nav"><td style="text-align: center;" colSpan="12"><% =GF_TRADUCIR("No hay informacion disponible en estos momentos") %></td></tr>		
		<%  
		end if 
		%>			
		</tbody>
		<tfoot>
			<td colspan="12"><div id="paginacion"></div></td>
		</tfoot>	
	</table>
</body>
</html>
