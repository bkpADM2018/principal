<!--#include file="../Includes/includeGeneracionArchivos.asp"-->
<!--#include file="../Includes/procedimientosCompras.asp"-->
<!--#include file="../Includes/procedimientosFechas.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientosMail.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientoslog.asp"-->
<!--#include file="../Includes/procedimientosPDF.asp"-->
<!--#include file="../Includes/procedimientossql.asp"-->
<!--#include file="../Includes/procedimientostraducir.asp"-->
<!--#include file="interfacturas.asp"-->
<%

Const FECHA_MIGRACION_SISTEMA = 20170914

Function obtenerFacturas(p_ptoVenta, p_nroCbte, p_proveedor, p_nroCAE, p_fechaCAE, p_fechaPagoDesde, p_FechaPagoHasta, p_tipo, p_estado, p_sqlOrder, p_cdDivision, p_cdConcepto, p_fechaFacturaDesde, p_fechaFacturaHasta)
	dim strSQL, strWhere, rsFacturas , conn	
	strSQL = "	SELECT	CAB.guid AS nroReg, "&_			 
			 "			Year(CAB.feccbt)*10000+Month(CAB.feccbt)*100+Day(CAB.feccbt) AS FechaFac, "&_
			 "	       CAB.succbt AS ptoVenta, "&_
			 "	       CAB.nrocbt AS nroCbte, "&_
			 "	       CAB.cliente AS nroProveedor, "&_
			 "	       CAB.codmone AS cdMoneda, "&_
			 "	       CASE WHEN EMP.nomemp IS NULL THEN '' ELSE EMP.nomemp END AS DSProveedor, "&_			 
			 "	       CAB.tipcbt Tipo, "&_
			 "	       CAB.letra TipoFacturaAoB, "&_
			 "	       CAB.cai AS nroCAE, "&_
			 "	       Year(CAB.vencai)*10000+Month(CAB.vencai)*100+Day(CAB.vencai) AS fechaCAE, "&_
			 "		   CAB.codcia IDDIVISION," &_
			 "		   DIV.descia AS DSDIVISION, "&_
			 "         EMP.nrodoc  AS IdproveedorMail, " &_
			 "		   EMP.nrodoc AS lista, "&_			 
			 "		   CAB.imptotcbt AS IMPORTE "&_
			 "	FROM FAC001A CAB " &_
			 "		LEFT JOIN MET001A EMP ON EMP.nroemp = CAB.cliente " &_
			 "		LEFT JOIN CGT001A DIV ON DIV.cia = CAB.CODCIA "	
	if (p_ptoVenta <> "") then Call mkWhere(strWhere, "CAB.SUCCBT", Cint(p_ptoVenta), "=", 1)
	if (p_nroCbte <> "") then Call mkWhere(strWhere, "CAB.NROCBT", CDbl(p_nroCbte), "=", 3)
	if (p_proveedor <> 0) then Call mkWhere(strWhere, "CAB.cliente", p_proveedor, "=", 1) 
	if (p_tipo <> "") then Call getSearchTipoFAC(p_tipo, strWhere)
	if (p_nroCAE <> "") then Call mkWhere(strWhere, "CAB.cai", p_nroCAE, "LIKE", 3)
	if (p_fechaCAE <> "") then Call mkWhere(strWhere, "CAB.vencai", GF_FN2DTCONTBLE(p_fechaCAE), "=", 3)
	if (p_cdDivision <> "") then Call mkWhere(strWhere, "CAB.codcia", p_cdDivision, "=", 3)
	    if (p_fechaFacturaDesde <> 0 and p_fechaFacturaHasta <> 0) then  
        Call mkWhere(strWhere, "CAB.feccbt" ,GF_FN2DTCONTABLE(p_fechaFacturaDesde),">=",3)
        Call mkWhere(strWhere, "CAB.feccbt" ,GF_FN2DTCONTABLE(p_fechaFacturaHAsta),"<=",3)
    end if
	strSQL = strSQL & strWhere	
	strSQL = strSQL & " ORDER BY  CAB.feccbt desc, CAB.codcia , CAB.succbt, CAB.nrocbt desc"
	Call executeQueryDb(DBSITE_SQL_MAGIC, rsFacturas, "OPEN", strSQL)
	Set obtenerFacturas = rsFacturas
end Function
'********************************************************************
Function addParam(p_strKey,p_strValue,ByRef p_strParam)
       if (not isEmpty(p_strValue)) then
          if (isEmpty(p_strParam)) then
             p_strParam = "?"
          else
             p_strParam = p_strParam & "&"
          end if
          p_strParam = p_strParam & p_strKey & "=" & p_strValue
       end if
end Function
'********************************************************************
Function getSearchTipoFAC(pTipo, byRef strWhere)
	dim auxTipo, auxAoB	
		
	auxTipo = CInt(Left(pTipo, 1))
	auxAoB = Right(pTipo, 1)
		
	if (auxTipo <> 0) then Call mkWhere(strWhere, "CAB.TIPCBT", auxTipo, "=", 1)
	if (auxAoB <> "") then Call mkWhere(strWhere, "CAB.LETRA", auxAoB, "=", 0)
	
end Function
'**********************************************************
'*****************  INICIO DE LA PAGINA  ******************
'**********************************************************
dim ptoVenta, nroCBTE, FAC_nroCAE, fechaCAE, tipofac, aux_FechaFac, fechaPagoDesde, fac_fechaPagoDesde, fechaPagoHasta, fac_fechaPagoHasta, fechaFacturaDesde, fechaFacturaHasta ,fac_fechaFacturaDesde,fac_fechaFacturaHasta
dim rsFacturas, paginaActual, mostrar, contReg, fac_fechaCAE, fac_estado, aux_permisoEXP
dim aux_nroReg, aux_nroFactura, aux_nroCAE, aux_fechaCAE, aux_tipofac, aux_estado
dim FAC_accion,fac_Division,responsable,cdResponsable,fac_concepto, fac_sector, myUsuarioConsulta

FAC_accion = GF_PARAMETROS7("FAC_accion","",6)
call addParam("FAC_accion", FAC_accion, params)
ptoVenta = GF_PARAMETROS7("ptoVenta","",6)
call addParam("ptoVenta", ptoVenta, params)
nroCBTE = GF_PARAMETROS7("nroCBTE","",6)
call addParam("nroCBTE", nroCBTE, params)
FAC_nroCAE = GF_PARAMETROS7("FAC_nroCAE","",6)
call addParam("FAC_nroCAE", FAC_nroCAE, params)
fechaCAE = GF_PARAMETROS7("fechaCAE","",6)
call addParam("fechaCAE", fechaCAE, params)
fac_fechaCAE = GF_DTE2FN(fechaCAE)

fechaPagoDesde = GF_PARAMETROS7("fechaPagoDesde","",6)
call addParam("fechaPagoDesde", fechaPagoDesde, params)
fac_fechaPagoDesde = GF_DTE2FN(fechaPagoDesde)
fechaPagoHasta = GF_PARAMETROS7("fechaPagoHasta","",6)
call addParam("fechaPagoHasta", fechaPagoHasta, params)
fac_fechaPagoHasta = GF_DTE2FN(fechaPagoHasta)
fechaFacturaDesde = GF_PARAMETROS7("fechaFacturaDesde","",6)
call addParam("fechaFacturaDesde", fechaFacturaDesde, params)
fac_fechaFacturaDesde = GF_DTE2FN(fechaFacturaDesde)
fechaFacturaHasta = GF_PARAMETROS7("fechaFacturaHasta","",6)
call addParam("fechaFacturaHasta", fechaFacturaHasta, params)
fac_fechaFacturaHasta = GF_DTE2FN(fechaFacturaHasta)

tipofac = GF_PARAMETROS7("tipofac", "",6)
call addParam("tipofac", tipofac, params)
fac_sector = GF_PARAMETROS7("fac_sector",0,6)
call addParam("fac_sector", fac_sector, params)
fac_estado = GF_PARAMETROS7("fac_estado",0,6)

sql_Campo_Order = GF_PARAMETROS7("sqlCampoOrder",0,6)   
call addParam("sqlCampoOrder", sql_Campo_Order, params) 
sql_Tipo_Order = GF_PARAMETROS7("sqlTipoOrder","",6)    
call addParam("sqlTipoOrder", sql_Tipo_Order, params)   

if not isToepfer(session("KCOrganizacion")) then 
	myUsuarioConsulta = FAC_USER_WEB
	fac_estado = FAC_AUTORIZADA
	if trim(session("KCOrganizacion")) = "" then 
		Response.Write "Sesi&oacuten vencida. Por favor inicie sesión nuevamente"
		Response.End 
	end if	
	responsable = getDescripcionProveedor(session("KCOrganizacion"))
	cdResponsable = session("KCOrganizacion")
else
	myUsuarioConsulta = session("Usuario")
	responsable = GF_PARAMETROS7("responsable","",6)
	cdResponsable = GF_PARAMETROS7("cdResponsable",0,6)
	if responsable = "" then cdResponsable = 0
end if
call addParam("responsable", responsable, params)
call addParam("cdResponsable", cdResponsable, params)

'if (fac_estado = 0) then fac_estado = FAC_PENDIENTE
call addParam("fac_estado", fac_estado, params)
fac_Division = GF_PARAMETROS7("fac_Division","",6)
call addParam("fac_Division", fac_Division, params)

fac_concepto = GF_PARAMETROS7("fac_concepto","",6)
call addParam("fac_concepto", fac_concepto, params)
paginaActual = GF_PARAMETROS7("numeroPagina",0,6)
if (paginaActual = 0) then paginaActual=1
mostrar = GF_PARAMETROS7("registrosPorPagina",0,6)
if (mostrar = 0) then mostrar = 50

GP_ConfigurarMomentos
set rsFacturas = obtenerFacturas(ptoVenta, nroCBTE, cdResponsable, FAC_nroCAE, fac_fechaCAE, fac_fechaPagoDesde, fac_fechaPagoHasta, tipofac, fac_estado, sqlOrder,fac_Division,fac_concepto, fac_fechaFacturaDesde,fac_fechaFacturaHasta)
Call setupPaginacion(rsFacturas, paginaActual, mostrar)

%>
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
	<title>SISTEMA DE FACTURACION - CONSULTA DE FACTURAS</title>
<link rel="stylesheet" href="../css/main.css" type="text/css">
<link rel="stylesheet" href="../css/paginar.css" type="text/css">
<link rel="stylesheet" href="../css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
<link rel="stylesheet" href="../css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="../css/calendar-win2k-2.css" type="text/css">
<style type="text/css">
.divOculto {
	display: none;
}
.titleStyle {
	font-weight: bold;
	font-size: 20px;
}
</style>
	<script type="text/javascript" src="../scripts/paginar.js"></script>
	<script type="text/javascript" src="../scripts/date.js"></script>
	<script type="text/javascript" src="../scripts/formato.js"></script>
    <script type="text/javascript" src="../scripts/Toolbar.js"></script>
	<script type="text/javascript" src="../scripts/channel.js"></script>
	<script type="text/javascript" src="../scripts/calendar.js"></script>
	<script type="text/javascript" src="../scripts/calendar-1.js"></script>
   	<script type="text/javascript" src="../scripts/jquery/jquery-1.3.2.min.js"></script>
	<script defer type="text/javascript" src="../scripts/pngfix.js"></script>
	<script type="text/javascript" src="../scripts/jqueryPopUp.js"></script>
	<script type="text/javascript" src="../scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>	
	<script type="text/javascript">
	
		var NBR_TOKEN = "_";
		
		jQuery.fn.getCheckboxValues = function(){
			var values = "";
			var i = 0;
			this.each(function(){
				if (this.name != "todos")
				{
					if (values == "")
						values += $(this).val();
					else
						values += NBR_TOKEN + $(this).val();
				}
			});
			return values;
		}
		
		function bodyOnLoad() {			
			var	tb = new Toolbar('toolbar');
			tb.addButton("toolbar-print", "Imprimir Seleccionadas", "printFacturaLote()");
			tb.addButton("../../images/mail-16.png", "Administrar Mail", "abrirDirecciones()");
			tb.addButton("../../images/document-16.png", "Reportes", "VerificarDestinoReporte()");
			//tb.addButton("../../images/document-16.png", "Generar Archivo Auto", "generateFileAuto()");
			tb.draw();
			autoCompleteEmpresaResponsable();	
			<% 	if (not rsFacturas.eof) then %>
					var pgn = new Paginacion("paginacion");							
					pgn.paginar(<% =paginaActual %>, <% =rsFacturas.RecordCount %>, <% =mostrar %>, 50, "interfacturasConsulta.asp<% =params %>");					
			<%	end if 	%>
			pngfix();
		}

		function CerrarCal(cal) {
			cal.hide();
		}
	
		function MostrarCalendario(p_objID, funcSel) {
			var dte= new Date();		    	    
			var elem= document.getElementById(p_objID);
			if (calendar != null) calendar.hide();		
			var cal = new Calendar(false, dte, funcSel, CerrarCal);
		    cal.weekNumbers = false;
			cal.setRange(1993, 2045);
			cal.create();
			calendar = cal;		
		    calendar.setDateFormat("dd/mm/y");
		    calendar.showAtElement(elem);
		}
		function VerificarDestinoReporte(){
		    <% if myUsuarioConsulta = FAC_USER_WEB then %>
                generateFile();
		    <%else%>
                generateReport();
            <%end if%>
		}
	
		function resetFechas(pOption) {
			if (pOption == '1') {
				document.getElementById("fechaCAEDiv").innerHTML = "";
				document.getElementById("fechaCAE").value = "";
			}
			else{
			    if (pOption == '2'){
			        document.getElementById("fechaPagoDesdeDiv").innerHTML = "";
			        document.getElementById("fechaPagoDesde").value = "";
			        document.getElementById("fechaPagoHastaDiv").innerHTML = "";
			        document.getElementById("fechaPagoHasta").value = "";	
			    }else{
			        
			        document.getElementById("fechaFacturaDesdeDiv").innerHTML = "";
			        document.getElementById("fechaFacturaDesde").value = "";
			        document.getElementById("fechaFacturaHastaDiv").innerHTML = "";
			        document.getElementById("fechaFacturaHasta").value = "";
			    }

			}
		}
		function SeleccionarCalDesde(cal, date) {
			var str= new String(date);		
			document.getElementById("fechaPagoDesdeDiv").innerHTML = str;
			document.getElementById("fechaPagoDesde").value = str;
			if (cal) cal.hide();	
			var text = document.getElementById("fechaPagoHasta").value;
			if (text.trim() == ""){
				document.getElementById("fechaPagoHastaDiv").innerHTML = document.getElementById("fechaPagoDesdeDiv").innerHTML + " <a href=javascript:resetFechas('2')><img src='../images/button_cancel.png'></a>";
				document.getElementById("fechaPagoHasta").value = document.getElementById("fechaPagoDesde").value;
			}else{
				//Comparar que no sea mayor a fecha hasta
				if((Date.parse(str)) > (Date.parse(text))){
					alert("Fecha desde no puede ser mayor a fecha hasta!")
					resetFechas('2');
				}
			}
		}

		function SeleccionarCalHasta(cal, date) {
			var str= new String(date);
			document.getElementById("fechaPagoHastaDiv").innerHTML = str;
			document.getElementById("fechaPagoHasta").value = str;
			if (cal) cal.hide();
			var text = document.getElementById("fechaPagoDesde").value;
			if (text.trim() == ""){
				document.getElementById("fechaPagoDesdeDiv").innerHTML = document.getElementById("fechaPagoHastaDiv").innerHTML;
				document.getElementById("fechaPagoDesde").value = document.getElementById("fechaPagoHasta").value;
			}else{
				//Comparar que no sea menor a fecha hasta
				if((Date.parse(str)) < (Date.parse(text))){
					alert("Fecha hasta no puede ser menor a fecha desde!")
					resetFechas('2');
					return 0;
				}
			}		
			document.getElementById("fechaPagoHastaDiv").innerHTML = document.getElementById("fechaPagoHastaDiv").innerHTML + " <a href=javascript:resetFechas('2')><img src='../images/button_cancel.png'></a>";
		}
		function SelecCalDesdeFacturas(cal, date) {
		    var str= new String(date);		
		    document.getElementById("fechaFacturaDesdeDiv").innerHTML = str;
		    document.getElementById("fechaFacturaDesde").value = str;
		    if (cal) cal.hide();	
		    var text = document.getElementById("fechaFacturaHasta").value;
		    if (text.trim() == ""){
		        document.getElementById("fechaFacturaHastaDiv").innerHTML = document.getElementById("fechaFacturaDesdeDiv").innerHTML + " <a href=javascript:resetFechas('3')><img src='../images/button_cancel.png'></a>";
		        document.getElementById("fechaFacturaHasta").value = document.getElementById("fechaFacturaDesde").value;
		    }else{
		        //Comparar que no sea mayor a fecha hasta
		        if((Date.parse(str)) > (Date.parse(text))){
		            alert("Fecha desde no puede ser mayor a fecha hasta!")
		            resetFechas('3');
		        }
		    }
		}
		function SelecCalHastaFacturas(cal, date) {
		    var str= new String(date);
		    document.getElementById("fechaFacturaHastaDiv").innerHTML = str;
		    document.getElementById("fechaFacturaHasta").value = str;
		    if (cal) cal.hide();
		    var text = document.getElementById("fechaFacturaDesde").value;
		    if (text.trim() == ""){
		        document.getElementById("fechaFacturaDesdeDiv").innerHTML = document.getElementById("fechaFacturaHastaDiv").innerHTML;
		        document.getElementById("fechaFacturaDesde").value = document.getElementById("fechaFacturaHasta").value;
		    }else{
		        //Comparar que no sea menor a fecha hasta
		        if((Date.parse(str)) < (Date.parse(text))){
		            alert("Fecha hasta no puede ser menor a fecha desde!")
		            resetFechas('3');
		            return 0;
		        }
		    }		
		    document.getElementById("fechaFacturaHastaDiv").innerHTML = document.getElementById("fechaFacturaHastaDiv").innerHTML + " <a href=javascript:resetFechas('3')><img src='../images/button_cancel.png'></a>";
		}

		function SeleccionarCal(cal, date) {
			var str= new String(date);		
			document.getElementById("fechaCAEDiv").innerHTML = str + "<a href=javascript:resetFechas('1')><img src='../images/button_cancel.png'></a>";
			document.getElementById("fechaCAE").value = str;
			if (cal) cal.hide();	
		}

		function printFactura(nroreg, isOld) {			
			if (isOld == 0) {
				var lote = "";				
				//SI NO TIENE CODIGO DE CONCEPTO SE GENERA LA IMPRESION DE EXPORTACION
				window.open("interfacturasPrintCommon.asp?lote=" + nroreg + "&tipoPDF=<%=PDF_STREAM_MODE%>","_blank");
			} else {				
				window.open("archive/" + editarCaracteres(""+ nroreg, "0", 8, CHR_FWD) + ".pdf","_blank");
			}
		}
		
		function printFacturaLote() {
		    var lote = "";
			//Primero verifico si hay algun documento que se pidió imprimir que aún no tiene CAE.			
			var vals = $("input:checked").getCheckboxValues();
			var controlarCAE = true;
			arr = vals.split(NBR_TOKEN);
			for(x in arr) {
				lote = lote + document.getElementById("NFAC" + arr[x]).value + "','";	
			}
			//Quito la ultima coma
			if (lote != "") lote = lote.substr(0, lote.length-3);
			//Si esta todo bien, imprime.
			if (lote != "")  window.open("interfacturasPrintCommon.asp?lote=" + lote + "&tipoPDF=<%=PDF_STREAM_MODE%>","_blank");			
		}		
			
		function buscar() {
			document.getElementById("frmSel").submit();
		}
		
		
		function seleccionar_todo(me)
		{
			for (i=0;i<document.frmSel.elements.length;i++){
			
			
				if(document.frmSel.elements[i].type == "checkbox")
					if (document.frmSel.elements[i].name != "todos")
					{
						if (me.checked == 0)
							document.frmSel.elements[i].checked=0;
						else
							document.frmSel.elements[i].checked=1;
					}
			}
		}
 
		function autoCompleteEmpresaResponsable()
		{
			$( "#responsable" ).autocomplete({
					minLength: 2,
					source: "../comprasStreamElementos.asp?tipo=JQEmpresas",
					focus: function( event, ui ) {
						$( "#responsable").val(ui.item.dsempresa);
						return false;
					},
					select: function( event, ui ) {
						$( "#responsable"    ).val (ui.item.dsempresa);
						$( "#cdResponsable"    ).val (ui.item.idempresa);
						return false;
					},
					change: function( event, ui ) {
						if (!ui.item) {
							$( "#responsable").val ("");
							$( "#cdResponsable").val ("");
						}
					}
				})
				.data( "autocomplete" )._renderItem = function( ul, item ) {
					return $( "<li></li>" )
						.data( "item.autocomplete", item )
						.append( "<a>" + item.idempresa + " - <font style='font-size:10;'>" + item.dsempresa + "</font></a>" )
						.appendTo( ul );
				};
		}	
		
		function abrirDirecciones(){
			window.open("interfacturasAdministrarMail.asp", "_blank", "location=no,menubar=no,statusbar=no,scrolling=yes,height=500,width=700",false);				
		}						
		function sendFacturaByMail(idFactura,pIdProveedor,pLista){
			window.open("interfacturasAdministrarMail.asp?idProveedor="+pIdProveedor+"&factura="+idFactura+"&tipo="+pLista+"&accion=<%=ACCION_EMAIL%>", "_blank", "location=no,menubar=no,statusbar=no,scrolling=yes,height=500,width=700",false);
		}
		function generateFile(){
			window.open("interfacturasGenerarArchivo.asp","_blank","toolbar=no, scrollbars=yes, resizable=yes, width=750, height=650");
		}	
		function generateReport(){
		    var puw = new winPopUp('popupReportes','interfacturasReportesPopUp.asp','520','465', 'Reportes' ,' ' );
		}
		function verFirmasRegistradas(idFactura,cdComprobante){
			var puw = new winPopUp('popupFirmasReg',"interfacturasFirmaPopUp.asp?idFactura="+idFactura+"&cdComprobante="+cdComprobante,'500','275','Firmas Registradas', '');
		}
		function refacturar(idFactura,cdComprobante){
		    if(confirm("Desea refacturar este comprobante?? (IMPORTANTE: La confección de la nota de credito deberá ser realizada por el usuario y no se generará en este proceso.)")) {
		        var puw = new winPopUp('popupFirmasReg',"../sincronizarDescargasRefacturacion.asp?nreg="+idFactura,'800','800','Refacturacion', '');
		    }
			
		}
		if(typeof String.prototype.trim !== 'function') {
			String.prototype.trim = function() {
			return this.replace(/^\s+|\s+$/g, ''); 
		}}

		function setOrder(p_campo,p_orden){ 
		    document.getElementById("sqlCampoOrder").value = p_campo;
		    document.getElementById("sqlTipoOrder").value = p_orden;
		    submitInfo();
		}
		function submitInfo(acc) {		
		    document.getElementById("frmSel").submit();
		}
	window.onload=bodyOnLoad;		
	</script>
</head>
<body>
<div id="toolbar"></div><br>
<form id="frmSel" name="frmSel" action="InterfacturasConsulta.asp" method="get">
        <input type="hidden" name="sqlCampoOrder" id="sqlCampoOrder" value="<%=sql_Campo_Order%>"> 
	    <input type="hidden" name="sqlTipoOrder" id="sqlTipoOrder" value="<%=sql_Tipo_Order%>">    


		<div class="tableaside size100"> <!-- BUSCAR -->

			<h3> Filtros </h3>
		  
			<div id="searchfilter" class="tableasidecontent">
		        <div class="col16 reg_header_navdos"> <%=GF_Traducir("Nº Comprobante:")%> </div>
		        <div class="col16"> 
					<input type="text" size="2" maxlength="4" id="ptoVenta" name="ptoVenta" value="<% =ptoVenta %>">
					-
					<input type="text" size="8" maxlength="8" id="nroCBTE" name="nroCBTE" value="<% =nroCBTE %>">
		        </div>
		        		        
		        <div class="col16 reg_header_navdos"> <%=GF_Traducir("Proveedor:")%> </div>
		        <div class="col16"> 
					<% 
						pTipo = "text"
						if not isToepfer(session("KCOrganizacion")) then 
							Response.Write left(cdResponsable & "-" & responsable,22) 
							pTipo = "hidden"
						end if
					%>
       				<input name="responsable" size="20" type="<%=pTipo%>" id="responsable" value="<%=responsable%>">
					<input type="hidden" name="cdResponsable" id="cdResponsable" value="<%=cdResponsable%>">	
		        </div>
		        
		        <div class="col16 reg_header_navdos"> <%=GF_Traducir("Nº CAE:")%> </div>
		        <div class="col16"> 
					<input type="text" size="21" maxlength="14" id="FAC_nroCAE" name="FAC_nroCAE" value="<% =FAC_nroCAE %>">
		        </div>
		        
		        <div class="col16 reg_header_navdos"> <%=GF_Traducir("Tipo de Comprobante:")%> </div>
		        <div class="col16"> 
					<select id="tipofac" name="tipofac">
						<option value="" <% if (tipofac = "") then response.write "selected" %>><% =GF_TRADUCIR("- Todas -") %>
						<option value="<% =CODIGO_CBTE_FAC_A   %>" <% if (tipofac = CODIGO_CBTE_FAC_A)   then response.write "selected" %>><% =GF_TRADUCIR("Factura A") %>
						<option value="<% =CODIGO_CBTE_FAC_B   %>" <% if (tipofac = CODIGO_CBTE_FAC_B)   then response.write "selected" %>><% =GF_TRADUCIR("Factura B") %>
						<option value="<% =CODIGO_CBTE_FAC_E   %>" <% if (tipofac = CODIGO_CBTE_FAC_E)	 then response.write "selected" %>><% =GF_TRADUCIR("Invoice") %>
						<option value="<% =CODIGO_CBTE_NCR_A   %>" <% if (tipofac = CODIGO_CBTE_NCR_A)   then response.write "selected" %>><% =GF_TRADUCIR("Nota de Credito A") %>
						<option value="<% =CODIGO_CBTE_NCR_B   %>" <% if (tipofac = CODIGO_CBTE_NCR_B)   then response.write "selected" %>><% =GF_TRADUCIR("Nota de Credito B") %>
						<option value="<% =CODIGO_CBTE_NCR_E   %>" <% if (tipofac = CODIGO_CBTE_NCR_E)   then response.write "selected" %>><% =GF_TRADUCIR("Nota de Credito Invoice") %>
						<option value="<% =CODIGO_CBTE_NDB_A   %>" <% if (tipofac = CODIGO_CBTE_NDB_A)   then response.write "selected" %>><% =GF_TRADUCIR("Nota de Debito A") %>
						<option value="<% =CODIGO_CBTE_NDB_B   %>" <% if (tipofac = CODIGO_CBTE_NDB_B)   then response.write "selected" %>><% =GF_TRADUCIR("Nota de Debito B") %>			
						<option value="<% =CODIGO_CBTE_NDB_B   %>" <% if (tipofac = CODIGO_CBTE_NDB_E)   then response.write "selected" %>><% =GF_TRADUCIR("Nota de Debito Invoice") %>			
					</select>
				</div>
		        
		        <div class="col16 reg_header_navdos"> <%=GF_Traducir("Fecha CAE:")%> </div>
		        <div class="col16"> 
   					<table>
						<tr>
							<td>
								<a href="javascript:MostrarCalendario('img_fechaCAE', SeleccionarCal)">
									<img id="img_fechaCAE" src="../images/calendar-16.png">
								</a>
							</td>	
							<td>
								<div id="fechaCAEDiv">
									<%
									if fechaCAE <> "" then
										Response.Write fechaCAE 
										Response.Write "<a href=javascript:resetFechas('1')><img src='../images/button_cancel.png'></a>"
									end if
									%>	
								</div>
							</td>	
						</tr>	
						<input type="hidden" id="fechaCAE" name="fechaCAE" value="<% =fechaCAE %>">
					</table>
		        </div>
		        
		        <div class="col16 reg_header_navdos"> <%=GF_Traducir("Division:")%> </div>
		        <div class="col16"> 
					<%	Call executeQueryDb(DBSITE_SQL_MAGIC, rsDiv, "OPEN", "Select * from CGT001A") %>
					<select style="z-index:-1;" name="fac_Division" id="fac_Division">
					    <option SELECTED value="">- <% =GF_TRADUCIR("Seleccione") %> -
					<%	while (not rsDiv.eof)
							selected = ""						
							if (Cstr(rsDiv("cia")) = Cstr(fac_Division)) then selected = "selected" %>
							<option value="<% =rsDiv("cia") %>" <% =selected %>><% =rsDiv("descia") %>
					<%		rsDiv.MoveNext()
						wend	%>
					</select>
		        </div>			        
		        <!--
		        <div class="col16 reg_header_navdos"> <%=GF_Traducir("Por Fecha de Cobro:")%> </div>
		        <div class="col16"> 
   					<table>
						<tr>
							<td>
								<a href="javascript:MostrarCalendario('img_fechaPagoDesde', SeleccionarCalDesde)">
										<img id="img_fechaPagoDesde" src="../images/calendar-16.png">
								</a>			
								<span id="fechaPagoDesdeDiv"><% =fechaPagoDesde %></span>
								<input type="hidden" id="fechaPagoDesde" name="fechaPagoDesde" value="<% =fechaPagoDesde %>">
								<b><%=GF_Traducir("Al")%></b>
								<a href="javascript:MostrarCalendario('img_fechaPagoHasta', SeleccionarCalHasta)">
									<img id="img_fechaPagoHasta" src="../images/calendar-16.png">
								</a>			
								<span id="fechaPagoHastaDiv"><% =fechaPagoHasta %>
									<%if fechaPagoHasta <> "" then %>
										<a title="Borrar Fechas" href="javascript:resetFechas('2')"><img id="img_resetFechas" src="../images/button_cancel.png"></a>
									<%end if%>
								</span>
								<input type="hidden" id="fechaPagoHasta" name="fechaPagoHasta" value="<% =fechaPagoHasta %>">
							</td>

						</tr>	
					</table>
		        </div>		
				-->
                <div class="col16 reg_header_navdos"> <%=GF_Traducir("Por Fecha de Factura:")%> </div>
                 <div class="col16"> 
   					<table>
						<tr>
							<td>
								<a href="javascript:MostrarCalendario('img_fechafacturacionDesde', SelecCalDesdeFacturas)">
										<img id="img_fechafacturacionDesde" src="../images/calendar-16.png">
								</a>			
								<span id="fechaFacturaDesdeDiv"><% =fechaFacturaDesde %></span>
								<input type="hidden" id="fechaFacturaDesde" name="fechaFacturaDesde" value="<% =fechaFacturaDesde %>">
								<b><%=GF_Traducir("Al")%></b>
								<a href="javascript:MostrarCalendario('img_fechafacturacionHasta', SelecCalHastaFacturas)">
									<img id="img_fechafacturacionHasta" src="../images/calendar-16.png">
								</a>			
								<span id="fechaFacturaHastaDiv"><% =fechaFacturaHasta %>
									<%if fechaFacturaHasta <> "" then %>
										<a title="Borrar Fechas" href="javascript:resetFechas('3')"><img id="img_resetFechas" src="../images/button_cancel.png"></a>
									<%end if%>
								</span>
								<input type="hidden" id="fechaFacturaHasta" name="fechaFacturaHasta" value="<% =fechaFacturaHasta %>">
							</td>

						</tr>	
					</table>
		        </div>
		        <BR>
		        <div class="col26 reg_header_navdos"> <%=GF_Traducir("*Por defecto se muestran los pendientes")%> </div>
		        
		    	<span class="btnaction"><input type="submit"  onclick='buscar()' value="<% =GF_TRADUCIR("Buscar") %>" id=submitir name=submitir></span>
			</div>
		</div>
		<br>
	<table class="datagrid" width="100%" align="center">
	    <thead>
	        <tr>
				<th class="thiconac" align="center" >			
					<% =GF_TRADUCIR("Incluir") %>
					<input style="border:none;background-color:#2E6B4D;cursor:pointer;" type="checkbox" name="todos" id="todos" title="Seleccionar todo" onclick="seleccionar_todo(this)">
				</th>
				<th class="thiconac" align="center" >
					<% =GF_TRADUCIR("Fecha") %>
				</th>
				<th class="thiconac" align="center" >
					<% =GF_TRADUCIR("Divisi&oacute;n") %>
				</th>
				<th class="thiconac" align="center" >
					<% =GF_TRADUCIR("Nº Comprobante") %>
				</th>
				<th class="thiconac" align="center" >
					<% =GF_TRADUCIR("Descripción Proveedor")%>
				</th>
				<th class="thiconac" align="center" >
					<% =GF_TRADUCIR("Tipo Comprobante") %>
				</th>
				<!--
				<th class="thiconac" align="center" >
					<% =GF_TRADUCIR("Division") %>
				</th>
				-->
				<th class="thiconac" align="center" >
					<% =GF_TRADUCIR("Nº CAE") %>
				</th>
				<th class="thiconac" align="center" >
					<% =GF_TRADUCIR("Fecha CAE") %>
				</th>
				<th class="thiconac" align="center" >
					<% =GF_TRADUCIR("Importe") %>
				</th>				
				<th class="thiconac" align="center" >
					.
				</th>
			<tr>
		</thead>
		<tbody>	
			<%
			if (rsFacturas.eof)then 
			%>
			<tr class="reg_Header_nav">
				<td align="center" colspan="13"><% =GF_TRADUCIR("No hay informacion disponible en estos momentos") %>			</td>
			</tr>
			<%
			else
				while ((not rsFacturas.eof)	and (contReg < mostrar))
					contReg = contReg + 1			
					flagOld = 0
					if (CLng(rsFacturas("FechaFac")) <= FECHA_MIGRACION_SISTEMA) then flagOld = 1
					%>
					<tr>
						<td align="center"> 
							<%	
								aux_nroReg = rsFacturas("nroReg")								
								aux_nroCAE = rsFacturas("nroCAE")
								if (flagOld = 0) then
							%>
							<input style="cursor:pointer;" type="checkbox" id="NFAC<% =aux_nroReg %>" name="NFAC<% =aux_nroReg %>" value="<% =aux_nroReg %>">							
							<%	end if 	%>
						</td>
						<td align="center" onclick="printFactura('<% =aux_nroReg %>', <% =flagOld %>)" style="cursor:pointer">
							<% 
								aux_FechaFac = GF_FN2DTE(rsFacturas("FechaFac")) 
								if (len(aux_FechaFac) < 8) then aux_FechaFac = "-" 
								Response.Write aux_FechaFac
							%>
						</td>
						<td align="center" onclick="printFactura('<% =aux_nroReg %>', <% =flagOld %>)" style="cursor:pointer">
							<% Response.Write Trim(rsFacturas("DSDivision")) %>
						</td>
						<td align="center" onclick="printFactura('<% =aux_nroReg %>', <% =flagOld %>)" style="cursor:pointer">
							<% 
								aux_nroFactura = GF_EDIT_CBTE(GF_nDigits(rsFacturas("ptoVenta"),4) & GF_nDigits(rsFacturas("nroCBTE"),8)) 
								Response.Write aux_nroFactura 
							%>
						</td>
					
						<td align="left" onclick="printFactura('<% =aux_nroReg %>', <% =flagOld %>)" style="cursor:pointer">
							<% =rsFacturas("DSProveedor") %>
						</td>
					
						<td align="center" onclick="printFactura('<% =aux_nroReg %>', <% =flagOld %>)" style="cursor:pointer">
							<% 
								aux_tipofac = getTipoFactura(rsFacturas("Tipo"))
								if (trim(rsFacturas("TipoFacturaAoB")) <> "") then aux_tipofac = aux_tipofac & " " & trim(rsFacturas("TipoFacturaAoB")) 
								Response.Write aux_tipofac 
							%>
						</td>
					
						<td align="center" onclick="printFactura('<% =aux_nroReg %>', <% =flagOld %>)" style="cursor:pointer">
							<% =aux_nroCAE %>
						</td>
						<td align="center" onclick="printFactura('<% =aux_nroReg %>', <% =flagOld %>)" style="cursor:pointer">
							<% 
								aux_fechaCAE = GF_FN2DTE(rsFacturas("fechaCAE")) 
								if (len(aux_fechaCAE) < 8) then aux_fechaCAE = "-" 
							    Response.Write aux_fechaCAE 
						   %>
						</td>						
							<td align="RIGHT" onclick="printFactura('<% =aux_nroReg %>', <% =flagOld %>)" style="cursor:pointer">						
						<%
								Response.write getSimboloMoneda(rsFacturas("cdMoneda")) & " " & GF_EDIT_DECIMALS(Cdbl(rsFacturas("IMPORTE"))*100,2)
							%>
						</td>

						<td align="center">						
							<a href="javascript:sendFacturaByMail('<% =aux_nroReg %>',<% = rsFacturas("IdproveedorMail")%>,'<% =rsFacturas("lista") %>');">						
								<img id="imgSendMail_<% =aux_nroReg %>" src="../images/mail-16.png" title="<% =GF_TRADUCIR("Enviar a proveedores") %>">
							</a>						
						</td>						
					</tr>
					<% rsFacturas.movenext %>
				<% Wend %>
		<% end if %>		
		<input type="hidden" id="FAC_accion" name="FAC_accion" value="">
	</table>
    <tfoot>
  		<td colspan="12"><div id="paginacion"></div></td>
  	</tfoot>	
	
</form>
</body>
</html>