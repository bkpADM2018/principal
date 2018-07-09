<!--#include file="../Includes/procedimientostraducir.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientosmail.asp"-->
<!--#include file="../Includes/procedimientosunificador.asp"-->
<!--#include file="../Includes/procedimientosFechas.asp"-->
<!--#include file="../Includes/procedimientosparametros.asp"-->
<!--#include file="interfacturas.asp"-->
<% 

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
'******************************************************************************************************************
'********************************************	COMIENZO DE LA PAGINA   *******************************************
'******************************************************************************************************************
Dim idProveedor, dsProveedor, accion, totalRegistros,pagina,lpp,tipo,factura,flagHayListas
dim dtDesde, dtHasta, myListaMails

'session("KCOrganizacion")= 8148
if session("KCOrganizacion") <> CD_TOEPFER then
	dsProveedor = getDescripcionProveedor(session("KCOrganizacion"))
	idProveedor = session("KCOrganizacion")
else
	dsProveedor = GF_PARAMETROS7("dsProveedor","",6)
	idProveedor = GF_PARAMETROS7("idProveedor",0,6)
end if
Call addParam("dsProveedor", dsProveedor, params)
Call addParam("idProveedor", idProveedor, params)

Set sp_ret = executeSP(rs, "TFFL.TF220F4_GET_BY_IDPROVEEDOR_TIPO", idProveedor & "||" & FACTURACION_LISTA_MAIL_ARCHIVO & "||1||0$$totalRegistros")

tipo = GF_PARAMETROS7("tipo", "", 6)
call addParam("tipo", tipo, params)
accion = GF_PARAMETROS7("accion", "", 6)
call addParam("accion", accion, params)

dtDesde = GF_PARAMETROS7("dtDesde", "", 6)
if (dtDesde = "") then dtDesde = "01" & "/" & month(date) & "/" & year(date)
dtDesde = GF_STANDARIZAR_FECHA_RTRN(dtDesde) 
Call addParam("dtDesde", dtDesde, params)
dtHasta = GF_PARAMETROS7("dtHasta", "", 6)
if (dtHasta = "") then dtHasta = GF_STANDARIZAR_FECHA_RTRN(date) 
Call addParam("dtHasta", dtHasta, params)

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta http-equiv="X-UA-Compatible" content="IE=Edge">
<link rel="stylesheet" href="../css/main.css" type="text/css">
<link rel="stylesheet" href="../css/paginar.css" type="text/css">
<link rel="stylesheet" href="../css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="../css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
<script type="text/javascript" src="../scripts/paginar.js"></script>
<script type="text/javascript" src="../scripts/controles.js"></script>
<script type="text/javascript" src="../scripts/Toolbar.js"></script>
<script type="text/javascript" src="../scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="../scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript" src="../scripts/channel.js"></script>
<script type="text/javascript" src="../Scripts/jQueryPopUp.js"></script>
<link rel="stylesheet" href="../css/calendar-win2k-2.css" type="text/css">
<script type="text/javascript" src="../scripts/calendar.js"></script>
<script type="text/javascript" src="../scripts/calendar-1.js"></script>
<style type="text/css">
.myFont {
	font-size: 12;
}
</style>
<script type="text/javascript">
	var ch = new channel();
	var subIndice = 0;
	function bodyOnload(){
		autoCompleteEmpresaResponsable();
	}
	
    function loadPopUpNew_callback(){
		submitInfo();
    } 
    function submitInfo() {     
	    document.getElementById("frmSel").action = "interfacturasGenerarArchivo.asp";
	    document.getElementById("frmSel").target = "_parent";		
        document.getElementById("frmSel").submit();
    }

    function autoCompleteEmpresaResponsable(){
			$( "#dsProveedor" ).autocomplete({
					minLength: 3,
					source: "../comprasStreamElementos.asp?tipo=JQEmpresas",
					focus: function( event, ui ) {
						$( "#dsProveedor").val(ui.item.dsempresa);
						return false;
					},
					select: function( event, ui ) {
						$( "#dsProveedor"    ).val (ui.item.dsempresa);
						$( "#idProveedor"    ).val (ui.item.idempresa);
						submitInfo();
						return false;
					},
					change: function( event, ui ) {
						if (!ui.item) {
							$( "#dsProveedor").val ("");
							$( "#idProveedor").val ("");
							submitInfo();
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
	function SeleccionarCalDesde(cal, date) {
		var str= new String(date);		
		document.getElementById("dtDesdeDiv").innerHTML = str;
	    document.getElementById("dtDesde").value = str;
		if (cal) cal.hide();	
		var text = document.getElementById("dtHasta").value;
		if (text.trim() == ""){
			document.getElementById("dtHastaDiv").innerHTML = document.getElementById("dtDesdeDiv").innerHTML;
			document.getElementById("dtHasta").value = document.getElementById("dtDesde").value;
		}else{
			//Comparar que no sea mayor a fecha hasta
			var fechaDesde = FormatearFecha(str);
			var fechaHasta = FormatearFecha(text);
			if((Date.parse(fechaDesde)) > (Date.parse(fechaHasta))){
				alert("Fecha desde no puede ser mayor a fecha hasta!");
				document.getElementById("dtDesdeDiv").innerHTML = document.getElementById("dtHastaDiv").innerHTML;
				document.getElementById("dtDesde").value = document.getElementById("dtHasta").value;				
				}
		}
	}		
	function FormatearFecha(pDate){
		var arrayFecha;
		arrayFecha = pDate.split("/") 
		var dia=arrayFecha[0] 
		var mes=(arrayFecha[1]-1) 
		var ano=(arrayFecha[2])
		return new Date(ano,mes,dia)
	}
	function SeleccionarCalHasta(cal, date) {
		var str= new String(date);		
		document.getElementById("dtHastaDiv").innerHTML = str;
	    document.getElementById("dtHasta").value = str;
		if (cal) cal.hide();	
		var text = document.getElementById("dtDesdeDiv").innerHTML;
		if (text.trim() == ""){
			document.getElementById("dtDesdeDiv").innerHTML = document.getElementById("dtHastaDiv").innerHTML;
			document.getElementById("dtDesde").value = document.getElementById("dtHasta").value;
		}else{
			//Comparar que no sea mayor a fecha hasta
			var fechaDesde = FormatearFecha(str);
			var fechaHasta = FormatearFecha(text);
			if((Date.parse(fechaDesde)) < (Date.parse(fechaHasta))){
				alert("Fecha hasta no puede ser menor a fecha desde!");
				document.getElementById("dtHastaDiv").innerHTML = document.getElementById("dtDesdeDiv").innerHTML;
				document.getElementById("dtHasta").value = document.getElementById("dtDesde").value;				
				}
		}	
	}		
	function CerrarCal(cal) {
		cal.hide();
	}	
	function generateFile(pTipo, pEnviar){
		var myAccion;
		if (document.getElementById("idProveedor").value == 0){
			alert("Debe ingresar un proveedor!")
			return 0;
		}
		if (document.getElementById("dtDesde").value == ""){
			alert("Debe ingresar la fecha desde!")
			return 0;
		}
		if (document.getElementById("dtHasta").value == ""){
			alert("Debe ingresar la fecha hasta!")
			return 0;
		}
		if (pEnviar == '1'){
			myAccion = "<%=ACCION_BACH%>";
			habilitarLoading("visible","relative");
		}else{
			myAccion = "<%=ACCION_SUBMITIR%>";
		}
		document.getElementById("accion").value = myAccion;
		
		if (pTipo=="XLS"){
		    document.getElementById("frmSel").action = "interfacturasGenerarXLSAjax.asp";
		}
		else {
		    document.getElementById("frmSel").action = "interfacturasGenerarArchivoAjax.asp";
		}
		document.getElementById("frmSel").target = "ifrm1";
		//submitInfo();
		document.getElementById("frmSel").submit();
	}	
	function sinDatos(){
		alert("No se encontraron facturas para el proveedor y periodo seleccionado!");
	}
	function showHideTbl(pTblName){
		if (document.getElementById(pTblName).style.visibility == 'visible'){
			document.getElementById(pTblName).style.visibility = 'hidden';
			document.getElementById(pTblName).style.position = 'absolute';
		}else{
			document.getElementById(pTblName).style.visibility = 'visible';
			document.getElementById(pTblName).style.position = 'relative';
		}
		
	}
	function abrirDirecciones(){
			window.open("interfacturasAdministrarMail.asp?idProveedor="+ document.getElementById("idProveedor").value +"&tipo=<%=FACTURACION_LISTA_MAIL_ARCHIVO%>&accion=<%=ACCION_EMAIL%>", "_blank", "location=no,menubar=no,statusbar=no,scrolling=yes,height=500,width=700",false);
		}	
	if(typeof String.prototype.trim !== 'function') {
		String.prototype.trim = function() {
			return this.replace(/^\s+|\s+$/g, ''); 
		}
	}
	function editMail(pIdProveedor,pTipo,pOrden) {		
		document.getElementById("txtmail_" + pIdProveedor +"_"+pTipo+"_"+ pOrden).style.display = 'none';
		document.getElementById("inputMail_" + pIdProveedor +"_"+pTipo+"_"+ pOrden).style.display = 'block';		
		var auxTipo = "'"+pTipo+"'";
		document.getElementById("imgMail_" + pIdProveedor +"_"+pTipo+"_"+ pOrden).innerHTML = '<img src="../images/save-16.png" onClick="updateMail('+pIdProveedor+','+auxTipo+','+pOrden+')">';
    }
    function deleteMail(pIdProveedor,pTipo,pOrden){
		if(confirm("Desea eliminar el mail ?" )){
			ch.bind("interfacturasAdministrarMailAjax.asp?idProveedor=" + pIdProveedor + "&tipo=" +pTipo+ "&orden=" + pOrden + "&accion=<%=ACCION_BORRAR%>", "eliminarMail_CallBack()");
			ch.send();
		}
    }	
	function updateMail(pIdProveedor,pTipo,pOrden) {		
		var actualMail = document.getElementById("actualMail_" + pIdProveedor +"_"+ pTipo + "_"+ pOrden).value;
		var newMail = document.getElementById("inputMail_" + pIdProveedor +"_"+pTipo+"_"+ pOrden).value;
		if ((actualMail != newMail) && (newMail != '')) {
			document.getElementById("imgMail_" + pIdProveedor +"_"+pTipo+"_"+ pOrden).innerHTML = '<img src="../images/loading_small_green.gif">';
			ch.bind("interfacturasAdministrarMailAjax.asp?idProveedor=" + pIdProveedor + "&tipo="+ pTipo +"&orden=" + pOrden +"&accion=<%=ACCION_PROCESAR%>&mail="+newMail, "updateMailCallBack("+pIdProveedor+",'"+pTipo+"',"+pOrden+")");
			ch.send();
		} else {
			alert('No se produjo ningun cambio en el mail');
			updateChange(pIdProveedor,pTipo,pOrden, actualMail);
		}
	}    
	function updateMailCallBack(pIdProveedor,pTipo,pOrden) {
		var actualMail = document.getElementById("actualMail_" + pIdProveedor + "_" + pTipo + "_" + pOrden).value;
		var resp = ch.response();
		if (resp == '<%=RESPUESTA_OK%>') {
			submitInfo();
		} else {
			alert('El mail ingresado es invalido');
			updateChange(pIdProveedor,pTipo,pOrden, actualMail);
		}
	}	
    function eliminarMail_CallBack(){
		submitInfo();
    }   
	function habilitarLoading(pVisibility, pPosition){
		document.getElementById("imgLoading").style.position = pPosition;
		document.getElementById("imgLoading").style.visibility  = pVisibility;
		document.getElementById("lblLoading").style.position = pPosition;
		document.getElementById("lblLoading").style.visibility  = pVisibility;
	}  
    function addMail() {
		if (subIndice > 0) {
			alert("Primero debe guardar la direccion ingresada!")
		}else{
			<%if (idProveedor > 0) then %>
				 var className = "thicon";
				 $("#MAIL_TABLE")
						.find('tfoot:last')
				         .append($('<tr>')
				             .addClass(className)
				             .append($('<td align=left>')				                 
				                 .append($('<input type=\"text\">')				                     
				                     .attr('size', 40)
				                     .attr('id', 'myMail')
				                     .attr('name', 'myMail')
				                 )
				             )
				             .append($('<td align=center>')
									.append($('<img style=\"cursor:pointer;\" src=\"../images/save-16.png\" onClick=\"saveMail(\'myMail\')\">'))
				             )
				             .append($('<td align=center>'))
				         );
				$('table#MAIL_TABLE tr:last').after($('#ACTION_ROW'));     
				subIndice = subIndice + 1;  
			<% end if %> 
       }
    }	
	function saveMail(pDir){
		ch.bind("interfacturasAdministrarMailAjax.asp?accion=<%=ACCION_GRABAR%>&idProveedor=<%=idProveedor%>&tipo=<%=FACTURACION_LISTA_MAIL_ARCHIVO%>&mail=" + document.getElementById(pDir).value , "saveMailCallBack()");
		ch.send();
	}
	function saveMailCallBack(){	
		if (ch.response() == "<%=RESPUESTA_OK%>"){
			submitInfo();
		}else{
			alert(ch.response());
		}
	}
	function fileSent(){
		alert("Archivo enviado con exito!");
	}

</script>
<body onload="bodyOnload()">
	<form name="post" id="frmSel" name="frmSel" action="interfacturasGenerarArchivo.asp">		
		<div class="tableaside size100">
		<h3><%=GF_Traducir("Generacion de Reportes de Facturas")%></h3>
		</div>
			<div class="tableasidecontent">
				<div class="col16 reg_header_navdos"> <%=GF_Traducir("Proveedor:")%> </div>				
				<div class="col56">
					<% 
						pTipo = "text"
						if session("KCOrganizacion") <> CD_TOEPFER then
							Response.Write idProveedor & "-" & dsProveedor
							pTipo = "hidden"
						end if
					%>
       				<input name="dsProveedor" size=29 type="<%=pTipo%>" id="dsProveedor" value="<%=dsProveedor%>">
					<input type="hidden" name="idProveedor" id="idProveedor" value="<%=idProveedor%>">	
				</div>
				<div class="col16 reg_header_navdos"> <%=GF_Traducir("Periodo:")%> </div>
				<div class="col56"> 
   					<table>
						<tr>
							<td>
								<a href="javascript:MostrarCalendario('imgDesde', SeleccionarCalDesde)"><img id="imgDesde" src="../images/calendar-16.png"></a>
							</td>	
							<td>
								<div id="dtDesdeDiv"><% =dtDesde %></div>
								<input type="hidden" id="dtDesde" name="dtDesde" value="<%=dtDesde%>">
							</td>	
							<td>&nbsp;Al&nbsp;</td>
							<td>
								<a href="javascript:MostrarCalendario('imgHasta', SeleccionarCalHasta)"><img id="imgHasta" src="../images/calendar-16.png"></a>
							</td>	
							<td>
								<div id="dtHastaDiv"><% =dtHasta %></div>
								<input type="hidden" id="dtHasta" name="dtHasta" value="<%=dtHasta%>">
							</td>	

					</table>
				</div>			
				<div class="col16 reg_header_navdos"> <%=GF_Traducir("Mails Registrados:")%> </div>
				<div class="col76"> 
					<table class="datagrid" id="MAIL_TABLE" width="70%" align="left">
						<thead>
							<tr>
								<th class="thiconac" align="center" width="90%" nowrap>	<% =GF_TRADUCIR("Mail") %></th>
								<th class="thiconac" align="center" width="5%" nowrap><% =GF_TRADUCIR("Editar") %></th>
								<th class="thiconac" align="center" width="5%" nowrap><% =GF_TRADUCIR("Eliminar") %></th>
							</tr>
						</thead>	
						<tbody>
						
							<% if (not rs.eof) then 
									flagHayListas = true
									while (not rs.eof) %>			
										<tr>
											<td class="thicon">
												<div id="txtmail_<%=rs("IDPROVEEDOR") %>_<%=rs("TIPO")%>_<%=rs("ORDEN")%>"><% =rs("MAIL") %></div>							
												<input type="hidden" id="actualMail_<%=rs("IDPROVEEDOR") %>_<%=rs("TIPO")%>_<%=rs("ORDEN")%>" value="<% =rs("MAIL") %>">
												<input type="text" id="inputMail_<%=rs("IDPROVEEDOR") %>_<%=rs("TIPO")%>_<%=rs("ORDEN")%>" style="display:none;" size="40" value="<% =rs("MAIL") %>">
											</td>
											<td class="thicon" align="center"><div id="imgMail_<%=rs("IDPROVEEDOR") %>_<%=rs("TIPO")%>_<%=rs("ORDEN")%>" style="cursor:pointer;"><img src="../images/edit-16.png" onClick="editMail(<%=rs("IDPROVEEDOR")%>,'<%=rs("TIPO")%>',<%=rs("ORDEN")%>)" title="Editar"></div></td>
											<td class="thicon" align="center"><img src="../images/cross-16.png" onClick="deleteMail(<%=rs("IDPROVEEDOR")%>,'<%=rs("TIPO")%>',<%=rs("ORDEN")%>)" style="cursor:pointer;" title="Eliminar"> </td>
										</tr>
							<%		  rs.MoveNext()			
									wend	
							   else %>
									<tr>
										<td colspan="3" align="center"><%=GF_TRADUCIR("No se encontraron resultados")%></td>
									</tr>
							<% end if %>								
						</tbody>
						<tfoot id="last">
						<tr id="ACTION_ROW">
							<td colspan="3" align="right" >
								<a id="btnmore" class="btnmore" href="javascript:addMail('')"><img src="../images/plus-16.png"><%=GF_Traducir("Agregar Direccion")%></a>
						    </td>
						</tr>	
						</tfoot>	
					</table>		
				</div>			
				<div class="col16 reg_header_navdos"> <%=GF_Traducir("Archivos:")%> </div>				
					<div class="col76">
						<table class="datagrid" width="70%" align="left">
						<tbody>
							<tr>
								<td rowspan=2 width="10%" bgcolor=white>
									<img id="imgXLS" src="../images/excel-50.png">
								</td>
								<td>
									<a style="cursor:pointer;" onclick="generateFile('XLS','0')"><%=GF_Traducir("Descargar Reporte de Facturas Emitidas (Solo Acondicionamiento)")%></a>
								</td>
							</tr>	
							<tr>	
								<td>
									<a style="cursor:pointer;" onclick="generateFile('XLS','1')"><%=GF_Traducir("Enviar por Mail")%></a>
								</td>
							</tr>
							<tr><td colspan="2" bgcolor=white></td></tr>
							<tr>
								<td rowspan=3 bgcolor=white>
									<img id="imgXLS" src="../images/document-50.png">
								</td>
								<td>
									<a style="cursor:pointer;" onclick="generateFile('TXT','0')"><%=GF_Traducir("Descargar Archivo para importar en Sistema")%></a>
								</td>
							</tr>
							<tr>
								<td>
									<a style="cursor:pointer;" onclick="generateFile('TXT','1')"><%=GF_Traducir("Enviar por Mail")%></a>
								</td>
							</tr>
							<tr>
								<td>	
									<a style="cursor:pointer;" onclick="showHideTbl('tblEspecificaciones');"><%=GF_Traducir("Ver Formato")%></a>
								</td>								
							</tr>										
						</tbody>
					</table>	
				</div>	         
			</div>	
			</div>
		<!--</div>	-->
<div class="col66"></div>
<table align="center" width="90%" border="0">
	<tr>
		<td align="center">
			<img style="position:absolute;visibility:hidden;" id="imgLoading" src="../images/Loading4.gif">
			<div style="position:absolute;visibility:hidden;" id="lblLoading"><b><br>Aguarde por favor...</b></div>
					
		</td>
	</tr>
</table>    		
		<div class="col76">
		<table width="60%" align="center" cellpadding="1" cellspacing=0 id=tblEspecificaciones style="visibility:hidden;position:absolute;">
			<thead>
				<tr>
					<th class="thiconac myFont" align="center"><%=GF_Traducir("Descripción")%></th>
					<th class="thiconac myFont" align="center"><%=GF_Traducir("Tamaño")%></th>
					<th class="thiconac myFont" align="center"><%=GF_Traducir("Tipo")%></th>
				</tr>	
			</thead>
			<tbody>	
				<tr>						
					<td class="myFont"><%=GF_Traducir("Fecha de Comprobante")%></td>
					<td class="myFont" align="center"><%=tFECHA%></td>
					<td class="myFont" align="center"><%=GF_Traducir("Numerico")%></td>
				</tr>
				<tr bgcolor=LightGrey>						
					<td class="myFont"><%=GF_Traducir("Vencimiento de Comprobante")%></td>
					<td class="myFont" align="center"><%=tFECHA%></td>
					<td class="myFont" align="center"><%=GF_Traducir("Numerico")%></td>
				</tr>
				<tr>						
					<td class="myFont"><%=GF_Traducir("Tipo de Comprobante<BR>(1=Factura, 2=N/Débito, 3=N/Crédito)")%></td>
					<td class="myFont" align="center"><%=tTCBT%></td>
					<td class="myFont" align="center"><%=GF_Traducir("Numerico")%></td>
				</tr>
				<tr bgcolor=LightGrey>						
					<td class="myFont"><%=GF_Traducir("Tipo de Comprobante.<BR>(A=Factura A, B=Factura B)")%></td>
					<td class="myFont" align="center"><%=tTCBT%></td>
					<td class="myFont" align="center"><%=GF_Traducir("Texto")%></td>
				</tr>
				<tr>						
					<td class="myFont"><%=GF_Traducir("Número de Comprobante")%></td>
					<td class="myFont" align="center"><%=tNUM%></td>
					<td class="myFont" align="center"><%=GF_Traducir("Numerico")%></td>
				</tr>
				<tr bgcolor=LightGrey>						
					<td class="myFont"><%=GF_Traducir("CUIT Corredor")%></td>
					<td class="myFont" align="center"><%=tCUIT%></td>
					<td class="myFont" align="center"><%=GF_Traducir("Numerico")%></td>
				</tr>
				<tr>						
					<td class="myFont"><%=GF_Traducir("CUIT Vendedor")%></td>
					<td class="myFont" align="center"><%=tCUIT%></td>
					<td class="myFont" align="center"><%=GF_Traducir("Numerico")%></td>
				</tr>
				<tr bgcolor=LightGrey>						
					<td class="myFont"><%=GF_Traducir("CUIT Emisor")%></td>
					<td class="myFont" align="center"><%=tCUIT%></td>
					<td class="myFont" align="center"><%=GF_Traducir("Numerico")%></td>
				</tr>
				<tr>						
					<td class="myFont"><%=GF_Traducir("Cto Toepfer")%></td>
					<td class="myFont" align="center"><%=tCtos%></td>
					<td class="myFont" align="center"><%=GF_Traducir("Texto")%></td>
				</tr>
				<tr bgcolor=LightGrey>						
					<td class="myFont"><%=GF_Traducir("Cto Corredor")%></td>
					<td class="myFont" align="center"><%=tCtos%></td>
					<td class="myFont" align="center"><%=GF_Traducir("Texto")%></td>						
				</tr>
				<tr>
					<td class="myFont"><%=GF_Traducir("Moneda de la factura (P=Pesos, D=D&oacutelares") %></td>
					<td class="myFont" align="center">1</td>
					<td class="myFont" align="center"><%=GF_Traducir("Texto")%></td>						
				</tr>
				<tr bgcolor=LightGrey>
					<td class="myFont"><%=GF_Traducir("Monto Gravado (Centavos)")%></td>
					<td class="myFont" align="center"><%=tIMP%></td>
					<td class="myFont" align="center"><%=GF_Traducir("Numerico")%></td>						
				</tr>
				<tr>						
					<td class="myFont"><%=GF_Traducir("Monto No Gravado (Centavos)")%></td>
					<td class="myFont" align="center"><%=tIMP%></td>
					<td class="myFont" align="center"><%=GF_Traducir("Numerico")%></td>
				</tr>
				<tr bgcolor=LightGrey>
					<td class="myFont"><%=GF_Traducir("Tasa IVA (Centesimas)")%></td>
					<td class="myFont" align="center"><%=tPORC%></td>
					<td class="myFont" align="center"><%=GF_Traducir("Numerico")%></td>
				</tr>
				<tr >
					<td class="myFont"><%=GF_Traducir("Monto del IVA (Centavos)")%></td>
					<td class="myFont" align="center"><%=tIMP%></td>
					<td class="myFont" align="center"><%=GF_Traducir("Numerico")%></td>
				</tr>
				<tr bgcolor=LightGrey>
					<td class="myFont"><%=GF_Traducir("Percepcion IVA (Centavos)")%></td>
					<td class="myFont" align="center"><%=tIMP%></td>
					<td class="myFont" align="center"><%=GF_Traducir("Numerico")%></td>
				</tr>
				<tr >
					<td class="myFont"><%=GF_Traducir("Percepcion IIBB (Centavos)")%></td>
					<td class="myFont" align="center"><%=tIMP%></td>
					<td class="myFont" align="center"><%=GF_Traducir("Numerico")%></td>
				</tr>
				<tr bgcolor=LightGrey> 
					<td class="myFont"><%=GF_Traducir("Total Comprobante (Centavos)")%></td>
					<td class="myFont" align="center"><%=tIMP%></td>
					<td class="myFont" align="center"><%=GF_Traducir("Numerico")%></td>
				</tr>
				<tr >
					<td class="myFont"><%=GF_Traducir("Tipo de Cambio (Milesimas)")%></td>
					<td class="myFont" align="center"><%=tCambio%></td>
					<td class="myFont" align="center"><%=GF_Traducir("Numerico")%></td>
				</tr>
				<tr bgcolor=LightGrey>						
					<td class="myFont"><%=GF_Traducir("Nro de CAE")%></td>
					<td class="myFont" align="center"><%=tCAE%></td>
					<td class="myFont" align="center"><%=GF_Traducir("Numerico")%></td>
				</tr>
				<tr >
					<td class="myFont"><%=GF_Traducir("Vencimiento de CAE")%></td>
					<td class="myFont" align="center"><%=tFECHA%></td>
					<td class="myFont" align="center"><%=GF_Traducir("Numerico")%></td>
				</tr>	
				<tr bgcolor=LightGrey>
					<td class="myFont"><%=GF_Traducir("Número de Carta de Porte")%></td>
					<td class="myFont" align="center"><%=tCTAPTE%></td>
					<td class="myFont" align="center"><%=GF_Traducir("Numerico")%></td>
				</tr>
				<tr >
					<td class="myFont"><%=GF_Traducir("Kilos Descarga")%></td>
					<td class="myFont" align="center"><%=tKILOS%></td>
					<td class="myFont" align="center"><%=GF_Traducir("Numerico")%></td>
				</tr>
				<tr bgcolor=LightGrey> 
					<td class="myFont"><%=GF_Traducir("Kilos de Merma")%></td>
					<td class="myFont" align="center"><%=tKILOS%></td>
					<td class="myFont" align="center"><%=GF_Traducir("Numerico")%></td>
				</tr>
				<tr >
					<td class="myFont"><%=GF_Traducir("Porcentaje de Humedad (Centesimas)")%></td>
					<td class="myFont" align="center"><%=tPORC%></td>
					<td class="myFont" align="center"><%=GF_Traducir("Numerico")%></td>
				</tr>
				<tr bgcolor=LightGrey>
					<td class="myFont"><%=GF_Traducir("Tarifa del Gasto (Centavos)")%></td>
					<td class="myFont" align="center"><%=tIMP%></td>
					<td class="myFont" align="center"><%=GF_Traducir("Numerico")%></td>
				</tr>
				<tr >
					<td class="myFont"><%=GF_Traducir("Tipo de Gasto")%></td>
					<td class="myFont" align="center"><%=tTIPO%></td>
					<td class="myFont" align="center"><%=GF_Traducir("Numerico")%></td>
				</tr>
			</tbody>									
		</table>
		</div>
		<input type="hidden" value="<%=ACCION_SUBMITIR%>" name="accion" id="accion">
	</form>		
	<iframe width=1 height=1 onload="habilitarLoading('hidden','absolute');" id="ifrm1" name="ifrm1"></iframe>
</body>
</html>