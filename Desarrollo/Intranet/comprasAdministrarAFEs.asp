<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientossql.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosAFE.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<%
Call comprasControlAccesoCM(RES_AFE)
'**********************************************************
'***	COMIENZO DE PAGINA
'**********************************************************
Dim rsAFE,  cdPedido, paginaActual, mostrar, cdObra
Dim cdUsuario, dsUsuario, cantFiles, imgAFEUpload

'variables para la busqueda
Dim afeSearch_cdAFE,afeSearch_Division,afeSearch_Title,afeSearch_import, rsDivision
Dim afeSearch_radioImport,searching,afeSearch_idObra,dsObra,afeSearch_Order
Dim afeSearch_pedido,listaDivisiones,totalRegistros

paginaActual = GF_PARAMETROS7("numeroPagina"       ,0,6)
mostrar      = GF_PARAMETROS7("registrosPorPagina" ,0,6)

if (paginaActual = 0) then paginaActual = 1
if (mostrar      = 0) then mostrar      = 10

'parametros de busqueda
searching             = 0
afeSearch_cdAFE 	  = UCASE(GF_PARAMETROS7("cdAFE"		 ,"",6))
afeSearch_Division	  =       GF_PARAMETROS7("afeDivision"	 ,0,6)
afeSearch_Title 	  = UCASE(GF_PARAMETROS7("afeTitle"		 ,"",6))
afeSearch_import	  =  	  GF_PARAMETROS7("afeImport"	 ,"",6)
afeSearch_radioImport = 	  GF_PARAMETROS7("radio_Import"	 ,"",6)
afeSearch_cdObra	  = 	  GF_PARAMETROS7("cdObra" 		 ,"",6)
afeSearch_Order		  = 	  GF_PARAMETROS7("afeOrder"		 ,"",6)
searching   		  =		  GF_PARAMETROS7("busquedaActiva","",6)
afeSearch_pedido	  = UCASE(GF_PARAMETROS7("afepedido"     ,"",6))

if (afeSearch_Order = "") then afeSearch_Order = "ORDER BY AFE.IDDIVISION, AFE.CDAFE DESC" end if

if (afeSearch_cdObra <> "") then
	'Busco la descripcion de la obra para agregarla al MagicSearch
	strSQL =          " SELECT * "
	strSQL = strSQl & " FROM   TBLDATOSOBRAS "
	strSQL = strSQl & " WHERE  Cdobra = '" & afeSearch_cdObra & "'"
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	dsObra = UCASE(rs("DSOBRA"))
end if


Call GP_ConfigurarMomentos()


listaDivisiones = getListaCargosAdmin()
if (listaDivisiones = "") then listaDivisiones = "0"
multiplo = 100
if inStr(afeSearch_import,",") then multiplo = 1
sp_parameter = 	afeSearch_cdAFE & "||" & afeSearch_Division & "||" & afeSearch_Title & "||" & afeSearch_cdObra & "||" & afeSearch_pedido & "||" & afeSearch_import * multiplo & "||" & afeSearch_radioImport & "||" & listaDivisiones &	"||" & afeSearch_Order & "||" & paginaActual & "||" & mostrar & "$$totalRegistros"
Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rsAFE, "TBLDATOSAFE_GET_BY_PARAMETERS", sp_parameter)
totalRegistros = sp_ret("totalRegistros")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title>Administración de Autorizaciones para Gastos</title>
<link rel="stylesheet" href="css/MagicSearch.css" type="text/css">
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />

<style type="text/css">
.titleStyle {
	font-weight: bold;
	font-size: 20px;
}


.busquedaTop
{
	border-top: 2px solid #4aa16a;
	border-left: 2px solid #4aa16a;
	border-right: 2px solid #4aa16a;
	-moz-border-radius:5px 5px 0px 0px;
}
.busquedaTop2
{
	border-top: 2px solid #4aa16a;
	border-right: 2px solid #4aa16a;
	-moz-border-radius:0px 5px 0px 0px;
}

.busquedaButton
{
	border-bottom: 2px solid #4aa16a;
	border-left: 2px solid #4aa16a;
	border-right: 2px solid #4aa16a;
	-moz-border-radius:0px 0px 5px 5px;
}

.busquedaVL
{
	border-left: 2px solid #4aa16a;
}

.busquedaVR
{
	border-right: 2px solid #4aa16a;
}

.busquedaHT
{
	border-top: 2px solid #4aa16a;
}

.busquedaHB
{
	border-bottom: 2px solid #4aa16a;
}


.divOculto {
	display: none;
}
</style>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/paginar.js"></script>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/MagicSearchObj.js"></script>
<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>

<script type="text/javascript">
	var puw;
	function lightOn(tr, estado) {
		if (estado == '<% =AFE_ANULADO %>') {
			tr.className = "reg_Header_navdosHL reg_header_rejected";
		} else {
			tr.className = "reg_Header_navdosHL";
		}
	}
	
	function lightOff(tr, estado) {
		if (estado == '<% =AFE_ANULADO %>') {
			tr.className = "reg_Header_navdos reg_header_rejected";
		} else {
			tr.className = "reg_Header_navdos";
		}
	}
	
	function abrirPedido(id) {
		window.open("comprasFichaPedidoCotizacion.asp?idPedido=" + id + "&tab=1", "_blank", "resizable=yes,location=no,scrollbars=yes,menubar=no,statusbar=no,height=500,width=500",false);
	}
	
	function abrirObra(id) {
		window.open("comprasTableroObra.asp?idObra=" + id, "_blank", "resizable=yes,location=no,menubar=no,scrollbars=yes,scrolling=yes",false);		
	}
	function abrirAFEPrint(id){
		window.open("comprasAFEPrint.asp?idAFE=" + id, "_blank", "resizable=yes,location=no,menubar=no,scrollbars=yes,scrolling=yes,height=500,width=700",false);
	}
	function abrirAFEFirma(id) {
		window.open("comprasAFEFirma.asp?idAFE=" + id, "_blank", "resizable=yes,location=no,menubar=no,scrollbars=yes,scrolling=yes",false);		
	}
	function editAFE(idAfe) {
		window.open('comprasAFE.asp?idAFE=' + idAfe);
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
		location.href = "comprasIndex.asp";
	}
		
	function irAdministracion() {
		location.href = "comprasAdministracion.asp";
	}
	
	function irObras() {
		location.href = "comprasObras.asp";
	}
	
	function seleccionarObra(ms) {
		var desc = ms.getSelectedItem();
		if (desc.indexOf('-') != -1) {
			var arr = desc.split('-');
			document.getElementById("cdObra").value = arr[0];
			ms.setValue(arr[1]);
		} else {
			if (desc == "") document.getElementById("cdObra").value = "";							
		}		
	}
	
	function startMagicSearch() {		
		var ms = new MagicSearch("", "divObra", 30, 2, "comprasStreamElementos.asp?tipo=obras");		
		ms.setToken(";");		
		ms.onBlur = seleccionarObra;	
		//ms.setValue(document.getElementById("idObra").value);
		ms.setValue('<%=dsObra%>');
	}	
	function irDirecta(){
		location.href = "comprasAdministrarCotizaciones.asp";
	}
	
	function irPedidos() {
		location.href = "comprasAdministrarPedidos.asp";
	}	
		
	function bodyOnLoad() {	
		
		var tb = new Toolbar('toolbar', 6, 'images/compras/');
		tb.addButton("Home-16x16.png", "Home", "irHome()");				
		tb.addButtonREFRESH("Recargar", "submitInfo()");		
		var swt = tb.addSwitcher("Search-16x16.png", "Buscar", "buscarOn()", "buscarOff()");		
		tb.addButton("Quote_purchase-16x16.png", "Ped. Precio", "irPedidos()");		
		tb.addButton("direct_Purchase-16x16.png", "Compra Directa", "irDirecta()");
		tb.addButton("document-16.png", "Gastos Asociados", "irGastosAsociados()");
		tb.draw();
		<%
		if (searching=1) then%>
			buscarOn();
			tb.changeState(swt);
		<%end if%>
		var pg = new Paginacion("paginacion");			
		pg.paginar(<% =paginaActual %>, <% =totalRegistros %>, <% =mostrar %>, 50, "comprasAdministrarAFEs.asp" + params());		
		pngfix();
	}
	
	function params(){
		var rtrn;
		rtrn = "?cdAFE="+document.getElementById("cdAFE").value;
		rtrn = rtrn+"&afeDivision="+document.getElementById("afeDivision").value;
		rtrn = rtrn+"&afeTitle="+document.getElementById("afeTitle").value;
		rtrn = rtrn+"&afeImport="+document.getElementById("afeImport").value;
		rtrn = rtrn+"&radio_Import="+document.getElementById("valorRadios").value;
		rtrn = rtrn+"&idObra="+document.getElementById("cdObra").value;
		rtrn = rtrn+"&afeOrder="+document.getElementById("afeOrder").value;
		rtrn = rtrn+"&busquedaActiva="+document.getElementById("busquedaActiva").value;
		rtrn = rtrn+"&afepedido="+document.getElementById("afePedido").value;
		return rtrn;
	}
	function setOrder(p_campo,p_orden){
		document.getElementById("afeOrder").value = ' ORDER BY '+p_campo+' '+p_orden;
		document.getElementById("frmSel").submit();
	}
	function anularAFE(idAFE){
		if (confirm("Esta seguro que desea eliminar este AFE?")) {
			window.scrollTo(0,0);
			puw = new winPopUp('popUpAnularAFE', 'comprasAFEAnulacion.asp?idAFE=' + idAFE, 500, 350, 'Anulación del AFE', 'location.reload()');
		}
	}
	
	function agregarArchivoAFE(pId)
	{
		winPopUp('AddAfe', 'comprasAFEUpload.asp?idafe='+pId, 400, 150, '<img src="images/compras/AFE-16X16.png" align="middle"></img> Agregar Archivo AFE', '');
	}
	function irGastosAsociados(){
		var puw = new winPopUp('popupGastosAsociadosAfe','comprasAFEPopUp.asp','900','200','Consumos Asociados de AFE', 'submitInfo()');
	}
	
	
</script>
</head>
<body onLoad="bodyOnLoad()">	
	<div id="toolbar"></div>
	<br>
	<!-- Seccion de Busqueda	-->
	<form name="frmSel" id="frmSel">
	<div id="busqueda" class="divOculto" >
	<table width="90%" cellspacing="0" cellpadding="0" align="center" border="0" >
       <input type="hidden" name="accion" id="accion" value="">
       <tr>
           <td colspan=3 class="busquedaTop" >&nbsp;</td>

           <td width="670">
         <td width="8">
         <td width="1"></td>
      </tr>
       <tr>
           <td width="8" rowspan="3" class="busquedaVL">&nbsp;</td>
           <td width="395" align="center" valign="center"><font class="big" color="#517b4a"><% =GF_TRADUCIR("Búsqueda") %></font></td>
           <td width="11" class="busquedaVR">&nbsp;</td>
         <td></td>
           <td></td>
       </tr>
       <tr>
           <td></td>
           <td width="11" >&nbsp;</td>
           <td colspan="2" class="busquedaTop2">&nbsp;</td>
       </tr>
       <tr>
           <td colspan="3">
                     <table width="100%" align="center" border="0" >
							<tr>
								<input type="hidden" name="afeOrder" id="afeOrder" value="<% =afeSearch_Order %>">
								<td align="right"><% = GF_TRADUCIR("AFE") %>:</td>
								<td><input type="text" name="cdAFE" id="cdAFE" value="<% =afeSearch_cdAFE %>"></td>
								<td align="right"><% = GF_TRADUCIR("Título") %>:</td>
								<td>
									<input type="text" id="afeTitle" name="afeTitle" value="<% =afeSearch_Title %>">								</td>								
							</tr>							
							<tr>
								<td align="right"><% = GF_TRADUCIR("Division")     %>:</td>
								<td>
									<%
									strSQL="Select * from TBLDIVISIONES"
									Call executeQueryDb(DBSITE_SQL_INTRA, rsDivision, "OPEN", strSQL)
									%>
										<select id="afeDivision" name="afeDivision">
											<option value="" <%if (afeSearch_Division = "") then %> selected='true' <%end if%>><% =GF_TRADUCIR("-Seleccione-") %></option>
											<%		
											while (not rsDivision.eof) 	
												if (checkPointAcceso(rsDivision("IDDIVISION"))) then
													 %>
														<option value="<% =rsDivision("IDDIVISION") %>" <% if (afeSearch_Division = rsDivision("IDDIVISION")) then response.write "selected='true'" %>><% =rsDivision("DSDIVISION") %>
											<%		
												end if
												rsDivision.MoveNext()
											wend	
											%>								
										</select>								</td>
								<td align="right"><% = GF_TRADUCIR("Importe")%>:</td>
								<td><input type="text" name="afeImport" id="afeImport" value="<%=afeSearch_import%>" onKeyPress="return controlIngreso(this, event, 'I')">u$s</td>
							</tr>
							<tr>
								<td align="right"><% = GF_TRADUCIR("Obra") %>:</td>
								<td><div id="divObra"></div></td>
								<input type="hidden" id="cdObra" name="cdObra" value="<% =afeSearch_cdObra %>">
								<td>								</td>
								<td align = 'left'>
									<input type="radio" name="radio_Import" id="radio_Import" value="Menor" <%if (afeSearch_radioImport = "Menor") then %>checked="checked"<%end if%> /><% = GF_TRADUCIR("Menor")%>
									<input type="radio" name="radio_Import" id="radio_Import" value="Igual" <%if (afeSearch_radioImport = "Igual") then %>checked="checked"<%end if%> /><% = GF_TRADUCIR("Igual")%>
									<input type="radio" name="radio_Import" id="radio_Import" value="Mayor" <%if (afeSearch_radioImport = "Mayor") then %>checked="checked"<%end if%> /><% = GF_TRADUCIR("Mayor")%>								</td>
							</tr>
							<tr>
								<td align="right"><% = GF_TRADUCIR("Pedido") %>:</td>
								<td>
									<input type="text" id="afePedido" name="afePedido" value="<% =afeSearch_pedido %>">								</td>
							</tr>
                            <tr>
								<td colspan="4" align="center"><input type="button" value="Buscar..." id=submit1 name=submit1 onclick='submitInfo();'></td>						
                            </tr>								
                     </table>         </td>
	           <td height="100%" class="busquedaVR">&nbsp;</td>
      </tr>
	       <tr>
	           <td colspan="5" class="busquedaButton">&nbsp;</td>
           </tr>
	</table>
	</div>
	<input type="hidden" name="busquedaActiva"     id="busquedaActiva"     value=<%=searching%>   >	
	<input type="hidden" name="registrosPorPagina" id="registrosPorPagina" value=<%=mostrar%>     >
	<input type="hidden" name="valorRadios"		   id="valorRadios" 	   value=<%=afeSearch_radioImport%>     >
	</form>
	<br>
	<!-- Seccion de Datos	-->
	<table align="center" width="90%" class="reg_Header">			
			<tr><td colspan="12"><div id="paginacion"></div></td></tr>						
			<tr class="reg_Header_nav">
				<td width="10%" style="text-align: center">            <img src="images\compras\arrow_up_12x12.gif" onclick='setOrder("AFE.cdafe"          ,"asc")' style="cursor:pointer">&nbsp <% =GF_TRADUCIR("AFE") 	      %>&nbsp <img src="images\compras\arrow_down_12x12.gif" onclick='setOrder("AFE.cdafe"          ,"desc")' style="cursor:pointer"></td>
				<td width="15%" style="text-align: center">            <img src="images\compras\arrow_up_12x12.gif" onclick='setOrder("AFE.titulo"         ,"asc")' style="cursor:pointer">&nbsp <% =GF_TRADUCIR("Titulo")        %>&nbsp <img src="images\compras\arrow_down_12x12.gif" onclick='setOrder("AFE.titulo"         ,"desc")' style="cursor:pointer"></td>
				<td width="10%" style="text-align: center">            <img src="images\compras\arrow_up_12x12.gif" onclick='setOrder("AFE.importedolares" ,"asc")' style="cursor:pointer">&nbsp <% =GF_TRADUCIR("Importe")       %>&nbsp <img src="images\compras\arrow_down_12x12.gif" onclick='setOrder("AFE.importedolares" ,"desc")' style="cursor:pointer"></td>				
				<td width="10%" style="text-align: center" colspan="2"><img src="images\compras\arrow_up_12x12.gif" onclick='setOrder("OBRA.CDOBRA"        ,"asc")' style="cursor:pointer">&nbsp <% =GF_TRADUCIR("Ptda. Presup.") %>&nbsp <img src="images\compras\arrow_down_12x12.gif" onclick='setOrder("OBRA.CDOBRA"        ,"desc")' style="cursor:pointer"></td>
				<td width="10%" style="text-align: center" colspan="2"><img src="images\compras\arrow_up_12x12.gif" onclick='setOrder("PEDIDO.CDPEDIDO"    ,"asc")' style="cursor:pointer">&nbsp <% =GF_TRADUCIR("Pedido")        %>&nbsp <img src="images\compras\arrow_down_12x12.gif" onclick='setOrder("PEDIDO.CDPEDIDO"    ,"desc")' style="cursor:pointer"></td>
				<td width="1%"  style="text-align: center">.</td>
				<td width="1%"  style="text-align: center">.</td>
				<td width="1%"  style="text-align: center">.</td>
				<td width="1%"  style="text-align: center">.</td>
			</tr>		
<%	if (not rsAFE.eof) then			
		while (not rsAFE.eof) %>
		<tr class="reg_Header_navdos <%If (rsAFE("Confirmado") = AFE_ANULADO) then Response.write "reg_header_rejected" %>" onMouseOver="javascript:lightOn(this, '<% =rsAFE("CONFIRMADO") %>')" onMouseOut="javascript:lightOff(this, '<% =rsAFE("CONFIRMADO") %>')">			
			<td style="text-align: center" onClick="abrirAFEPrint(<% =rsAFE("IDAFE") %>)"><% =rsAFE("CDAFE") %></td>
			<td  onclick="abrirAFEPrint(<% =rsAFE("IDAFE") %>)">
				<% if ( len(rsAFE("TITULO")) )>30 then
					response.write Left(rsAFE("TITULO"),30) & "..."
				else
					response.write rsAFE("TITULO")
				end if %>			</td>
			<td style="text-align: right" onClick="abrirAFEPrint(<% =rsAFE("IDAFE") %>)"><% =getSimboloMoneda(MONEDA_DOLAR) %>&nbsp;&nbsp;<% =GF_EDIT_DECIMALS(rsAFE("IMPORTEDOLARES"),2) %></td>
			
			<td style="text-align: center" onClick="abrirAFEPrint(<% =rsAFE("IDAFE") %>)">
			<%	if (rsAFE("IDOBRA") > 0) then	%>
				<% =rsAFE("OBRA_CDOBRA") & "-" & rsAFE("IDAREA") & "-" & rsAFE("IDDETALLE") %>
			<%	end if	%>	
			</td>				
			
			<td style="text-align: center" width="1%">
				<%	if (rsAFE("IDOBRA") > 0) then	%>
				<span style="cursor:pointer" onClick="abrirObra(<% =rsAFE("IDOBRA") %>)"><img src="images/compras/obr-16x16.png" title="<% =GF_TRADUCIR("Ver Partida") %>" /></span>
				<%	end if	%>			</td>
			<td style="text-align: center" onClick="abrirAFEPrint(<% =rsAFE("IDAFE") %>)"><% =rsAFE("PEDIDO_CDPEDIDO") %></td>
			<td style="text-align: center" width="1%">
				<%	if (rsAFE("IDPEDIDO") > 0) then	%>
				<span style="cursor:pointer" onClick="abrirPedido(<% =rsAFE("IDPEDIDO") %>)"><img src="images/compras/pct-16x16.png" title="Ver Pedido de Cotizacion" /></span>
				<%	end if	%>			</td>
			<%
			strSQL = "Select count(*) cant from TBLDATOSAFE where filescan is not null and idAfe= " & rsAFE("IDAFE")
			Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
			cantFiles = 0
			if (not rs.EoF) then cantFiles = rs("cant")
			if (cantFiles = 0) then
				imgAFEUpload = "AFE_upload-16.png"
			else
				imgAFEUpload = "AFE_confirm-16x16.png"
			end if
			%>
			<td style="text-align: center"><img src="images/compras/<% =imgAFEUpload %>" title="Agregar Archivo" onClick="agregarArchivoAFE('<% =rsAFE("IDAFE") %>')"></td>						
			<td style="text-align: center"><% =getEditAFEIcon(rsAFE("IDAFE")) %></td>
			<td style="text-align: center">
				<%
					textoAux = ""
					if ((IsNumeric(rsAFE("CONFIRMADO"))) or (rsAFE("CONFIRMADO") = AFE_NO_CONFIRMADO)) then
						'JAS - cdUsuario = getUsuarioAFirmar(rsAFE("IDAFE"))
						dsUsuario = getDSUsuarioAFirmar(rsAFE("IDAFE"))
						'if (dsUsuario = "") then dsUsuario = cdUsuario
						textoAux = "Se esta esperando la firma del usuario " & dsUsuario
					else 
					    if (rsAFE("CONFIRMADO") = AFE_ESPERA_HAMBURGO) then
					        textoAux = "Se espera la aprobacion de Hamburgo."
					    end if
					end if
					if (textoAux <> "") then
				%>		<span style="cursor:pointer" onClick="alert('<% =textoAux %>')"><img style="cursor:pointer" src="images/compras/action_warning-16x16.png" title="<% =textoAux %>"></span>				
				<%	else	%>
						<span style="cursor:pointer" onClick="abrirAFEPrint(<% =rsAFE("IDAFE") %>)"><img style="cursor:pointer" src="images/compras/printer-16x16.png" title="<% =GF_TRADUCIR("Imprimir AFE") %>"></span>
				<%	end if	%>			
			</td>
			<td style="text-align: center"><% =getRejectAFEIcon(rsAFE("IDAFE")) %></td>
		</tr>
<%		rsAFE.MoveNext()
		wend
	else	%>
		<tr class="TDNOHAY"><td colSpan="12"><% =GF_TRADUCIR("No hay informacion disponible en estos momentos") %></td></tr>		
<%	end if	%>
		</table>
</body>
</html>
