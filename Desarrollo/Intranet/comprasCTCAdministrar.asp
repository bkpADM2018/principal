<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientossql.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->
<!--#include file="Includes/procedimientosTitulos.asp"-->

<%
Call comprasControlAccesoCM(RES_CC)

'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'ADMINISTRACVION DE CONTRATOS, SE LISTAN LOS CTC Y SE DESDE AQUI PUEDEN --
'INGRESAR AL TABLERO DE CONTROL, CONFIRMARLOS Y VER EL ARCHIVO ADJUNTO ---
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'***********************************************
'*************  COMIENZO DE PAGINA  ************
'***********************************************
Dim myOrder, mostrar, paginaActual, myWhere, reg, tituloCTC,totalRegistros

Call GP_ConfigurarMomentos()

mostrar = GF_PARAMETROS7("registrosPorPagina",0,6)
paginaActual = GF_PARAMETROS7("numeroPagina",0,6)
if (mostrar = 0) then mostrar = 10
if (paginaActual = 0) then paginaActual = 1

myOrder = GF_PARAMETROS7("myOrder","",6)
CTC_cdContrato = GF_PARAMETROS7("cdContrato","",6)
CTC_idDivision = GF_PARAMETROS7("idDivision",0,6)
CTC_cdPedido = GF_PARAMETROS7("cdPedido","",6)
CTC_cdObra = GF_PARAMETROS7("cdObra","",6)
CTC_dsObra = GF_PARAMETROS7("dsObra","",6)
CTC_idProveedor = GF_PARAMETROS7("idProveedor",0,6)
CTC_dsProveedor = GF_PARAMETROS7("dsProveedor","",6)
CTC_cdResponsable = GF_PARAMETROS7("cdResponsable","",6)
CTC_dsResponsable = getUserDescription(CTC_cdResponsable)
CTC_estado = GF_PARAMETROS7("estado",0,6)
CTC_Titulo = UCase(GF_PARAMETROS7("titulo","",6))
CTC_Tipo = GF_PARAMETROS7("tipo","",6)

sp_parameter = CTC_cdContrato &"||"& CTC_cdPedido &"||"& CTC_cdObra &"||"& CTC_idProveedor &"||"& CTC_cdResponsable &"||"&_
			   CTC_estado &"||"& CTC_idDivision &"||"& myOrder &"||"& CTC_Titulo &"||"& CTC_Tipo &"||"& paginaActual &"||"& mostrar &"$$totalRegistros"
Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rsCTC, "TBLOBRACONTRATOS_GET_BY_PARAMETERS", sp_parameter)
totalRegistros = sp_ret("totalRegistros")

%>
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
		<title>Administraci�n de Contratos</title>
		<link rel="stylesheet" href="css/MagicSearch.css" type="text/css">
		<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
		<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
		<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
    <link rel="stylesheet" href="css/main.css" type="text/css">
	<script type="text/javascript" src="scripts/Toolbar.js"></script>
	<script type="text/javascript" src="scripts/paginar.js"></script>
	<script type="text/javascript" src="scripts/channel.js"></script>
	<script type="text/javascript" src="scripts/MagicSearchObj.js"></script>
	<script type="text/javascript" src="scripts/controles.js"></script>	
	<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
	<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
	<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
	<script type="text/javascript">
			
		var ch = new channel();
			
		function modificarContrato(id) {
			window.open('comprasCTCNuevo.asp?CTC_idContrato=' + id, "_blank", "resizable=yes,location=no,menubar=no,statusbar=no,height=550,width=650",false);
		}
			
		function irNuevo() {
			window.open('comprasCTCNuevo.asp', "_blank", "resizable=yes,location=no,menubar=no,statusbar=no,height=640,width=650",false);
		}
		function bodyOnLoad(){
			var tb = new Toolbar('toolbar', 6, 'images/compras/');
			tb.addButton("Home-16x16.png", "Home", "irHome()");
			tb.addButtonREFRESH("Recargar", "submitInfo()");				
			tb.addButton("CTC_New-16x16.png", "Nuevo", "irNuevo()");							
			tb.addButton("Quote_purchase-16x16.png", "Ped. Precio", "irPedidos()");
			tb.addButton("Obra_16x16.png", "Obras", "irObras()");				
			tb.addButton("excel-16.png", "Ctos. c/ Saldo", "irSaldosPendientesPrintXLS()");
			tb.draw();
			var pg = new Paginacion("paginacion");			
			pg.paginar("<% =paginaActual %>", "<% =totalRegistros %>", "<% =mostrar %>", 50, "comprasCTCAdministrar.asp" + params());				
			startMagicSearch()
		}
		function irSaldosPendientesPrintXLS(){
			window.open("comprasctcsaldospendientesprintxls.asp","_blank",false);
		}
		function irHome(){
			location.href = "comprasIndex.asp";
		}
		function submitInfo(){
			document.getElementById("frmSearch").submit();
		}		
		function irObras(){
			location.href = "comprasObras.asp";
		}
		function irPedidos(){
			location.href = "comprasAdministrarPedidos.asp";
		}
		function abrirPedido(id) {
			window.open("comprasFichaPedidoCotizacion.asp?idPedido=" + id + "&tab=1", "_blank", "resizable=yes,location=no,scrollbars=yes,menubar=no,statusbar=no,height=500,width=500",false);
		}
		function abrirObra(id) {
			window.open("comprasTableroObra.asp?idObra=" + id, "_blank", "resizable=yes,location=no,menubar=no,scrollbars=yes,scrolling=yes",false);
		}
		function abrirCTC(id) {
			window.open("comprasCTC.asp?idContrato=" + id, "_blank", "resizable=yes,location=no,menubar=no,scrollbars=yes,scrolling=yes",false);
		}
		function abrirAdjunto(id) {
			window.open("comprasOpenArchivo.asp?idContrato=" + id, "_blank", "resizable=yes,location=no,menubar=no,scrollbars=yes,scrolling=yes,height=200,width=300",false);
		}
		function confirmarContrato(id) {
			var puw = new winPopUp('popUpCTCConfirm', 'comprasCTCConfirmar.asp?idContrato='+id, 600,420, 'Confirmar Contrato', 'submitInfo()');
		}
		function startMagicSearch() {
			var ms = new MagicSearch("", "divObra", 30, 2, "comprasStreamElementos.asp?tipo=obras");
			ms.setToken(";");
			ms.onBlur = seleccionarObra;
			ms.setValue('<% =CTC_dsObra %>');
			var msProv = new MagicSearch("", "companyName0", 30, 2, "comprasStreamElementos.asp?tipo=empresas");
			msProv.setMinChar(3);
			msProv.setToken(";");
			msProv.onBlur = SeleccionarProveedor;
			msProv.setValue('<% =CTC_dsProveedor %>');
			var msResp = new MagicSearch("", "divResp", 30, 2, "comprasStreamElementos.asp?tipo=personas");
			msResp.setToken(";");
			msResp.onBlur = seleccionarResponsable;
			msResp.setValue('<% =CTC_dsResponsable %>');
		}
		function seleccionarObra(ms) {
			var desc = ms.getSelectedItem();
			if (desc.indexOf('-') != -1) {
				var arr = desc.split('-');
				document.getElementById("cdObra").value = arr[0];
				document.getElementById("dsObra").value = arr[1];
				ms.setValue(arr[1]);
			} else {
				if (desc == ""){
					document.getElementById("cdObra").value = "";
					document.getElementById("dsObra").value = "";
				}
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
		function seleccionarResponsable(ms) {
			var desc = ms.getSelectedItem();
			if (desc.indexOf('-') != -1) {
				var arr = desc.split('-');
				document.getElementById("cdResponsable").value = arr[0];
				ms.setValue(arr[1]);
			} else {
				if (desc == "") document.getElementById("cdResponsable").value = "";
			}
		}
		function setOrder(p_campo,p_orden){
			document.getElementById("myOrder").value = p_campo+' '+p_orden;
			submitInfo();
		}
		function params(){
			var rtrn;
			rtrn = "?cdContrato="+document.getElementById("cdContrato").value;
			rtrn = rtrn+"&idDivision="+document.getElementById("idDivision").value;
			rtrn = rtrn+"&cdObra="+document.getElementById("cdObra").value;
			rtrn = rtrn+"&myOrder="+document.getElementById("myOrder").value;			
			rtrn = rtrn+"&cdPedido="+document.getElementById("cdPedido").value;
			rtrn = rtrn+"&idProveedor="+document.getElementById("idProveedor").value;
			rtrn = rtrn+"&cdResponsable="+document.getElementById("cdResponsable").value;
			rtrn = rtrn+"&estado="+document.getElementById("estado").value;
			rtrn = rtrn+"&tipo="+document.getElementById("tipo").value;
			return rtrn;
		}
		function lightOn(tr, estado) {
			if (estado == '<% =ESTADO_CTC_CANCELADO %>') {
				tr.className = "reg_Header_navdosHL reg_header_rejected";
			} else {
				tr.className = "reg_Header_navdosHL";
			}
		}
		function lightOff(tr, estado) {
			if (estado == '<% =ESTADO_CTC_CANCELADO %>') {
				tr.className = "reg_Header_navdos reg_header_rejected";
			} else {
				tr.className = "reg_Header_navdos";
			}
		}
		function anularCTCCallback(pId){	
			var resp = ch.response();
			var myImg = document.getElementById(pId);		
			if (resp != "<% =RESPUESTA_OK %>") {
				myImg.src="images/compras/CTZ_cancel-16x16.png";			
				alert(resp);			
			} else {
				location.reload();
			}
		}
	
		function anularContrato(idContrato, img){
			if (confirm("Esta seguro que desea anular este Contrato?")) {
				img.src = "images/loading_small_green.gif"
				ch.bind("comprasCTCAnularContratoAjax.asp?idContrato=" + idContrato, "anularCTCCallback('" + img.id + "')");
				ch.send();			
			}		
		}
	</script>
    <style type="text/css">
        td{font-weight:bold;}
        .reg_header_navdosHL {
            font-style:italic;
        }
    </style>
</head>
<body onLoad="bodyOnLoad()">
	<div id="toolbar"></div>

	<!-- Seccion de Busqueda	-->
	<form id="frmSearch" name="frmSearch">
        <div id="divSearch" style="border-bottom:none;" class="tableaside size100">
            <h3><% =GF_TRADUCIR("B�SQUEDA") %></h3>
            <div id="searchfilter" class="tableasidecontent">
                <div class="col16 reg_header_navdos"> <% =GF_TRADUCIR("Cod. Cto.:")%></div>
                <div class="col16">
                    <input type="text" name="cdContrato" id="cdContrato" value="<% =CTC_cdContrato %>">
                </div>
                <div class="col16 reg_header_navdos"> <%=GF_TRADUCIR("Titulo:")%> </div>
                <div class="col16">
                    <input type="text" id="titulo" name="titulo" value="<% =CTC_Titulo %>">
                </div>

                <div class="col16 reg_header_navdos"> <% =GF_TRADUCIR("Obra:")%> </div>
                <div class="col16">
                    <div id="divObra"></div>
                    <input type="hidden" id="cdObra" name="cdObra" value="<% =CTC_cdObra %>">
                    <input type="hidden" id="dsObra" name="dsObra" value="<% =CTC_dsObra %>">
                </div>

                <div class="col16 reg_header_navdos"> <% =GF_TRADUCIR("Division:")%> </div>
                <div class="col16">
                    <%
                        strSQL="Select * from TBLDIVISIONES"
                        Call executeQueryDb(DBSITE_SQL_INTRA, rsDivision, "OPEN", strSQL)
                    %>
					<select id="idDivision" name="idDivision">
						<option value="" <%if (CTC_idDivision = "") then %> selected='true' <%end if%>><% =GF_TRADUCIR("Todas") %></option>
					<%		
						while (not rsDivision.eof) 	
							if (checkPointAcceso(rsDivision("IDDIVISION"))) then
								if not isAuditor(rsDivision("IDDIVISION")) then %>
									<option value="<% =rsDivision("IDDIVISION") %>" <% if (CTC_idDivision = rsDivision("IDDIVISION")) then response.write "selected='true'" %>><% =rsDivision("DSDIVISION") %>
					<%			end if
							end if
							rsDivision.MoveNext()
						wend	
					%>								
					</select>
                </div>
                <div class="col16 reg_header_navdos"><% =GF_TRADUCIR("Proveedor:")%>  </div>
                <div class="col16">
                    <div id="companyName0"></div>
                    <input type="hidden" id="idProveedor" name="idProveedor" value="<% =CTC_idProveedor %>">
					<input type="hidden" id="dsProveedor" name="dsProveedor" value="<% =CTC_dsProveedor %>">
                </div>
                <div class="col16 reg_header_navdos"><% =GF_TRADUCIR("Pedido:")%>  </div>
                <div class="col16">
                    <input type="text" id="cdPedido" name="cdPedido" value="<% =CTC_cdPedido %>">
                </div>
                <div class="col16 reg_header_navdos"><% =GF_TRADUCIR("Estado:")%> </div>
                <div class="col16">
                    <select id="estado" name="estado">
						<option value="0"							 <%if (CTC_estado = 0) then						%> selected='true' <%end if%>><% =GF_TRADUCIR("Activos") %></option>
						<option value="<% =ESTADO_CTC_PENDIENTE  %>" <%if (cint(CTC_estado) = ESTADO_CTC_PENDIENTE) then  %> selected='true' <%end if%>><% =GF_TRADUCIR("Pendientes") %></option>
						<option value="<% =ESTADO_CTC_AUTORIZADO %>" <%if (cint(CTC_estado) = ESTADO_CTC_AUTORIZADO) then %> selected='true' <%end if%>><% =GF_TRADUCIR("Autorizados") %></option>
						<option value="<% =ESTADO_CTC_FINALIZADO %>" <%if (cint(CTC_estado) = ESTADO_CTC_FINALIZADO) then %> selected='true' <%end if%>><% =GF_TRADUCIR("Finalizados") %></option>
						<option value="<% =ESTADO_CTC_CANCELADO  %>" <%if (cint(CTC_estado) = ESTADO_CTC_CANCELADO) then  %> selected='true' <%end if%>><% =GF_TRADUCIR("Cancelados") %></option>
					</select>
                </div>
                <div class="col16 reg_header_navdos"> <% =GF_TRADUCIR("Responsable:")%> </div>
                <div class="col16">
                    <div id="divResp"></div>
                    <input type="hidden" id="cdResponsable" name="cdResponsable" value="<% =CTC_cdResponsable %>">
                </div>
                <!--Nuevo filtro-->
                <div class="col16 reg_header_navdos"><% =GF_TRADUCIR("Tipo:")%></div>
                <div class="col16">
                    <select id="tipo" name="tipo">
                        <option value="" <%if(CTC_Tipo = "") then %> selected='true' <%end if%>><% =GF_TRADUCIR("Todos") %></option>
                        <option value="CONTRATOS" <%if(CTC_Tipo = "CONTRATOS") then %> selected='true' <%end if%>><% =GF_TRADUCIR("Contratos") %></option>
                        <option value="<%=CONTRATO_TIPO_SERVICIO%>" <%if(CTC_Tipo = CONTRATO_TIPO_SERVICIO) then %> selected='true' <%end if%>><% =GF_TRADUCIR("Servicios") %></option>
                    </select>
                </div>
                <!---->
                <span class="btnaction" style="margin-bottom:15px; margin-top:10px;">
                    <input type="button" value="Buscar" id=submit1 name=submit1 onclick='submitInfo();'>
                </span>
            </div>
        </div>
		<input type="hidden" name="myOrder" id="myOrder" value="">
		<input type="hidden" name="registrosPorPagina" id="registrosPorPagina" value=10>
		<input type="hidden" name="numeroPagina" id="numeroPagina" value=1>
	</form>
    <br/><br>
    <table class="datagrid" width="90%" align="center">
        <thead>
            <tr>
                <th align="center" width="30%" colspan="2">
                    <img src="images\arrow_plus_up.gif" onclick='setOrder("CTC.CDCONTRATO", "asc")' style="cursor:pointer">
                    &nbsp;<% =GF_TRADUCIR("Contrato") %>&nbsp;
                    <img src="images\arrow_plus_down.png" onclick='setOrder("CTC.CDCONTRATO", "desc")' style="cursor:pointer">
                </th>
                <th align="center" width="20%" colspan="2">
                    <img src="images\arrow_plus_up.gif" onclick='setOrder("OBR.CDOBRA", "asc")' style="cursor:pointer">
                    &nbsp;<% =GF_TRADUCIR("Partida Activa") %>&nbsp;
                    <img src="images\arrow_plus_down.png" onclick='setOrder("OBR.CDOBRA", "desc")' style="cursor:pointer">
                </th>
                <th align="center" width="10%" colspan="2">
                    <img src="images\arrow_plus_up.gif" onclick='setOrder("PCT.CDPEDIDO", "asc")' style="cursor:pointer">
                    &nbsp;<% =GF_TRADUCIR("Pedido") %>&nbsp;
                    <img src="images\arrow_plus_down.png" onclick='setOrder("PCT.CDPEDIDO", "desc")' style="cursor:pointer">
                </th>
                <th align="center" width="20%">
                    <img src="images\arrow_plus_up.gif" onclick='setOrder("CTC.IDPROVEEDOR", "asc")' style="cursor:pointer">
                    &nbsp;<% =GF_TRADUCIR("Proveedor") %>&nbsp;
                    <img src="images\arrow_plus_down.png" onclick='setOrder("CTC.IDPROVEEDOR", "desc")' style="cursor:pointer">
                </th>
                <th align="center" width="11%">
                    <img src="images\arrow_plus_up.gif" onclick='setOrder("CTC.FECHAVTO", "asc")' style="cursor:pointer">
                    &nbsp;<% =GF_TRADUCIR("Fecha Vto") %>&nbsp;
                    <img src="images\arrow_plus_down.png" onclick='setOrder("CTC.FECHAVTO", "desc")' style="cursor:pointer">
                </th>
                <td align="center" width="32px">.</td>
                <td align="center" width="32px">.</td>
                <td align="center" width="32px" style="border-top-right-radius:8px;">.</td>
            </tr>
        </thead>
<%			if (not rsCTC.eof) then
			While (not rsCTC.eof)
%>
        <tr class="<% if (CInt(rsCTC("ESTADO")) = ESTADO_CTC_CANCELADO) then Response.Write "reg_header_rejected" %>" style="cursor:pointer">
            <td align="center" width="10%" onclick='abrirCTC(<% =rsCTC("IDCONTRATO") %>)'>
				<% =rsCTC("CDCONTRATO") %>
			</td>
            <td align="left" onclick='abrirCTC(<% =rsCTC("IDCONTRATO") %>)'>
				<%tituloCTC = rsCTC("TITULO")
				if (Len(tituloCTC) > 30) then tituloCTC = Left(tituloCTC,30) & "..."
				response.write tituloCTC %>
			</td>
            <td align="center" onclick='abrirCTC(<% =rsCTC("IDCONTRATO") %>)'>
				<%if (rsCTC("IDOBRA") > 0) then Response.write rsCTC("CDOBRA")%>
			</td>
            <td align="center" width="2%">
				<%if ((CLng(rsCTC("IDOBRA")) > 0) and (CLng(rsCTC("IDOBRA")) <> OBRA_GEID)) then %>
					<img src="images\compras\Obra_16x16.png" onclick='abrirObra(<% =rsCTC("IDOBRA") %>)'>
				<%end if %>	
			</td>
            <td align="center" onclick='abrirCTC(<% =rsCTC("IDCONTRATO") %>)'>
				<%if (rsCTC("IDPEDIDO") > 0) then 										
					Response.Write rsCTC("CDPEDIDO") 
				end if%>
			</td>
            <td align="center" width="2%">
                <%if (rsCTC("IDPEDIDO") > 0) then%>
                    <img src="images\compras\PCT-16X16.png" onclick='abrirPedido(<% =rsCTC("IDPEDIDO") %>)'>
                <%end if%>
            </td>
            <td align="left" onclick='abrirCTC(<% =rsCTC("IDCONTRATO") %>)'>
                <% =rsCTC("IDPROVEEDOR") & "-" & getDescripcionProveedor(rsCTC("IDPROVEEDOR")) %>
            </td>
            <td onclick='abrirCTC(<% =rsCTC("IDCONTRATO") %>)' align="center">
				<%if(not isNull(rsCTC("FECHAVTO")))then 
						Response.Write GF_FN2DTE(rsCTC("FECHAVTO"))
				else
			    	Response.Write "Sin fecha definida"
				end if%>
			</td>
            <td align="center">
				<% 'consulto si es de legales (para que confirme o para que suba archivos)
				if (canConfirmCTC(session("Usuario"), rsCTC("IDCONTRATO"))) then %>
					<img src="images\round_up_arrow16.png" onClick="confirmarContrato(<% =rsCTC("IDCONTRATO") %>);" title="<% =GF_TRADUCIR("Confirmar") %>">
				<%end if %>
			</td>
            <td align="center">
                <% if (Trim(rsCTC("ARCHIVO_EXT")) <> "") then
					'tiene archivos adjuntos %>
					<img src="images\compras\CTC-16X16.png" onclick='abrirAdjunto(<% =rsCTC("IDCONTRATO") %>)' title="<% =GF_TRADUCIR("Ver Adjunto") %>">
				<%end if %>
            </td>
            <td align="center" width="32px" style="cursor:pointer">
                <%if ((canConfirmCTC(session("Usuario"), rsCTC("IDCONTRATO"))) and ((rsCTC("ESTADO") = ESTADO_CTC_PENDIENTE) or (rsCTC("ESTADO") = ESTADO_CTC_AUTORIZADO))) then %>
                    <img title="<%=GF_TRADUCIR("Anular Contrato")%>" id="ID_<%=rsCTC("IDCONTRATO")%>" src="images/compras/CTZ_cancel-16x16.png" onclick="anularContrato(<% =rsCTC("IDCONTRATO") %>, this)">
                    <%end if %>
                </td>
            </tr>
<%		rsCTC.MoveNext
	Wend 
	else%>
        <tr class="reg_header_nav">
			<td colspan="12" align="center"><% =GF_TRADUCIR("No hay informacion disponible en estos momentos") %></td>
		</tr>
    <%end if%>      
    <tfoot>
        <tr>
            <td colspan="12">
                <div id="paginacion"></div>
            </td>
        </tr>
    </tfoot>
</table>       
</body>
</html>