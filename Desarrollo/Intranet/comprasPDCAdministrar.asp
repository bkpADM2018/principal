<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientos.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosTitulos.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<%
Const BUSQUEDA_CDPEDIDO 	= 1
Const BUSQUEDA_NROPDC		= 2
Const BUSQUEDA_ASEGURADORA	= 3
Const BUSQUEDA_TOMADOR		= 4
Const BUSQUEDA_VENCIMIENTO	= 5
'----------------------------------------------------------------------------------

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
'-----------------------------------------------------------------------------------
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
'-----------------------------------------------------------------------------------
'Solo actualiza aquellas polizas que signe vigente, no afectara a las que se devolvieron, a las mismas vencidas ,
' y las anuladas (a estas tres si no se aplicara ese filtro todas estarian vencidas )y las pendientes 
'(en ese momento no tiene vencimiento)
Function buscarPDCvencidos()
	Dim strSQL 	
	strSQL = "			 SELECT IDPDC					    "
	strSQL = strSQL & "	 FROM TBLPOLIZASCAUCION	"
	strSQL = strSQL & "	 WHERE VENCIMIENTO < " & Left(session("MmtoSistema"),8) 
	strSQL = strSQL & "		AND ESTADO = " & ESTADO_PDC_RECIBIDA 
	strSQL = strSQL & "	 ORDER BY IDPDC "
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if(not rs.Eof)then
		while not rs.Eof
			auxPDC = auxPDC & rs("IDPDC") & ","
			rs.MoveNext()
		wend
		auxPDC = Left(auxPDC,Len(auxPDC)-1)
		strSQL = "UPDATE TBLPOLIZASCAUCION SET ESTADO = " & ESTADO_PDC_VENCIDA & " WHERE IDPDC IN ("& auxPDC & ")"		
		Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
	end if
End Function
'--------------------------------------------------------------------------------------
Function getSqlOrder(ByRef strOrder, pCampoOrder, pTipoOrder)
	if (strOrder	= "") then
		strOrder = " ORDER BY "
	else
		strOrder = myOrder & ", "
	end if
	Select case (pCampoOrder)
		case BUSQUEDA_NROPDC:
			strOrder = strOrder & " POL.IDPDC " & pTipoOrder
		case BUSQUEDA_CDPEDIDO:			
			strOrder = strOrder & " SUBSTRING(PCT.CDPEDIDO, 0, Len(PCT.CDPEDIDO) - 6 ) " & pTipoOrder
			strOrder = strOrder & " , SUBSTRING(pct.cdpedido, 8, 2 ) " & pTipoOrder
			strOrder = strOrder & " , SUBSTRING(pct.cdpedido, (LEN(pct.cdpedido) - 5), 3) " & pTipoOrder
		case BUSQUEDA_ASEGURADORA:
			strOrder = strOrder & " SEC.DSASEGURADORA " & pTipoOrder
		case BUSQUEDA_TOMADOR:
			strOrder = strOrder & " EMP.NOMEMP " & pTipoOrder
		case BUSQUEDA_VENCIMIENTO:
			strOrder = strOrder & " POL.VENCIMIENTO " & pTipoOrder		
		case else:
			strOrder = strOrder & " POL.ESTADO ASC "
	end Select	
End Function
'*********************************************************************************************************************'
'**************************************************		INICIO DE PAGINA    ******************************************'
'*********************************************************************************************************************'
Dim gv_IdPDC, gv_cdPedido, gv_NroPoliza, gv_idAseguradora, gv_Tomador, gv_Monto, gv_Vencimiento, idEstado, gv_DsAseguradora,setOrder
Dim gv_idDivision, gv_Fecha, fromAdmPed, flagAdmin, flagUser, rs, paginaActualmostrar,lineasTotales, accion, hayBusqueda,params

Call comprasControlAccesoCM(RES_PDC)
Call buscarPDCvencidos()
gv_IdPDC = GF_PARAMETROS7("IdPDC", 0, 6)
call addParam("IdPDC", gv_IdPDC, params)
gv_cdPedido = UCase(GF_PARAMETROS7("cdPedido","",6))
call addParam("cdPedido", gv_cdPedido, params)
gv_NroPoliza = GF_PARAMETROS7("NroPoliza", "", 6)
call addParam("NroPoliza", gv_NroPoliza, params)
gv_DsAseguradora = Trim(Ucase(GF_PARAMETROS7("dsAseguradora", "", 6)))
call addParam("dsAseguradora", gv_DsAseguradora, params)
if(gv_DsAseguradora <> "")then 
	gv_idAseguradora = GF_PARAMETROS7("idAseguradora", 0, 6)
	call addParam("idAseguradora", gv_idAseguradora, params)
end if	
gv_Monto = GF_PARAMETROS7("Monto", 0, 6)
call addParam("Monto", gv_Monto, params)
gv_Vencimiento = GF_PARAMETROS7("Vencimiento", 0, 6)
call addParam("Vencimiento", gv_Vencimiento, params)
idEstado = GF_PARAMETROS7("idEstado",0,6)
call addParam("idEstado", idEstado, params)
gv_idDivision = GF_PARAMETROS7("idDivision",0,6)
call addParam("idDivision", gv_idDivision, params)
gv_Tomador = GF_PARAMETROS7("idTomador", 0, 6)
call addParam("idTomador", gv_Tomador, params)
gv_dsTomador = GF_PARAMETROS7("dsTomador", "", 6)
call addParam("dsTomador", gv_dsTomador, params)
gv_Importe = GF_PARAMETROS7("importeTot", 2, 6)
call addParam("importeTot", gv_Importe, params)
gv_Moneda  = GF_PARAMETROS7("tipoMoneda" ,"",6)
call addParam("tipoMoneda", gv_Moneda, params)
gv_ImporteAprox  = GF_PARAMETROS7("radio_Import" ,"",6)
call addParam("radio_Import", gv_ImporteAprox, params)
gv_AVenc = GF_PARAMETROS7("txtAnioEmision","",6)
call addParam("txtAnioEmision", gv_AVenc, params)
gv_MVenc = GF_PARAMETROS7("txtMesEmision","",6)
call addParam("txtMesEmision", gv_MVenc, params)
gv_DVenc = GF_PARAMETROS7("txtDiaEmision","",6)
call addParam("txtDiaEmision", gv_DVenc, params)
gv_Fecha = Trim(GF_nDigits(gv_AVenc,4)&GF_nDigits(gv_MVenc,2)&GF_nDigits(gv_DVenc,2))
accion = GF_PARAMETROS7("accion", "", 6)
mostrar = GF_PARAMETROS7("registrosPorPagina",0,6)
paginaActual = GF_PARAMETROS7("numeroPagina",0,6)
if (mostrar = 0) then mostrar = 10
if (paginaActual = 0) then paginaActual = 1
hayBusqueda = GF_PARAMETROS7("busquedaActiva",0,6)
call addParam("busquedaActiva", hayBusqueda, params)

setOrderTipo = UCASE(GF_PARAMETROS7("setOrderTipo", "", 6))
call addParam("setOrderTipo", setOrderTipo, params)
setOrder = GF_PARAMETROS7("setOrder", 0, 6)
call addParam("setOrder", setOrder, params)
Call getSQLOrder(strOrder, setOrder, setOrderTipo)


Set rsPDC = readPDC(gv_IdPDC, gv_cdPedido, gv_NroPoliza, gv_idAseguradora, gv_Tomador, idEstado, gv_idDivision, strOrder, gv_Importe, gv_Moneda, gv_ImporteAprox, gv_Fecha, false)
Call setupPaginacion(rsPDC, paginaActual, mostrar)
lineasTotales = rsPDC.recordcount

%>
<HTML>
<HEAD>
<META http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<TITLE><% =GF_TRADUCIR("Sistema de Compras - Administrar PDC") %></TITLE>
<LINK rel="stylesheet" href="css/MagicSearch.css" type="text/css">
<LINK href="css/ActisaIntra-1.css" rel="stylesheet" type="text/css" />
<LINK rel="stylesheet" href="css/JQueryUpload2.css"	 type="text/css">
<LINK rel="stylesheet" href="css/tabs.css" TYPE="text/css" MEDIA="screen">
<LINK rel="stylesheet" href="css/tabs-print.css" TYPE="text/css" MEDIA="print">
<LINK rel="stylesheet" href="css/Toolbar.css" type="text/css">
<LINK rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css"	 type="text/css">
<LINK href="css/jqueryUI/custom-theme/jquery-ui-1.8.15.custom.css" rel="stylesheet" type="text/css" />
<STYLE type="text/css">
.titleStyle {
	font-weight: bold;
	font-size: 20px;
}

.divOculto {
	display: none;
}
.titleVencido
{
    BORDER-BOTTOM: #f80800 1px solid;
    BORDER-LEFT: #f40800 1px solid;
    BACKGROUND-COLOR: #F5340E;	
    FONT-FAMILY: verdana,arial,san-serif;
    HEIGHT: 19px;
    COLOR: #ffffff;
    FONT-SIZE: 10px;
    BORDER-TOP: #f40800 1px solid;
    FONT-WEIGHT: bold;
    BORDER-RIGHT: #f40800 1px solid;
    TEXT-DECORATION: none
}

</STYLE>
<SCRIPT type="text/javascript" src="scripts/channel.js"></SCRIPT>
<SCRIPT type="text/javascript" src="scripts/formato.js"></SCRIPT>
<SCRIPT type="text/javascript" src="scripts/controles.js"></SCRIPT>
<SCRIPT type="text/javascript" src="scripts/paginar.js"></SCRIPT>
<SCRIPT type="text/javascript" src="scripts/Toolbar.js"></SCRIPT>
<SCRIPT type="text/javascript" src="Scripts/jquery/jquery-1.5.1.min.js"></SCRIPT>
<SCRIPT type="text/javascript" src="Scripts/jqueryUI/jquery-ui-1.8.15.custom.min.js"></SCRIPT>
<SCRIPT type="text/javascript" src="Scripts/botoneraPopUp.js"></SCRIPT>
<SCRIPT type="text/javascript" src="Scripts/jQueryPopUp.js"></SCRIPT>
<SCRIPT type="text/javascript" src="scripts/jQueryAutocomplete.js"></SCRIPT>
<SCRIPT type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></SCRIPT>
<SCRIPT type="text/javascript" src="scripts/MagicSearchObj.js"></SCRIPT>
<SCRIPT defer type="text/javascript" src="scripts/pngfix.js"></SCRIPT>      
<SCRIPT type="text/javascript" src="scripts/script_fechas.js"></SCRIPT>
<SCRIPT type="text/javascript">

	var ch = new channel();	
	var popUpPDC;
	
	function onLoadPage(){
		tb = new Toolbar('toolbar', 6,'images/compras/');
		tb.addButton("Home-16x16.png", "Home", "irHome()");		
<%  	if isAdminInAny() then  %>
		    tb.addButton("../add.gif", "Agregar Aseguradora", "irAseguradora()");		
<%		end if  %>
		tb.addButtonREFRESH("Recargar", "submitInfo()");		
		var swt = tb.addSwitcher("Search-16x16.png", "Buscar", "buscarOn()", "buscarOff()");		
		tb.draw();
		autocompleteAseguradora();	
		loadPDCvencidas();
		<%	if (cint(hayBusqueda) = 1) then %>				
				tb.changeState(swt);				
				startMagicSearch();
		<%	end if  %>
		<% 	if (not rsPDC.eof) then %>
				var pgn = new Paginacion("paginacion");
				pgn.paginar(<% =paginaActual %>, <% =lineasTotales %>, <% =mostrar %>, 50, "comprasPDCAdministrar.asp<% =params %>");
		<%	end if 	%>
		
		pngfix();	
	}
	
	function loadPDCvencidas() {
		$.ajax({
		  url: "comprasPDCAjax.asp?cdPedido=<%=gv_cdPedido%>&nroPoliza=<%=gv_NroPoliza%>&idAseguradora=<%=gv_idAseguradora%>&idTomador=<%=gv_Tomador%>&cdMoneda=<%=gv_Moneda%>&fecha=<%=gv_Fecha%>&idDivision=<%=gv_idDivision%>&importeAprox=<%=gv_ImporteAprox%>&importe=<%=gv_Importe%>&idEstado=<%=idEstado%>",
		  success: function(data){
		    document.getElementById("divPDCvencidas").innerHTML = data;
		  }
		});
	}
	
	function irAseguradora(){
		popUpPDC = winPopUp('Iframe', "comprasPDCAseguradoraPopUp.asp", "450", "200", '<%=GF_Traducir("Nueva Aseguradora")%>', 'submitInfo()');
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
		
	function irHome() {
		location.href = "comprasIndex.asp";
	}
	
	function submitInfo() {	
		document.getElementById("frmSel").submit();
	}	
		
	function setOrder(p_campo,p_orden){		
		
		document.getElementById("setOrder").value = p_campo;
		document.getElementById("setOrderTipo").value = p_orden;
		submitInfo();
	}
	function abrirPedido(idPedido){			
			window.open("comprasFichaPedidoCotizacion.asp?idPedido=" + idPedido + "&tab=1", "_blank", "location=no,scrollbars=yes,menubar=no,statusbar=no,height=500,width=500",false);
	}
	
	function anularPDC(idPoliza, img){
		if (confirm("Esta seguro que desea anular esta Poliza de Caucion?")) {
			img.src = "images/loading_small_green.gif"
			ch.bind("comprasPDCAjax.asp?IdPoliza=" + idPoliza + "&idEstado=<%=ESTADO_PDC_ANULADA%>", "anularPDCCallback('" + img.id + "')");
			ch.send();			
		}		
	}
	
	function devolverPDC(idPoliza, img){
		if (confirm("Esta seguro que desea devolver esta Poliza de Caucion?")) {
			img.src = "images/loading_small_green.gif"
			ch.bind("comprasPDCAjax.asp?IdPoliza=" + idPoliza + "&idEstado=<%=ESTADO_PDC_DEVUELTA%>", "devolverPDCCallback('" + img.id + "')");
			ch.send();			
		}		
	}
		
	function recibirPDC(idPoliza, idPedido){		
		popUpPDC = winPopUp('Iframe', "comprasPDCPopUp.asp?idPoliza=" + idPoliza + "&idPedido=" + idPedido, "400", "250", '<%=GF_Traducir("Completar PDC")%>', 'submitInfo()');
	}	
	
	function reloadPage(){
		window.location.reload();
	}	
		
	function anularPDCCallback(img){
		submitInfo();
	}
	
	function devolverPDCCallback(img){
		submitInfo();
	}
	
	function autocompleteAseguradora()
	{	
		$(function() {
		$( "#dsAseguradora" ).autocomplete({
		minLength: 2,
		source: function(request,response){
			$.ajax({
				url: "comprasStreamElementos.asp",
				dataType: "json",
			data: {				
				term : request.term,
				Tipo : "JQAseguradoras",
				DsLista : document.getElementById("dsAseguradora").value
				 },
		    success: function(data) {				
				response(data);
				}
			});	
		},		
		focus: function( event, ui ) {
				$( "#dsAseguradora").val(ui.item.descr);
				$( "#idAseguradora").val(ui.item.id);
				return false;
			},
		select: function( event, ui ) {
				$( "#dsAseguradora").val (ui.item.descr);				
				$( "#idAseguradora").val (ui.item.id);				
				return false;
			}		
		})
		.data( "autocomplete" )._renderItem = function( ul, item ) {
			return $( "<li></li>" )
			.data( "item.autocomplete", item )
			.append( "<a><font style='font-size:10;'>" + item.descr + "</font></a>" )
			.appendTo( ul );
			};
		});
	}
	
	function SeleccionarProveedor(ms){
		var desc = ms.getSelectedItem();
		if (desc.indexOf('-') != -1) {
			var arr = desc.split('-');
			document.getElementById("idTomador").value = arr[0];
			document.getElementById("dsTomador").value = arr[1];
			ms.setValue(arr[1]);
		} else {
			if (desc == ""){
				document.getElementById("idTomador").value = 0;
				document.getElementById("dsTomador").value = "";
				ms.setValue("");
			}	
		}				
	}
	
	
	
	
	
	function startMagicSearch(){	
		var msProveedor = new MagicSearch("", "companyName0", 30, 2, "comprasStreamElementos.asp?tipo=empresas");
		msProveedor.setMinChar(3);
		msProveedor.setToken(";");
		msProveedor.onBlur = SeleccionarProveedor;
		msProveedor.setValue('<% =gv_dsTomador %>');	
	}
		
</SCRIPT>

</HEAD>
<% call GF_TITULO2("kogge64.gif","Administracion de Poliza de Caucion") %>
<BODY onload="onLoadPage();">
<DIV id="toolbar"></DIV>
<FORM name="frmSel" id="frmSel" method="post" action="comprasPDCAdministrar.asp">	
	<DIV id="busqueda" class="divOculto">
	<BR><BR>
	<TABLE id="tblBusqueda" width="60%" cellspacing="0" cellpadding="0" align="center" border="0">
       <TR>
           <TD width="8"><IMG src="images/marco_r1_c1.gif"></TD>
           <TD width="25%"><IMG src="images/marco_r1_c2.gif" width="100%" height="8"></TD>
           <TD width="8"><IMG src="images/marco_r1_c3.gif"></TD>
           <TD width="75%"><TD>
           <TD></TD>
       </TR>
       <TR>
           <TD width="8"><IMG src="images/marco_r2_c1.gif"></TD>
           <TD align="center" valign="center"><font class="big" color="#517b4a"><% =GF_TRADUCIR("Busqueda") %></font></TD>
           <TD width="8"><IMG src="images/marco_r2_c3.gif"></TD>
           <TD align="right"></TD>
           <TD></TD>
       </TR>
       <TR>
           <TD><IMG src="images/marco_r2_c1.gif" height="8"  width="8"></TD>
           <TD></TD>
           <TD><IMG src="images/marco_c_s_d.gif" height="8" width="8"></TD>
           <TD><IMG src="images/marco_r1_c2.gif" width="100%" height="8"></TD>
           <TD width="8"><IMG src="images/marco_r1_c3.gif"></TD>
       </TR>
       <TR>
           <TD height="100%"><IMG src="images/marco_r2_c1.gif" height="100%" width="8"></TD>
           <TD colspan="3">
                     <TABLE width="95%" align="center" border="0">
                            <TR>
								<INPUT type="hidden" name="setOrder" id="setOrder" value=<% =setOrder %>>
								<INPUT type="hidden" name="setOrderTipo" id="setOrderTipo" value="<% =setOrderTipo %>">
								<TD width="15%" align="right"><% = GF_TRADUCIR("Cod. Pedido") %>:</TD>
								<TD width="25%">
									<INPUT type="text"  id="cdPedido" name="cdPedido" value="<%=gv_cdPedido%>">
								</TD>
								<TD width="13%" align="right"><% = GF_TRADUCIR("Nro Poliza") %>:</TD>
								<TD width="20%">
									<INPUT type="text"  id="NroPoliza" name="NroPoliza" value="<%=gv_NroPoliza%>">
								</TD>								
                            </TR>
                            <TR>
								<TD width="15%" align="right"><% = GF_TRADUCIR("Aseguradora") %>:</TD>
								<TD width="20%">
									<INPUT id="dsAseguradora" name="dsAseguradora" value="<%=gv_DsAseguradora%>">									
									<INPUT type="hidden" id="idAseguradora" name="idAseguradora" value="<%=gv_idAseguradora%>">
								</TD>								
								<TD width="13%" align="right"><% = GF_TRADUCIR("Tomador") %>:</TD>
								<TD width="20%">
									<DIV id="companyName0"></DIV>
									<INPUT type="hidden" id="idTomador" name="idTomador" value="<%=gv_Tomador%>">
									<INPUT type="hidden" id="dsTomador" name="dsTomador" value="<%=gv_dsTomador%>">
								</TD>								
							</TR>
							<TR>
								<TD align="right"><% = GF_TRADUCIR("Monto") %>: </TD>
								<TD>
									<INPUT type="text" size="12" name="importeTot" id="importeTot" value="<% if (gv_Importe <> 0) then response.write gv_Importe  %>" onKeyPress="return controlIngreso(this, event, 'I')">
									<INPUT type="radio" name="tipoMoneda" id="tipoMoneda" value="<%=MONEDA_PESO%>" <%if (gv_Moneda = MONEDA_PESO) then %>checked="checked"<%end if%> ><% = GF_TRADUCIR("$")%>
									<INPUT type="radio" name="tipoMoneda" id="tipoMoneda" value="<%=MONEDA_DOLAR%>" <%if (gv_Moneda = MONEDA_DOLAR) then %>checked="checked"<%end if%> ><% = GF_TRADUCIR("US$")%>
								</TD>
								<TD align="right"><% =GF_TRADUCIR("Vencimiento") %>:</TD>
								<TD >
									<INPUT type="text" size="1" maxLength="2" value="<% =gv_DVenc %>" name="txtDiaEmision" onBlur="javascript:ControlarDia(this);"> /
									<INPUT type="text" size="1" maxLength="2" value="<% =gv_MVenc %>" name="txtMesEmision" onBlur="javascript:ControlarMes(this);"> /
									<INPUT type="text" size="3" maxLength="4" value="<% =gv_AVenc %>" name="txtAnioEmision" onBlur="javascript:ControlarAnio(this);">
								</TD>								
							</TR>
							<TR>	
								<TD colspan=2 align="center">
									<INPUT type="radio" name="radio_Import" id="radio_Import" value="Menor" <%if (gv_ImporteAprox = "Menor") then %>checked="checked"<%end if%> ><% = GF_TRADUCIR("Menor")%>
									<INPUT type="radio" name="radio_Import" id="radio_Import" value="Igual" <%if (gv_ImporteAprox = "Igual") then %>checked="checked"<%end if%> ><% = GF_TRADUCIR("Igual")%>
									<INPUT type="radio" name="radio_Import" id="radio_Import" value="Mayor" <%if (gv_ImporteAprox = "Mayor") then %>checked="checked"<%end if%> ><% = GF_TRADUCIR("Mayor")%>
								</TD>
                            </TR>
                            <TR>
								<TD width="13%" align="right"><% = GF_TRADUCIR("Estado") %>:</TD>
								<TD width="20%">
									<SELECT name="idEstado" id="idEstado">
										<option VALUE="0"><%=GF_TRADUCIR("Seleccione...")%></option>
										<% =createOption(ESTADO_PDC_PENDIENTE, GF_TRADUCIR("Pendiente"), idEstado) %>
										<% =createOption(ESTADO_PDC_RECIBIDA, GF_TRADUCIR("Recibida"), idEstado) %>										
										<% =createOption(ESTADO_PDC_VENCIDA, GF_TRADUCIR("Vencida"), idEstado) %>										
										<% =createOption(ESTADO_PDC_DEVUELTA, GF_TRADUCIR("Devuelta"), idEstado) %>
										<% =createOption(ESTADO_PDC_ANULADA, GF_TRADUCIR("Anulada"), idEstado) %>
									</SELECT>
								</TD>                            
								<%
								strSQL = "Select divi.IDDIVISION, divi.DSDIVISION from TBLDIVISIONES divi"
								Call executeQueryDb(DBSITE_SQL_INTRA, rsDivisiones, "OPEN", strSQL)
								%>                                
                                <TD align="right"><% =GF_TRADUCIR("Division") %>:</TD>
                                <TD>                                
									<SELECT style="z-index:-1;" name="idDivision">
									        <option SELECTED value="<% =SIN_DIVISION %>">- <% =GF_TRADUCIR("Seleccione") %> -
									<%		while (not rsDivisiones.eof)		
												selected = ""										
												if (CLng(rsDivisiones("IDDIVISION")) = CLng(gv_idDivision)) then selected = "selected"
									%>
												<option value="<% =rsDivisiones("IDDIVISION") %>" <% =selected %>><% =rsDivisiones("DSDIVISION") %>                                        
									<%			rsDivisiones.MoveNext()
											wend 	
									%>
									</SELECT>
                                </TD>			
                            </TR>
							<TR>															
								<TD colspan="4"  align="center"><INPUT type="submit" value="Buscar" id="submit1" name="submit1" onclick='submitInfo();'></TD>
							</TR>		
								
                     </TABLE>
	           </TD>
	           <TD height="100%"><IMG src="images/marco_r2_c3.gif" width="8" height="100%"></TD>
	       </TR>
	       <TR>
	           <TD width="8"><IMG src="images/marco_r3_c1.gif"></TD>
	           <TD width="100%" align=center colspan="3"><IMG src="images/marco_r3_c2.gif" width="100%" height="8"></TD>
	           <TD width="8"><IMG src="images/marco_r3_c3.gif"></TD>
	       </TR>
	</TABLE>
	</DIV> 
	
	<INPUT type="hidden" name="busquedaActiva" id="busquedaActiva" value="0">
	<BR>
	<DIV id="divPDCvencidas" name="divPDCvencidas"></DIV>
	<BR>	
	<TABLE class="reg_header" width="100%" cellspacing="1" cellpadding="1" align="center" border="0">	
	
		<TR><TD colspan="11"><DIV id="paginacion"></DIV></TD></TR>
		<TR><TD colspan="11" align="center" class="reg_header_nav"><%= GF_TRADUCIR("Polizas vigentes")%></TD></TR>
		<TR>
	<% 	if (not rsPDC.eof) then %>
			<TD  class="reg_header_nav" width="11%" style="text-align: center">
				<IMG src="images\arrow_up.gif" onclick='setOrder(<%=BUSQUEDA_CDPEDIDO%>,"ASC")' style="cursor:pointer" title="Ascendente">
					&nbsp <%=GF_Traducir("Pedido")%>&nbsp 
				<IMG src="images\arrow_down.gif" onclick='setOrder(<%=BUSQUEDA_CDPEDIDO%>,"DESC")' style="cursor:pointer" title="Descendente">
			</TD>
			<TD  class="reg_header_nav" width="2%" align="center"></TD>
			<TD  class="reg_header_nav" width="10%" align="center">
				<IMG src="images\arrow_up.gif" onclick='setOrder(<%=BUSQUEDA_NROPDC%>,"ASC")' style="cursor:pointer" title="Ascendente">
					<%=GF_Traducir("Nro Poliza")%>
				<IMG src="images\arrow_down.gif" onclick='setOrder(<%=BUSQUEDA_NROPDC%>,"DESC")' style="cursor:pointer" title="Descendente">
			</TD>			
			<TD  class="reg_header_nav" width="30%" align="center">
				<IMG src="images\arrow_up.gif" onclick='setOrder(<%=BUSQUEDA_ASEGURADORA%>,"ASC")' style="cursor:pointer" title="Ascendente">
					<%=GF_Traducir("Aseguradora")%>
				<IMG src="images\arrow_down.gif" onclick='setOrder(<%=BUSQUEDA_ASEGURADORA%>,"DESC")' style="cursor:pointer" title="Descendente">
			</TD>
			<TD  class="reg_header_nav" width="23%" align="center">
				<IMG src="images\arrow_up.gif" onclick='setOrder(<%=BUSQUEDA_TOMADOR%>,"ASC")' style="cursor:pointer" title="Ascendente">
					<%=GF_Traducir("Tomador")%>
				<IMG src="images\arrow_down.gif" onclick='setOrder(<%=BUSQUEDA_TOMADOR%>,"DESC")' style="cursor:pointer" title="Descendente">
			</TD>
			<TD  class="reg_header_nav" width="10%" align="center">				
					<%=GF_Traducir("Monto")%>			
			</TD>
			<TD  class="reg_header_nav" width="10%" align="center">
				<IMG src="images\arrow_up.gif" onclick='setOrder(<%=BUSQUEDA_VENCIMIENTO%>,"ASC")' style="cursor:pointer" title="Ascendente">
					<%=GF_Traducir("Vencimiento")%>
				<IMG src="images\arrow_down.gif" onclick='setOrder(<%=BUSQUEDA_VENCIMIENTO%>,"DESC")' style="cursor:pointer" title="Descendente">
			</TD>
			<TD  class="reg_header_nav" width="3%" align="center">				
					<%=GF_Traducir(".")%>			
			</TD>
			<TD  class="reg_header_nav" width="3%" align="center">				
					<%=GF_Traducir(".")%>			
			</TD>
		</TR>		
		<% 		
				while ((not rsPDC.eof) and (CInt(reg) < CInt(mostrar)))				
					reg = reg + 1
					flagAdmin = isAdmin(rsPDC("IDDIVISION"))
					flagUser  = isUser(rsPDC("IDDIVISION"))	%>
					<TR>
						<TD class="reg_header_navdos" align="center"><%=rsPDC("CDPEDIDO")%></TD>						
						<TD style="text-align: center; cursor:pointer;" class="reg_header_navdos"><img onclick="abrirPedido(<% =rsPDC("IDPEDIDO") %>)" src="images/compras/PCT-16X16.png" title="Ver Ficha de Pedido"></td>
						<TD class="reg_header_navdos" align="center"><%=rsPDC("NROPOLIZA")%></TD>
						<TD class="reg_header_navdos" align="center"><%=rsPDC("DSASEGURADORA")%></TD>						
						<TD class="reg_header_navdos" align="center"><%=Trim(rsPDC("DSEMPRESA"))%></TD>						
						<TD class="reg_header_navdos" align="center"><%=getSimboloMoneda(rsPDC("CDMONEDA")) & " " & GF_EDIT_DECIMALS(rsPDC("IMPORTE"),2)%></TD>
						<TD class="reg_header_navdos" align="center"><%=GF_FN2DTE(rsPDC("VENCIMIENTO"))%></TD>
					<% if((rsPDC("ESTADO") = ESTADO_PDC_PENDIENTE)and((flagAdmin)or(flagUser)))then %>
						<TD class="reg_header_navdos" align="center"><IMG style="cursor:pointer;" title="<%=GF_TRADUCIR("Anular Poliza")%>" id="anular_<%=rsPDC("IDPDC")%>" src="images\compras\CTZ_cancel-16x16.png" onclick="anularPDC(<%=rsPDC("IDPDC")%>, this)"></TD>
						<TD class="reg_header_navdos" align="center"><IMG style="cursor:pointer;" title="<%=GF_TRADUCIR("Recibir Poliza")%>" id="recibir_<%=rsPDC("IDPDC")%>" src="images\almacenes\arrow_reception-16x16.png" onclick="recibirPDC(<%=rsPDC("IDPDC")%>,<%=rsPDC("IDPEDIDO")%>)"></TD>
					<% else %>
						<TD class="reg_header_navdos" align="center"></TD>
						<TD class="reg_header_navdos" align="center">
						<% if((rsPDC("ESTADO") = ESTADO_PDC_RECIBIDA)and((flagAdmin)or(flagUser)))then %>
							<IMG style="cursor:pointer;" title="<%=GF_TRADUCIR("Devolvel Poliza")%>" id="devolver_<%=rsPDC("IDPDC")%>" src="images\almacenes\arrow_loan-16x16.png" onclick="devolverPDC(<%=rsPDC("IDPDC")%>, this)">
						<% end if %>							
						</TD>
					<% end if %>						
					</TR>
					<% 
					rsPDC.movenext
				wend 
			else%>
			<TR class="TDNOHAY"><TD colSpan="4"><% =GF_TRADUCIR("No hay informacion disponible en estos momentos") %></TD></TR>		
			<%end if%>			
	</TABLE>
	<INPUT TYPE="HIDDEN" ID="accion" NAME="accion" VALUE=<%=ACCION_SUBMITIR%>>
</FORM>	
</BODY>
</HTML>