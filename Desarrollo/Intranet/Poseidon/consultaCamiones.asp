<!--#include file="consultaCamionesCommon.asp"-->

<HTML>
<HEAD>
<title>Consulta de Camiones</title>

<meta http-equiv="X-UA-Compatible" content="IE=Edge">

<link rel="stylesheet" href="../css/main.css" type="text/css">
<link rel="stylesheet" href="../css/tables.css" type="text/css">
<link rel="stylesheet" href="../css/paginar.css" type="text/css">
<link rel="stylesheet" href="../css/MagicSearch.css" type="text/css">
<link rel="stylesheet" href="../css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />

<script type="text/javascript" src="../scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="../scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="../scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript" src="../scripts/channel.js"></script>
<script type="text/javascript" src="../scripts/formato.js"></script>
<script type="text/javascript" src="../scripts/controles.js"></script>
<script type="text/javascript" src="../scripts/paginar.js"></script>
<script type="text/javascript" src="../scripts/MagicSearchObj.js"></script>
<script type="text/javascript">
	
	function abrirInfoAnalisis(p_idCamion, p_dtContable, p_ctaPorte){		
		myPopUp = new winPopUp('Iframe', 'InfoAnalisisCamion.asp?Pto=<%=pto%>&Sector=<%=g_strSector%>&Camion=' + p_idCamion + '&dtContable=' + p_dtContable + '&ctaPorte=' + p_ctaPorte , '780', '550', 'Informacion de Camiones');		
	}
	
	function dimensionarIframe(p_width, p_height){
		myPopUp.resize(p_width, p_height);
	}

	function onLoadPage(){
	<%  if (isToepfer(session("KCOrganizacion"))) then %>
			var msCliente = new MagicSearch("", "divCliente", 25, 4, "puertosStreamElementos.asp?tipo=clientes&pto=<%=pto%>");
			msCliente.setToken(";");
			msCliente.minChar = 3			
			msCliente.onBlur = seleccionarCliente;
			msCliente.setValue('<% =dsCliente %>');
			var msCorredor = new MagicSearch("", "divCorredor", 25, 4, "puertosStreamElementos.asp?tipo=corredores&pto=<%=pto%>");
			msCorredor.setToken(";");
			msCorredor.minChar = 3
			msCorredor.onBlur = seleccionarCorredor;
			msCorredor.setValue('<% =dsCorredor %>');
			var msVendedor = new MagicSearch("", "divVendedor", 25, 4, "puertosStreamElementos.asp?tipo=vendedores&pto=<%=pto%>");
			msVendedor.setToken(";");
			msVendedor.minChar = 3
			msVendedor.onBlur = seleccionarVendedor;
			msVendedor.setValue('<% =dsVendedor %>');
    <% end if %>			
			var msChofer = new MagicSearch("", "divChofer", 25, 4, "puertosStreamElementos.asp?tipo=choferes&pto=<%=pto%>");
			msChofer.setToken(";");
			msChofer.minChar = 3
			msChofer.onBlur = seleccionarChofer;
			msChofer.setValue('<% =dsChofer %>');
			var msTransportista = new MagicSearch("", "divTransportista", 25, 4, "puertosStreamElementos.asp?tipo=transportistas&pto=<%=pto%>");
			msTransportista.setToken(";");
			msTransportista.minChar = 3
			msTransportista.onBlur = seleccionarTransportista;
			msTransportista.setValue('<% =dsTransportista %>');
			<% if not hayError() then 
				if (not rsLista.eof) then		%>								
					var pgn = new Paginacion("paginacion");							
					pgn.paginar(<% =paginaActual %>, <% =lineasTotales %>, <% =mostrar %>, 200, "consultacamiones.asp<% =params %>");
			<%	end if 
			   end if
			%>
			
	}
	function generateExcel() {
		document.getElementById("frmSel").action = "consultaCamionesPrintXLS.asp";
		document.getElementById("accion").value = '<% =ACCION_PROCESAR %>';
		document.getElementById("frmSel").submit();
	}
	function seleccionarCliente(ms) {
		var desc = ms.getSelectedItem();
		if (desc.indexOf('|') != -1) {
			var arr = desc.split('|');
			document.getElementById("cdCliente").value = arr[0];
			document.getElementById("dsCliente").value = arr[1];
			ms.setValue(arr[1]);
		} else {
			if (desc == ""){
				document.getElementById("cdCliente").value = "";
				document.getElementById("dsCliente").value = "";
			}
		}		
	}		
	function seleccionarCorredor(ms) {
		var desc = ms.getSelectedItem();
		if (desc.indexOf('|') != -1) {
			var arr = desc.split('|');
			document.getElementById("cdCorredor").value = arr[0];
			document.getElementById("dsCorredor").value = arr[1];
			ms.setValue(arr[1]);
		} else {
			if (desc == ""){
				document.getElementById("cdCorredor").value = "";
				document.getElementById("dsCorredor").value = "";
			}
		}		
	}		
	function seleccionarVendedor(ms) {
		var desc = ms.getSelectedItem();
		if (desc.indexOf('|') != -1) {
			var arr = desc.split('|');
			document.getElementById("cdVendedor").value = arr[0];
			document.getElementById("dsVendedor").value = arr[1];
			ms.setValue(arr[1]);
		} else {
			if (desc == ""){
				document.getElementById("cdVendedor").value = "";
				document.getElementById("dsVendedor").value = "";
			}
		}		
	}	
	function seleccionarChofer(ms) {
		var desc = ms.getSelectedItem();
		if (desc.indexOf('|') != -1) {
			var arr = desc.split('|');
			document.getElementById("cdChofer").value = arr[0];
			document.getElementById("dsChofer").value = arr[1];
			ms.setValue(arr[1]);
		} else {
			if (desc == ""){
				document.getElementById("cdChofer").value = "";
				document.getElementById("dsChofer").value = "";
			}
		}		
	}		
	function seleccionarTransportista(ms) {
		var desc = ms.getSelectedItem();
		if (desc.indexOf('|') != -1) {
			var arr = desc.split('|');
			document.getElementById("cdTransportista").value = arr[0];
			document.getElementById("dsTransportista").value = arr[1];
			ms.setValue(arr[1]);
		} else {
			if (desc == ""){
				document.getElementById("cdTransportista").value = "";
				document.getElementById("dsTransportista").value = "";
			}
		}		
	}	
	function submitInfo() {	
	    document.getElementById("frmSel").action = "consultaCamiones.asp";
		document.getElementById("accion").value = '';	
		document.getElementById("frmSel").submit();
	}		
	function setSortBy(pTxt){
		document.getElementById("sortBy").value = pTxt;
		submitInfo();
	}
	function abrirNotaRecepcion(pIdCamion,pCartaPorte,pDtContable,pTipoCamion){
	    window.open("NotaRecepcionPrint.asp?pto=<%=g_strPuerto%>&idCamion="+pIdCamion+"&cartaPorte="+pCartaPorte+"&dtContable="+pDtContable+"&tipoCamion="+pTipoCamion);
	}	
    function editCtaPte(p_CtaPte, p_DtContable, p_idCamion){
        window.open("cartadeporteEdit.asp?pto=<%=g_strPuerto%>&cartaPorte="+p_CtaPte+"&dtContable="+p_DtContable + "&idcamion=" + p_idCamion);
    }
</script>

</HEAD>
<BODY onload="onLoadPage();">
<form name="frmSel" id="frmSel" method="post" action="consultaCamiones.asp">
	<div class="tableaside size100"> <!-- BUSCAR -->
		<h3> filtro - Consulta Facturaci&oacuten Calidad - <% =g_strPuerto %></h3>
		
		<div id="searchfilter" class="tableasidecontent">
						 
			<div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Id Camion") %>: </div>
			<div class="col16"> <input type="text" SIZE="9" MAXLENGTH="10" id="idCamion" name="idCamion" onKeyPress="return controlIngreso (this, event, 'N');" value="<% =idCamion %>"> </div>
			
			<div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Fec. Descarga Desde") %>: </div>
			<div class="col16"> 
				<input type="text" size="1" maxLength="2" value="<% =fecContableD%>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecContableD"> /
				<input type="text" size="1" maxLength="2" value="<% =fecContableM %>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecContableM"> /
				<input type="text" size="2" maxLength="4" value="<% =fecContableA %>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecContableA"><br>
			</div>
			
			<div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Fec. Descarga Hasta") %>: </div>
			<div class="col16"> 
				<input type="text" size="1" maxLength="2" value="<% =fecContableHD%>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecContableHD"> /
				<input type="text" size="1" maxLength="2" value="<% =fecContableHM %>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecContableHM"> /
				<input type="text" size="2" maxLength="4" value="<% =fecContableHA %>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecContableHA">
			</div>
						
			<div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Nro C. Porte") %>: </div>
			<div class="col16"> <input type="text" SIZE="2" MAXLENGTH="4" id="nuCartaPorte1" name="nuCartaPorte1" onKeyPress="return controlIngreso (this, event, 'N');" value="<% =nuCartaPorte1 %>">-<input type="text" SIZE="7" MAXLENGTH="8" id="nuCartaPorte2" onKeyPress="return controlIngreso (this, event, 'N');" name="nuCartaPorte2" value="<% =nuCartaPorte2 %>"> </div>
			
			<div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Pat. Chasis") %>: </div>
			<div class="col16"> <input type="text" SIZE="2" MAXLENGTH="3" id="patChasis1" name="patChasis1" value="<% =patChasis1 %>">-<input type="text" SIZE="2" MAXLENGTH="3" id="patChasis2" name="patChasis2" onKeyPress="return controlIngreso (this, event, 'N');" value="<% =patChasis2 %>"> </div>						
			
			<div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Pat. Acoplado") %>: </div>
			<div class="col16"> <input type="text" SIZE="2" MAXLENGTH="3" id="patAcoplado1" name="patAcoplado1" value="<% =patAcoplado1 %>">-<input type="text" SIZE="2" MAXLENGTH="3" id="patAcoplado2" name="patAcoplado2" onKeyPress="return controlIngreso (this, event, 'N');" value="<% =patAcoplado2 %>"> </div>					
			
			<div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Chofer") %>: </div>
			<div class="col16"> 
				<div id="divChofer"></div>																		
				<input type="hidden" id="cdChofer" name="cdChofer" value="<%=cdChofer%>">
				<input type="hidden" id="dsChofer" name="dsChofer" value="<%=dsChofer%>">
			</div>
			
			<div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Cod. Cupo") %>: </div>
			<div class="col16"> <input type="text" SIZE="10" MAXLENGTH="11" id="codCupo" name="codCupo" value="<% =codCupo %>"> </div>
				
			<div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Tipo Camión") %>: </div>
			<div class="col16"> 
				<select name="cdCircuito">
						<option value="<% =CIRCUITO_CAMION_TODOS %>"> <%=GF_Traducir("Seleccionar...")%></option>
						<option value="<% =CIRCUITO_CAMION_CARGA %>" <% if (cdCircuito = CIRCUITO_CAMION_CARGA)  then response.write "selected" %>> <%=GF_Traducir("CARGA")%></option>
						<option value="<% =CIRCUITO_CAMION_DESCARGA %>" <% if (cdCircuito = CIRCUITO_CAMION_DESCARGA)  then response.write "selected" %>> <%=GF_Traducir("DESCARGA")%></option>
				</select>
			</div>
			
			<div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Producto") %>: </div>
			<div class="col16"> 
				<%									
				Call executeQueryDb(pto, rsProductos, "OPEN", "SELECT * FROM PRODUCTOS ORDER BY CDPRODUCTO")								
				%>
				<select name="cdProducto" value="<%=cdProducto%>">
					<option value="0"> <%=GF_Traducir("Seleccionar...")%></option>
				<%
					while not rsProductos.eof
						mySelected = ""
						if cint(rsProductos("CDPRODUCTO")) = cint(cdProducto) then mySelected = "SELECTED"
						%>
						<option value="<%=rsProductos("CDPRODUCTO")%>" <%=mySelected%>> <%=rsProductos("DSPRODUCTO")%></option>
						<%	
						rsProductos.movenext
					wend
				%>
				</select>
			</div>											
								
			<div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Solo c/Muestras Extra") %>: </div>
			<div class="col16"> <input type="checkbox" name="chkMuestrasAud" value="<% =MUESTRAS_AUDITORIA_ONLY %>" <% if (chkMuestrasAud = MUESTRAS_AUDITORIA_ONLY) then response.write "checked" %>/> </div>
									
			<div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Estado") %>: </div>
			<div class="col16"> 
				<%
				Call executeQueryDb(pto, rsEstados, "OPEN", "SELECT * FROM ESTADOS ORDER BY CDESTADO")										
				%>
				<select name="cdEstado" value="<%=cdEstado%>">
					<option value="0"> <%=GF_Traducir("Seleccionar...")%></option>
				<%
					while not rsEstados.eof
						mySelected = ""
						if cint(rsEstados("CDESTADO")) = cint(cdEstado) then mySelected = "SELECTED"
						%>
						<option value="<%=rsEstados("CDESTADO")%>" <%=mySelected%>> <%=rsEstados("DSESTADO")%></option>
						<%	
						rsEstados.movenext
					wend
				%>
				</select>
			</div>									
			
			<div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Transportista") %>: </div>
			<div class="col16"> 
				<div id="divTransportista"></div>																		
				<input type="hidden" id="cdTransportista" name="cdTransportista" value="<%=cdTransportista%>">
				<input type="hidden" id="dsTransportista" name="dsTransportista" value="<%=dsTransportista%>">
			</div>
						
			<% if (isToepfer(session("KCOrganizacion"))) then %>
			
				<div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Cliente") %>: </div>
				<div class="col16"> 
					<div id="divCliente"></div>																		
					<input type="hidden" id="cdCliente" name="cdCliente" value="<%=cdCliente%>">
					<input type="hidden" id="dsCliente" name="dsCliente" value="<%=dsCliente%>">
				</div>
			
				<div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Corredor") %>: </div>
				<div class="col16"> 
					<div id="divCorredor"></div>																		
					<input type="hidden" id="cdCorredor" name="cdCorredor" value="<%=cdCorredor%>">
					<input type="hidden" id="dsCorredor" name="dsCorredor" value="<%=dsCorredor%>">
				</div>
			
				<div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Vendedor") %>: </div>
				<div class="col16"> 
					<div id="divVendedor"></div>																		
					<input type="hidden" id="cdVendedor" name="cdVendedor" value="<%=cdVendedor%>">
					<input type="hidden" id="dsVendedor" name="dsVendedor" value="<%=dsVendedor%>">
				</div>
			
			<% end if %>
			
			<span class="btnaction">
				<input type="submit" value="Buscar"> 
				<input type="button" value="Exportar" onclick="generateExcel();">
			</span>
			
		</div>
					
			
	</div><!-- END BUSCAR -->
	<INPUT TYPE="hidden" NAME="pto" VALUE="<%=pto%>">
	<INPUT SIZE=50 TYPE="hidden" ID="sortBy" NAME="sortBy" VALUE="<%=sortBy%>">
	<INPUT TYPE="hidden" NAME="accion" id="accion" VALUE="">
</form>
<br>
	<%
	if hayError() then 	
	%>
		<table width="90%" cellspacing="0" cellpadding="0" align="center" border="0">
			<tr>
				<td colspan=3>
					<%			
					call showErrors()
					%>
				</td>
			</tr>	
		</table>		
	<%
	else
		Call crearTabla(mostrar, accion)
	end if
	%>
	
	
</BODY>
</HTML>
