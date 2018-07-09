<!--#include file="../../Includes/procedimientosUnificador.asp"-->
<!--#include file="../../Includes/procedimientostraducir.asp"-->
<!--#include file="../../Includes/procedimientosfechas.asp"-->
<!--#include file="../../Includes/procedimientosformato.asp"-->
<!--#include file="../../Includes/procedimientosParametros.asp"-->
<!--#include file="../../Includes/procedimientosSQL.asp"-->
<!--#include file="../../Includes/procedimientos.asp"-->
<!--#include file="../../Includes/procedimientosExcel.asp"-->
<!--#include file="includes/procedimientosVIC.asp"-->
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
End Function
'********************************************************************
'					INICIO PAGINA
'********************************************************************
Dim division,verPagosEfectuados,pto,idcamion,search_radio,fecContable
Dim accion,nuCartaPorte1,nuCartaPorte2,nuCartaPorte3,fecContableD,fecContableM,fecContableA
Dim flagCall,cdProducto,cdVendedor,dsVendedor,cdCorredor,dsCorredor,cdCliente, dsCliente
Dim strSQLPro,rsProductos,cdEntregador,dsEntregador, fileCode
totalVagones = 0
totalKilosNetos = 0
Call GP_CONFIGURARMOMENTOS()

pto = GF_PARAMETROS7("pto", "", 6)
call addParam("pto", pto, params)
accion = GF_PARAMETROS7("accion", "", 6)
Call getParametros()
if not hayError() then	
	strSQL = generarSQL()	
	call GF_BD_Puertos(pto, rsGeneral, "OPEN", strSQL)
end if
%>
<html>
<head>
<title><%=GF_TRADUCIR("Puertos - Visteo Calada")%></title>
<link rel="stylesheet" href="../../css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="../../css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="../../css/iwin.css" type="text/css">
<link rel="stylesheet" href="../../css/MagicSearch.css" type="text/css">
<link rel="stylesheet" href="../../css/calendar-win2k-2.css" type="text/css">
<link rel="stylesheet" href="../../css/main.css" type="text/css"> 
<style type="text/css">
.labelStyle {
	font-weight: bold;
	text-align: center;
}
.numberStyle {
	font-weight: bold;
	font-size: 14px;
}
</style>
<script type="text/javascript" src="../../scripts/formato.js"></script>
<script type="text/javascript" src="../../scripts/channel.js"></script>
<script type="text/javascript" src="../../scripts/controles.js"></script>
<script type="text/javascript" src="../../scripts/Toolbar.js"></script>
<script type="text/javascript" src="../../scripts/MagicSearchObj.js"></script>

<script type="text/javascript">	
	var ch = new channel();		
	var changeFilters = false;
	function bodyOnLoad() {

	    tb = new Toolbar('toolbar');
	    tb.addButton("toolbar-excel", "Generar XLS", "Generar('XLS')");
	    tb.addButton("toolbar-excel", "Generar TxT", "Generar('TXT')");
	    tb.draw();		
		
		var msCoordinador = new MagicSearch("", "divCoordinador", 25, 4, "../puertosStreamElementos.asp?tipo=empresas&pto=<%=pto%>");
			msCoordinador.setToken(";");
			msCoordinador.minChar = 1			
			msCoordinador.onBlur = seleccionarCoordinador;
			msCoordinador.setValue('<% =myDsCoordinador %>');
		var msCoordinado = new MagicSearch("", "divCoordinado", 25, 4, "../puertosStreamElementos.asp?tipo=clientes&pto=<%=pto%>");
			msCoordinado.setToken(";");
			msCoordinado.minChar = 1			
			msCoordinado.onBlur = seleccionarCoordinado;
			msCoordinado.setValue('<% =myDsCoordinado %>');
		var msCorredor = new MagicSearch("", "divCorredor", 25, 4, "../puertosStreamElementos.asp?tipo=corredores&pto=<%=pto%>");
			msCorredor.setToken(";");
			msCorredor.minChar = 3
			msCorredor.onBlur = seleccionarCorredor;
			msCorredor.setValue('<% =myDsCorredor %>');
		var msVendedor = new MagicSearch("", "divVendedor", 25, 4, "../puertosStreamElementos.asp?tipo=vendedores&pto=<%=pto%>");
			msVendedor.setToken(";");
			msVendedor.minChar = 3
			msVendedor.onBlur = seleccionarVendedor;
			msVendedor.setValue('<% =myDsVendedor %>');
		var msVendedor = new MagicSearch("", "divEntregador", 25, 4, "../puertosStreamElementos.asp?tipo=entregadores&pto=<%=pto%>");
			msVendedor.setToken(";");
			msVendedor.minChar = 3
			msVendedor.onBlur = seleccionarEntregador;
			msVendedor.setValue('<% =myDsEntregador %>');
	}
	function Generar(pOpcion){
		document.getElementById("results").innerHTML = "";
		if (pOpcion=='XLS'){
			if (changeFilters) {  
				alert ("Atencion!\nSe cambiaron los filtros de búsqueda, por favor genere nuevamente el informe.");
				return 0;
			}
			window.open("vagonesInformeCamaraXLS.asp<%=params%>");		
			//ch.bind("vagonesInformeCamaraXLS.asp<%=params%>", "generate_Callback()");
			//ch.send();	
		}
		else if (pOpcion=='TXT'){
			if (changeFilters) {  
				alert ("Atencion!\nSe cambiaron los filtros de búsqueda, por favor genere nuevamente el informe.");
				return 0;
			}
			habilitarLoading("visible", "relative")
			document.getElementById("results").innerHTML = "";
			ch.bind("vagonesInformeCamaraTXT.asp<%=params%>", "generate_Callback()");
			ch.send();		
		}
		else{
		habilitarLoading("visible", "relative")
		
		document.getElementById("frmSel").submit();
		}
	}
	function generate_Callback(){
		habilitarLoading("hidden", "absolute")
		document.getElementById("actionLabel").innerHTML = "Archivo generado con exito<br>Click <u><a href='" + ch.response() + "' style='cursor:pointer;' >aqui</a></u> para ir al archivo.";
	}

	function volver() {	
		location.href = "../puertosReportes.asp?pto=<%=pto%>";
	}

	function seleccionarCoordinador(ms) {
		cambioBusqueda();	
		var desc = ms.getSelectedItem();
		if (desc.indexOf('|') != -1) {
			var arr = desc.split('|');
			document.getElementById("cdCoordinador").value = arr[0];
			document.getElementById("dsCoordinador").value = arr[1];
			ms.setValue(arr[1]);
		} else {
			if (desc == ""){
				document.getElementById("cdCoordinador").value = "";
				document.getElementById("dsCoordinador").value = "";
			}
		}		
	}		
	
	function seleccionarCoordinado(ms) {
		cambioBusqueda();	
		var desc = ms.getSelectedItem();
		if (desc.indexOf('|') != -1) {
			var arr = desc.split('|');
			document.getElementById("cdCoordinado").value = arr[0];
			document.getElementById("dsCoordinado").value = arr[1];
			ms.setValue(arr[1]);
		} else {
			if (desc == ""){
				document.getElementById("cdCoordinado").value = "";
				document.getElementById("dsCoordinado").value = "";
			}
		}		
	}	
	function seleccionarCorredor(ms) {
		cambioBusqueda();	
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
		cambioBusqueda();
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
	function seleccionarEntregador(ms) {
		cambioBusqueda();
		var desc = ms.getSelectedItem();
		if (desc.indexOf('|') != -1) {
			var arr = desc.split('|');
			document.getElementById("cdEntregador").value = arr[0];
			document.getElementById("dsEntregador").value = arr[1];
			ms.setValue(arr[1]);
		} else {
			if (desc == ""){
				document.getElementById("cdEntregador").value = "";
				document.getElementById("dsEntregador").value = "";
			}
		}		
	}	

	function habilitarLoading(pVisibility, pPosition){
		document.getElementById("imgLoading").style.position = pPosition;
		document.getElementById("imgLoading").style.visibility  = pVisibility;
		document.getElementById("lblLoading").style.position = pPosition;
		document.getElementById("lblLoading").style.visibility  = pVisibility;
		if (pVisibility=='visible')
			document.getElementById("actionLabel").style.visibility  = "hidden";
		else	
			document.getElementById("actionLabel").style.visibility  = "visible";
	}

	function lightOn(tr) {
		tr.className = "reg_Header_navdosHL";
	}
	
	function lightOff(tr) {
		tr.className = "reg_Header_navdos";
	}
	function cambioBusqueda(){
		changeFilters = true;
	}		
</script>
</head>
<body onLoad="bodyOnLoad()">	

<div id="toolbar"></div>

<form id="frmSel" name="frmSel" method="POST">	

<div class="tableaside size100"> <!-- BUSCAR -->
    <h3> filtro - <%=GF_Traducir("Informe Camara Vagones")%> </h3>
    
    <div id="searchfilter" class="tableasidecontent">
        <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Turno Desde") %> </div>
        <div class="col16"> <input type="text" onchange="cambioBusqueda();" id="Text1" maxLength="6" size="5" name="turnoDesde" value="<% =myTurnoDesde %>" onKeyPress="return controlIngreso (this, event, 'N');"> </div>
        <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Turno Hasta") %> </div>
        <div class="col16"> <input type="text" onchange="cambioBusqueda();" id="Text2" maxLength="6" size="5" name="turnoHasta" value="<% =myTurnoHasta %>" onKeyPress="return controlIngreso (this, event, 'N');"> </div>
        <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Incluir") %> </div>
        <div class="col16"> 
            <select onchange="cambioBusqueda();" name="Incluir">
				<option value="T"><% = GF_TRADUCIR("TODOS") %></option>
				<option value="C" <%if myIncluir="C" then Response.Write "SELECTED"%>><% = GF_TRADUCIR("CON ANALISIS") %></option>
				<option value="S" <%if myIncluir="S" then Response.Write "SELECTED"%>><% = GF_TRADUCIR("SIN ANALISIS") %></option>
			</select>
        </div>
        
        <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Operativo") %> </div>
        <div class="col16"> <input onchange="cambioBusqueda();" type="text" id="operativo" name="operativo" value="<% =myOperativo %>" onKeyPress="return controlIngreso (this, event, 'N');"> </div>
        <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Aceptación") %> </div>
        <div class="col16"> 
            <%
			 strSQLPro = "SELECT * FROM ACEPTACIONCALIDAD WHERE ICACEPTACION = 1 ORDER BY CDACEPTACION"
			 call GF_BD_Puertos(pto, rsAceptacion, "OPEN",strSQLPro)
			 'Response.Write "ACA(" & rsAceptacion.eof  & ")"
			 %>
			<select onchange="cambioBusqueda();" name="cdAceptacion" value="<%=myCdAceptacion%>">
				<option value=""> <%=GF_Traducir("TODAS")%></option>
				<%while not rsAceptacion.eof
					mySelected = ""
					if trim(rsAceptacion("CDACEPTACION")) = trim(myCdAceptacion) then mySelected = "SELECTED"%>
					<option value="<%=rsAceptacion("CDACEPTACION")%>" <%=mySelected%>> <%=rsAceptacion("DSACEPTACION")%></option>
					<%
					rsAceptacion.movenext
				 wend
				 %>
			</select>
        </div>
        
        <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Fecha Desde") %> </div>
        <div class="col16"> 
            <input type="text" onchange="cambioBusqueda();" size="1" maxLength="2" value="<% =myFecContableD%>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecContableD" id="fecContableD"> /
			<input type="text" onchange="cambioBusqueda();" size="1" maxLength="2" value="<% =myFecContableM %>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecContableM" id="fecContableM"> /
			<input type="text" onchange="cambioBusqueda();" size="2" maxLength="4" value="<% =myFecContableA %>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecContableA" id="fecContableA"> &nbsp;&nbsp;			
			<input type="text" onchange="cambioBusqueda();" size="1" maxLength="2" value="<% =myFecContableH %>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecContableH" id="fecContableH"> :
			<input type="text" onchange="cambioBusqueda();" size="1" maxLength="2" value="<% =myFecContableN %>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecContableN" id="fecContableN"> 
        </div>
        <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Sticker Desde") %> </div>
        <div class="col16"> <input type="text" onchange="cambioBusqueda();" id="stickerDesde" maxLength="10" size="11" name="stickerDesde" value="<% =myStickerDesde %>" onKeyPress="return controlIngreso (this, event, 'N');"> </div>
        
        <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Sticker Hasta") %> </div>
        <div class="col16"> <input type="text" onchange="cambioBusqueda();" id="stickerHasta" maxLength="10" size="11" name="stickerHasta" value="<% =myStickerHasta %>" onKeyPress="return controlIngreso (this, event, 'N');"> </div>
        <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Fecha Hasta") %> </div>
        <div class="col16"> 
            <input type="text" onchange="cambioBusqueda();" size="1" maxLength="2" value="<% =myFecContableDH%>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecContableDH" id="fecContableDH"> /
			<input type="text" onchange="cambioBusqueda();" size="1" maxLength="2" value="<% =myFecContableMH %>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecContableMH" id="fecContableMH"> /
			<input type="text" onchange="cambioBusqueda();" size="2" maxLength="4" value="<% =myFecContableAH %>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecContableAH" id="fecContableAH"> &nbsp;&nbsp;		
			<input type="text" onchange="cambioBusqueda();" size="1" maxLength="2" value="<% =myFecContableHH %>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecContableHH" id="fecContableHH"> :
			<input type="text" onchange="cambioBusqueda();" size="1" maxLength="2" value="<% =myFecContableNH %>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecContableNH" id="fecContableNH"> 
        </div>
        
        <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Coordinador") %> </div>
        <div class="col16"> 
            <div id="divCoordinador"></div>																		
			<input type="hidden" id="cdCoordinador" name="cdCoordinador" value="<%=myCdCoordinador%>">
			<input type="hidden" id="dsCoordinador" name="dsCoordinador" value="<%=myDsCoordinador%>">
        </div>
        <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Coordinado") %> </div>
        <div class="col16"> 
            <div id="divCoordinado"></div>																		
			<input type="hidden" id="cdCoordinado" name="cdCoordinado" value="<%=myCdCoordinado%>">
			<input type="hidden" id="dsCoordinado" name="dsCoordinado" value="<%=myDsCoordinado%>">
        </div>
        <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Producto") %> </div>
        <div class="col16"> 
            <%
			 strSQLPro = "SELECT * FROM PRODUCTOS ORDER BY DSPRODUCTO"
			 call GF_BD_Puertos(pto, rsProducto, "OPEN",strSQLPro)
			 %>
				<select onchange="cambioBusqueda();" name="cdProducto" value="<%=myCdProducto%>">
					<option value=""> <%=GF_Traducir("TODOS")%></option>
					<%while not rsProducto.eof
						mySelected = ""
						if trim(rsProducto("CDPRODUCTO")) = trim(myCdProducto) then mySelected = "SELECTED"%>
						<option value="<%=rsProducto("CDPRODUCTO")%>" <%=mySelected%>> <%=rsProducto("DSPRODUCTO")%></option>
						<%
						rsProducto.movenext
					 wend%>
			</select>		
        </div>
        
        <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Corredor") %> </div>
        <div class="col16"> 
            <div id="divCorredor"></div>																		
			<input type="hidden" id="cdCorredor" name="cdCorredor" value="<%=myCdCorredor%>">
			<input type="hidden" id="dsCorredor" name="dsCorredor" value="<%=myDsCorredor%>">
        </div>
        <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Vendedor") %> </div>
        <div class="col16"> 
            <div id="divVendedor"></div>																		
			<input type="hidden" id="cdVendedor" name="cdVendedor" value="<%=myCdVendedor%>">
			<input type="hidden" id="dsVendedor" name="dsVendedor" value="<%=myDsVendedor%>">
        </div>
        <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Entregador") %> </div>
        <div class="col16"> 
            <div id="divEntregador"></div>																		
			<input type="hidden" id="cdEntregador" name="cdEntregador" value="<%=myCdEntregador%>">
			<input type="hidden" id="dsEntregador" name="dsEntregador" value="<%=myDsEntregador%>">			            
        </div>
        
        <span class="btnaction"><input type="button" value="Buscar" id=submit1 name=submit1 onclick="Generar()"></span>
    </div>
</div><!-- END BUSCAR -->

<div class="col66"></div>

<% Call showErrors() %>

<div id="results">
<%
if not hayError() then	
	if not rsGeneral.eof then%>
		<TABLE class="datagrid" id="TAB1" align="center" width="100%" border="0">
		    <thead>
			<TR>
				<TH align="center">	<%=GF_Traducir("Turno")%> </TH>
				<TH align="center">	<%=GF_Traducir("Fecha")%> </TH>
				<!--<TD align="center">	<%=GF_Traducir("Coordinador")%> </TD>-->
				<TH align="center">	<%=GF_Traducir("Coordinado")%> </TH>
				<TH align="center">	<%=GF_Traducir("Producto")%> </TH>
				<TH align="center">	<%=GF_Traducir("Corredor")%> </TH>
				<TH align="center">	<%=GF_Traducir("Entregador")%> </TH>
				<TH align="center">	<%=GF_Traducir("Vendedor")%> </TH>
				<TH align="center">	<%=GF_Traducir("Localidad")%> </TH>
				<TH align="center">	<%=GF_Traducir("Nro Vagon")%> </TH>
				<TH align="center">	<%=GF_Traducir("Carta Porte")%> </TH>
				<TH align="center">	<%=GF_Traducir("Merma")%> </TH>
				<TH align="center">	<%=GF_Traducir("Kilos Netos")%> </TH>
				<TH align="center">	<%=GF_Traducir("Barras")%> </TH>
				<TH align="center">	<%=GF_Traducir("Grado")%> </TH>
				<TH align="center">	<%=GF_Traducir("Aceptacion")%> </TH>
				<TH align="center">	<%=GF_Traducir("Hora")%> </TH>
			</TR>
			</thead>
			<tbody>
			<%
			CargarGrados
			while not rsGeneral.eof
				myKilosNetos = Clng(rsGeneral("Bruto"))-Clng(rsGeneral("Tara"))
				myGradoParticular =  VerGrado (pto, rsGeneral("cdProducto"), rsGeneral("cdAceptacion"), rsGeneral("Barras"), rsGeneral("Fecha"),myIncluir)
				If myGradoParticular <> "XXX" Then
					totalVagones = totalVagones + 1	
					totalNetoAcumulado = totalNetoAcumulado + myKilosNetos
					call Sumar_Totales (myKilosNetos, totalVagones)
					call SumarResumen (myGradoParticular,myKilosNetos)
				End If
					%>
					<TR onMouseOver="javascript:lightOn(this)" onMouseOut="javascript:lightOff(this)">
						<TD align="center">	<%=rsGeneral("Turno")%> </TD>
						<TD align="center">	<%=GF_FN2DTE(Left(rsGeneral("DTPESADA"),8))%></TD>
						<!--<TD align="center">	<%=rsGeneral("Coordinador")%> </TD>-->
						<TD align="center">	<%=rsGeneral("Coordinado")%> </TD>
						<TD align="center">	<%=rsGeneral("Producto")%> </TD>
						<TD align="center">	<%=rsGeneral("Corredor")%> </TD>
						<TD align="center">	<%=rsGeneral("Entregador")%> </TD>
						<TD align="center">	<%=rsGeneral("Vendedor")%> </TD>
						<TD align="center">	<%=rsGeneral("Localidad")%> </TD>
						<TD align="center">	<%=rsGeneral("NoVagon")%> </TD>
						<TD align="center">	<%=rsGeneral("CartaPorte")%> </TD>
						<TD align="center">	<%=rsGeneral("Merma")%> </TD>
						<TD align="center">	<%=GF_EDIT_DECIMALS(cdbl(myKilosNetos),0)%> </TD>
						<TD align="center">	<%=rsGeneral("Barras")%> </TD>
						<TD align="center">	<%=myGradoParticular%> </TD>
						<TD align="center">	<%=rsGeneral("Aceptacion")%> </TD>
						<TD align="center">	<%=Right(GF_FN2DTE(rsGeneral("DTPESADA")),8)%></TD>
					</TR>
					<%
				rsGeneral.movenext
			wend	
			'Response.Write "<hr>EE1(" & totalVagones & ")2(" & totalNetoAcumulado & ")"
				%>
				</tbody>
				<tfoot>
            	<tr>
					<td colspan=9><%=GF_Traducir("TOTAL DE VAGONES ") & totalVagones%></td>
					<td colspan=2><%=GF_Traducir("TOTAL DE KILOS")%></td>
					<td colspan=1 align="center"><B><%=GF_EDIT_DECIMALS(cdbl(totalNetoAcumulado),0)%></td>
					<td colspan=4>&nbsp;</td>
				</tr>
				</tfoot>
            </TABLE>	    		
            <div class="col66"></div>		
						<table class="datagrid" align="center" width="50%" >
						    <thead>
						    <tr>
						        <th colspan="5" align="center">RESUMEN</th>
						    </tr>	
							<tr>
								<!--<td>&nbsp;</td>-->
								<td width="40%" align="center" rowspan="2"><b><%=GF_Traducir("Items")%></b></td>
								<td align="center" colspan="2"><B><%=GF_Traducir("VAGONES")%></B></td>
								<td align="center" colspan="2"><B><%=GF_Traducir("KILOGRAMOS")%></B></td>
							</tr>	
							<tr>
								
								<td width="15%" align="center"><b><%=GF_Traducir("Cantidad")%></b></td>
								<td width="15%" align="center"><b><%=GF_Traducir("Porcentaje")%></b></td>
								<td width="15%" align="center"><b><%=GF_Traducir("Cantidad")%></b></td>
								<td width="15%" align="center"><b><%=GF_Traducir("Porcentaje")%></b></td>
							</tr>	
							</thead>
							<tbody>
							<%
							call SumarPorcentajeResumen(totalVagones, totalNetoAcumulado, totalVagonesRegistrados, totalKilosNetosRegistrados)
							For i = 0 To 13
								Response.Write "<tr class='reg_Header_navdos'>"
									Response.Write "<td colspan=1>" & myGrado(i,1) & "</td>"
									Response.Write "<td align='right' colspan=1>" & myGrado(i,2) & "</td>"
									Response.Write "<td align='right' colspan=1>" & myGrado(i,3) & "</td>"
									Response.Write "<td align='right' colspan=1>" & GF_EDIT_DECIMALS(cdbl(myGrado(i,4)),0) & "</td>"
									Response.Write "<td align='right' colspan=1>" & myGrado(i,5) & "</td>"
								Response.Write "</tr>"
							Next
							%>
							</tbody>
							<tfoot>
							<tr>
								<td colspan=1><%=GF_Traducir("TOTAL")%></td>
								<td align="right"><b><%=totalVagonesRegistrados%></b></td>
								<td align="right"><b><%=GF_EDIT_DECIMALS(10000,2)%></b></td>
								<td align="right"><b><%=GF_EDIT_DECIMALS(totalKilosNetosRegistrados,0)%></b></td>
								<td align="right"><b><%=GF_EDIT_DECIMALS(10000,2)%></b></td>		
							</tr>
							</tfoot>
						</table>
				</tbody>
		
<%
	end if	
end if%>
</div>	
<br>
		<table align="center" width="90%" border="0">
			<tr>
				<td align="center">
					<img style="position:absolute;visibility:hidden;" id="imgLoading" src="../images/Loading4.gif">
					<div style="position:absolute;visibility:hidden;" id="lblLoading"><b><br>Aguarde por favor...</b></div>
					
				</td>
			</tr>
		</table>             

<div align="center"><div id="actionLabel" class="round_border_all TDSUCCESS" style="width:80%;visibility:hidden;"></div></div>
</form>
</body>
</html>
