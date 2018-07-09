<!--#include file="../../Includes/procedimientosUnificador.asp"-->
<!--#include file="../../Includes/procedimientostraducir.asp"-->
<!--#include file="../../Includes/procedimientosfechas.asp"-->
<!--#include file="../../Includes/procedimientosformato.asp"-->
<!--#include file="../../Includes/procedimientosParametros.asp"-->
<!--#include file="../../Includes/procedimientosSQL.asp"-->
<!--#include file="../../Includes/procedimientos.asp"-->
<!--#include file="../../Includes/procedimientosPuertos.asp"-->
<%
Const PDF_REPORT = "PDF"
Const XLS_REPORT = "XLS"

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
'********************************************************************
'					INICIO PAGINA
'********************************************************************
dim  division,verPagosEfectuados,pto,idcamion,search_radio,params,fecContable
dim accion,nuCartaPorte1,nuCartaPorte2,nuCartaPorte3,fecContableD,fecContableM,fecContableA
dim flagCall,cdProducto,cdVendedor,dsVendedor,cdCorredor,dsCorredor,cdCliente, dsCliente
dim strSQLPro,rsProductos,cdEntregador,dsEntregador, fileCode, cdEstado, cdTransporte

Call GP_CONFIGURARMOMENTOS()

ty = GF_PARAMETROS7("ty", "", 6)

pto = GF_PARAMETROS7("pto", "", 6)
call addParam("pto", pto, params)
idcamion = GF_PARAMETROS7("idcamion", 0, 6)
idcamion = GF_nDigits(idcamion, 10)
call addParam("idcamion", idcamion, params)

accion = GF_PARAMETROS7("accion", "", 6)

fecContableD = GF_PARAMETROS7("fecContableD", "", 6)
if (fecContableD = "") then fecContableD=Day(Now())
call addParam("fecContableD", fecContableD, params)

fecContableM = GF_PARAMETROS7("fecContableM", "", 6)
if (fecContableM = "") then fecContableM=Month(Now())
Call addParam("fecContableM", fecContableM, params)

fecContableA = GF_PARAMETROS7("fecContableA", "", 6)
if (fecContableA = "") then fecContableA=Year(Now())
Call addParam("fecContableA", fecContableA, params)


fecContableDH = GF_PARAMETROS7("fecContableDH", "", 6)
if (fecContableDH = "") then fecContableDH=Day(Now())
call addParam("fecContableDH", fecContableDH, params)

fecContableMH = GF_PARAMETROS7("fecContableMH", "", 6)
if (fecContableMH = "") then fecContableMH=Month(Now())
call addParam("fecContableMH", fecContableMH, params)

fecContableAH = GF_PARAMETROS7("fecContableAH", "", 6)
if (fecContableAH = "") then fecContableAH=Year(Now())
call addParam("fecContableAH", fecContableAH, params)


nuCartaPorte1 = GF_PARAMETROS7("nuCartaPorte1", "", 6)
if (nuCartaPorte1 <> "") then nuCartaPorte1 = GF_nDigits(nuCartaPorte1, 4)
call addParam("nuCartaPorte1", nuCartaPorte1, params)
nuCartaPorte2 = GF_PARAMETROS7("nuCartaPorte2", "", 6)
if (nuCartaPorte2 <> "") then nuCartaPorte2 = GF_nDigits(nuCartaPorte2, 8)
call addParam("nuCartaPorte2", nuCartaPorte2, params)
nuCartaPorte3 = GF_PARAMETROS7("nuCartaPorte3", "", 6)
if (nuCartaPorte3 <> "") then nuCartaPorte3 = GF_nDigits(nuCartaPorte3, 4)
call addParam("nuCartaPorte3", nuCartaPorte3, params)
'------------------------Nuevos Filtros----------------------
cdProducto = GF_PARAMETROS7("cdProducto", 0, 6)
call addParam("cdProducto", cdProducto, params)

cdVendedor = GF_PARAMETROS7("cdVendedor", "", 6)
call addParam("cdVendedor", cdVendedor, params)
dsVendedor = GF_PARAMETROS7("dsVendedor", "", 6)
call addParam("dsVendedor", dsVendedor, params)

cdCorredor = GF_PARAMETROS7("cdCorredor", "", 6)
call addParam("cdCorredor", cdCorredor, params)
dsCorredor = GF_PARAMETROS7("dsCorredor", "", 6)
call addParam("dsCorredor", dsCorredor, params)

cdCliente = GF_PARAMETROS7("cdCliente", "", 6)
call addParam("cdCliente", cdCliente, params)
dsCliente = GF_PARAMETROS7("dsCliente", "", 6)
call addParam("dsCliente", dsCliente, params)

cdEntregador = GF_PARAMETROS7("cdEntregador", "", 6)
call addParam("cdEntregador", cdEntregador, params)
dsEntregador = GF_PARAMETROS7("dsEntregador", "", 6)
call addParam("dsEntregador", dsEntregador, params)

cdEstado = GF_PARAMETROS7("estado", 0, 6)
call addParam("estado", cdEstado, params)

cdTransporte = GF_PARAMETROS7("transporte", 0, 6)
if (cdTransporte = 0) then cdTransporte=TIPO_TRANSPORTE_CAMION
call addParam("transporte", cdTransporte, params)


'---------------------------------------------------------
fileCode = GF_PARAMETROS7("fileCode", "", 6)
if (fileCode = "") then fileCode = session("Usuario") & "_" & session("MMTODATO")

flagCall=false
if (accion = ACCION_SUBMITIR) then
	'CONTROLAR!!!!
	ret = GF_CONTROL_PERIODO(fecContableD, fecContableDH, fecContableM, fecContableMH, fecContableA, fecContableAH)
	Select case (ret)
	case 0
		'Si el control resulto exitoso
		flagCall=true
	case 1
		Call setError(FECHA_INICIO_INCORRECTA)
	case 2
		Call setError(FECHA_FIN_INCORRECTA)
	case 3
		Call setError(PERIODO_ERRONEO)
	end select
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
<script defer type="text/javascript" src="../../scripts/pngfix.js"></script>

<script type="text/javascript">	
		
	function bodyOnLoad() {			
		tb = new Toolbar('toolbar', 6,'images/');
		tb.addButton("DocumentoTexto-16x16.png", "Imprimir PDF (Solo Camiones)", "GenerarPDF()");
		tb.addButton("excel.gif", "Imprimir XLS", "GenerarXLS()");
		tb.draw();		
		pngfix();
		var msCliente = new MagicSearch("", "divCliente", 25, 4, "../puertosStreamElementos.asp?tipo=clientes&pto=<%=pto%>");
			msCliente.setToken(";");
			msCliente.minChar = 3			
			msCliente.onBlur = seleccionarCliente;
			msCliente.setValue('<% =dsCliente %>');
		var msCorredor = new MagicSearch("", "divCorredor", 25, 4, "../puertosStreamElementos.asp?tipo=corredores&pto=<%=pto%>");
			msCorredor.setToken(";");
			msCorredor.minChar = 3
			msCorredor.onBlur = seleccionarCorredor;
			msCorredor.setValue('<% =dsCorredor %>');
		var msVendedor = new MagicSearch("", "divVendedor", 25, 4, "../puertosStreamElementos.asp?tipo=vendedores&pto=<%=pto%>");
			msVendedor.setToken(";");
			msVendedor.minChar = 3
			msVendedor.onBlur = seleccionarVendedor;
			msVendedor.setValue('<% =dsVendedor %>');
		var msEntregador = new MagicSearch("", "divEntregador", 25, 4, "../puertosStreamElementos.asp?tipo=entregadores&pto=<%=pto%>");
			msEntregador.setToken(";");
			msEntregador.minChar = 3
			msEntregador.onBlur = seleccionarEntegador;
			msEntregador.setValue('<% =dsEntregador %>');
	<%	if (flagCall) then 
			if (ty = XLS_REPORT) then	%>
			generateXLS();
	<%		else	%>
			window.open("reporteVisteosCaladaPrint.asp<% =params %>");
	<%		end if
		end if %>
	}

	function GenerarPDF() {			
		document.getElementById("frmSel").action="reporteVisteosCalada.asp";
		document.getElementById("frmSel").target="";
		document.getElementById("ty").value = "<% =PDF_REPORT %>";
		document.getElementById("frmSel").submit();
	}
	
	function GenerarXLS() {			
		document.getElementById("frmSel").action="reporteVisteosCalada.asp";
		document.getElementById("frmSel").target="";
		document.getElementById("ty").value = "<% =XLS_REPORT %>";
		document.getElementById("frmSel").submit();
	}
	
	var maxSegments;
	var currSegment=0;
	var MS_X_DAY = 86400000 //Milisegundos por día.	
	
	function calculateSegments() {
		var d = document.getElementById("fecContableD").value;
		var m = document.getElementById("fecContableM").value-1; //El Month de Date trabaja de 0 a 11
		var y = document.getElementById("fecContableA").value;
		var fd = new Date(y, m, d, 0, 0, 0, 0);		
		d = document.getElementById("fecContableDH").value;
		m = document.getElementById("fecContableMH").value-1; //El Month de Date trabaja de 0 a 11
		y = document.getElementById("fecContableAH").value;
		var fh = new Date(y, m, d, 0, 0, 0, 0);		
		maxSegments = Math.round((fh.getTime() - fd.getTime())/MS_X_DAY)		
	}
	
	function generateExcel() {		
	    var tran = document.getElementById("transporte").value;
		document.getElementById("actionLabel").innerHTML = "Generando Excel...";
		setTimeout("document.getElementById('actionLabel').style.visibility = 'hidden'", 3000);
	    document.getElementById("frmSel").action="reporteVisteosCaladaPrintE2.asp";
		document.getElementById("frmSel").target="";
		document.getElementById("frmSel").submit();		
	}
	
	function generateSegment_callback() {		
		if (currSegment < maxSegments) {
			currSegment += 1; 
			generateSegment(currSegment);
		} else {			
			generateExcel();
		}
	}
	
	function generateSegment(currSegment) {
		document.getElementById("actionLabel").innerHTML = "Recopilando datos...  ( " + (currSegment+1) + " / " + (maxSegments+1) + " )";
		var d = document.getElementById("fecContableD").value;
		var m = document.getElementById("fecContableM").value-1; //El Month de Date trabaja de 0 a 11
		var y = document.getElementById("fecContableA").value;
		var fd = new Date(y, m, d, 0, 0, 0, 0);		
		var d = new Date(fd.getTime() + (MS_X_DAY*currSegment));
		document.getElementById("fecContableDS").value = d.getDate();
		document.getElementById("fecContableMS").value = d.getMonth()+1;	//getMonth() entrega el nro de mes de 0 a 11.
		document.getElementById("fecContableAS").value = d.getFullYear();
		document.getElementById("frmSel").submit();
	}
	
	function generateXLS() {		
		var tran = document.getElementById("transporte").value;
		var d = new Date();
		document.getElementById("fileCode").value = document.getElementById("usr").value + "_" + d.getTime();
		document.getElementById("actionLabel").style.visibility = 'visible';
		document.getElementById("actionLabel").innerHTML = "Inicializando... ";
		calculateSegments();
		if (tran == "<% =TIPO_TRANSPORTE_CAMION %>") {
		    document.getElementById("frmSel").action="reporteVisteosCaladaPrintE1.asp";
		} else {
		    document.getElementById("frmSel").action="reporteVisteosCaladaPrintV1.asp";
		}
		document.getElementById("frmSel").target="ifrmXLS";
		generateSegment(currSegment)
	}
			
	
	function seleccionarEntegador(ms) {
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
			
</script>
</head>
<body onLoad="bodyOnLoad()">	

<div id="toolbar"></div>
<br>		
<form id="frmSel" name="frmSel" action="reporteVisteosCalada.asp" method="POST">	
<table id="TAB0" align="center" width="80%" >				
	<tr>
		<td>
			<% Call showErrors() %>
		</td>
	</tr>
</table>	
<table class="reg_Header" id="TAB1" align="center" width="80%" border="0">				
	<tr>
		<td class="reg_Header_nav" align="left" colspan="6">
			<font class="big"><%=GF_Traducir("Reporte de carga de datos Calada")%></big>
		</td>
	</tr>
	<tr>		
		<td class="reg_Header_navdos"><% = GF_TRADUCIR("ID Camión") %></td>
		<td>
			<input type="text" id="idcamion" name="idcamion" value="<% =idcamion %>" onKeyPress="return controlIngreso (this, event, 'N');">
		</td>
		<td class="reg_Header_navdos"><% = GF_TRADUCIR("C.Porte") %>:</td>
		<td width="15%">
			<input type="text" SIZE="2" MAXLENGTH="4" id="nuCartaPorte1" name="nuCartaPorte1" onKeyPress="return controlIngreso (this, event, 'N');" value="<% =nuCartaPorte1 %>">-
			<input type="text" SIZE="8" MAXLENGTH="8" id="nuCartaPorte2" name="nuCartaPorte2" onKeyPress="return controlIngreso (this, event, 'N');"  value="<% =nuCartaPorte2 %>">
		</td>
	</tr>
	<tr id="filaFecha">		
		<td width="13%" align="left" class="reg_Header_navdos"><% = GF_TRADUCIR("Fecha Contable") %>:</td>
		<td>
			<input type="text" size="1" maxLength="2" value="<% =fecContableD%>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecContableD" id="fecContableD"> /
			<input type="text" size="1" maxLength="2" value="<% =fecContableM %>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecContableM" id="fecContableM"> /
			<input type="text" size="2" maxLength="4" value="<% =fecContableA %>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecContableA" id="fecContableA">			
		</td>

		<td class="reg_Header_navdos"><% = GF_TRADUCIR("Hasta") %></td>
		<td>

			<input type="text" size="1" maxLength="2" value="<% =fecContableDH%>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecContableDH" id="fecContableDH"> /
			<input type="text" size="1" maxLength="2" value="<% =fecContableMH %>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecContableMH" id="fecContableMH"> /
			<input type="text" size="2" maxLength="4" value="<% =fecContableAH %>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecContableAH" id="fecContableAH">			
		</td>

	</tr>	
	<tr>
		<td class="reg_Header_navdos" width="10%" align="left"><% = GF_TRADUCIR("Producto") %>:</td>
		<td width="8%">
			<%
			 strSQLPro = "SELECT * FROM PRODUCTOS ORDER BY CDPRODUCTO"
			 call GF_BD_Puertos(pto, rsProductos, "OPEN",strSQLPro)
			 %>
				<select name="cdProducto" value="<%=cdProducto%>">
					<option value="0"> <%=GF_Traducir("Seleccionar...")%></option>
					<%while not rsProductos.eof
						mySelected = ""
						if cint(rsProductos("CDPRODUCTO")) = cint(cdProducto) then mySelected = "SELECTED"%>
						<option value="<%=rsProductos("CDPRODUCTO")%>" <%=mySelected%>> <%=rsProductos("DSPRODUCTO")%></option>
						<%rsProductos.movenext
					 wend%>
			</select>
			
		</td>
		<td class="reg_Header_navdos" width="13%" align="left"><% = GF_TRADUCIR("Corredor") %>:</td>
		<td width="20%">
			<div id="divCorredor"></div>																		
			<input type="hidden" id="cdCorredor" name="cdCorredor" value="<%=cdCorredor%>">
			<input type="hidden" id="dsCorredor" name="dsCorredor" value="<%=dsCorredor%>">
		</td>
	</tr>
	<tr>
		<td class="reg_Header_navdos" width="13%" align="left"><% = GF_TRADUCIR("Vendedor") %>:</td>
		<td width="8%">
			<div id="divVendedor"></div>																		
			<input type="hidden" id="cdVendedor" name="cdVendedor" value="<%=cdVendedor%>">
			<input type="hidden" id="dsVendedor" name="dsVendedor" value="<%=dsVendedor%>">
		</td>
		<td class="reg_Header_navdos" width="13%" align="left"><% = GF_TRADUCIR("Cliente") %>:</td>
		<td width="20%">
			<div id="divCliente"></div>																		
			<input type="hidden" id="cdCliente" name="cdCliente" value="<%=cdCliente%>">
			<input type="hidden" id="dsCliente" name="dsCliente" value="<%=dsCliente%>">
		</td>		
	</tr>	
	<tr>
		<td class="reg_Header_navdos" width="13%" align="left"><% = GF_TRADUCIR("Entregador") %>:</td>
		<td width="20%">
			<div id="divEntregador"></div>																		
			<input type="hidden" id="cdEntregador" name="cdEntregador" value="<%=cdEntregador%>">
			<input type="hidden" id="dsEntregador" name="dsEntregador" value="<%=dsEntregador%>">
		</td>		
		<td class="reg_Header_navdos" width="13%" align="left"><% = GF_TRADUCIR("Estado") %>:</td>
		<td width="20%">
			<%
			 strSQLPro = "SELECT * FROM ESTADOS ORDER BY DSESTADO"
			 call GF_BD_Puertos(pto, rsEstados, "OPEN",strSQLPro)
			 %>
				<select name="estado" value="<%=cdEstado%>">
					<option value="0"> <%=GF_Traducir("Descargados OK")%></option>
					<%while not rsEstados.eof
						mySelected = ""
						if cint(rsEstados("CDESTADO")) = cint(cdEstado) then mySelected = "SELECTED"%>
						<option value="<%=rsEstados("CDESTADO")%>" <%=mySelected%>> <%=rsEstados("DSESTADO")%></option>
						<%rsEstados.movenext
					 wend%>
			</select>
		</td>			
	</tr>
	<tr>
		<td class="reg_Header_navdos" width="13%" align="left"><% = GF_TRADUCIR("Transporte") %>:</td>
		<td width="20%">
			<select  name="transporte" id="transporte">
		        <option value="<% =TIPO_TRANSPORTE_CAMION %>" <% if (cdTransporte = TIPO_TRANSPORTE_CAMION) then response.write "selected"%>> <%=GF_Traducir("CAMIONES")%></option>
		        <option value="<% =TIPO_TRANSPORTE_VAGON %>" <% if (cdTransporte = TIPO_TRANSPORTE_VAGON) then response.write "selected"%>> <%=GF_Traducir("VAGONES")%></option>
            </select>
		</td>		
	</tr>	
</table>
<br>
<div align="center"><div id="actionLabel" class="round_border_all TDSUCCESS" style="width:80%;visibility:hidden;"></div></div>

<input type="hidden" id="accion" name="accion" value="<% =ACCION_SUBMITIR %>">	
<input type="hidden" id="pto" name="pto" value="<% =pto %>">	
<input type="hidden" id="ty" name="ty" value="<% =PDF_REPORT %>">
<input type="hidden" id="fecContableDS" name="fecContableDS">
<input type="hidden" id="fecContableMS" name="fecContableMS">
<input type="hidden" id="fecContableAS" name="fecContableAS">
<input type="hidden" id="usr" name="usr" value="<% =session("Usuario") %>">
<input type="hidden" id="fileCode" name="fileCode" value="">
</form>
<iframe name="ifrmXLS" id="ifrmXLS" width="1px" height="1px" style="visibility:hidden"></iframe>
</body>
</html>
