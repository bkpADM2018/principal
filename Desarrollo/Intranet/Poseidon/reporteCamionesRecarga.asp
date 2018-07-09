<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientostraducir.asp"-->
<!--#include file="../Includes/procedimientosfechas.asp"-->
<!--#include file="../Includes/procedimientosformato.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosSQL.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->


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
dim flagCall,cdProducto,cdVendedor,dsVendedor,cdDestinatario,dsDestinatario,cdCoordinado, dsCoordinado
dim strSQLPro,rsProductos,cdEntregador,dsEntregador, fileCode

Call GP_CONFIGURARMOMENTOS()

ty = GF_PARAMETROS7("ty", "", 6)

pto = GF_PARAMETROS7("pto", "", 6)
Call addParam("pto", pto, params)
Call addParam("idcamion", idcamion, params)

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

cdProducto = GF_PARAMETROS7("cdProducto", 0, 6)
call addParam("cdProducto", cdProducto, params)

cdVendedor = GF_PARAMETROS7("cdVendedor", "", 6)
call addParam("cdVendedor", cdVendedor, params)
dsVendedor = GF_PARAMETROS7("dsVendedor", "", 6)
call addParam("dsVendedor", dsVendedor, params)

cdDestinatario = GF_PARAMETROS7("cdDestinatario", "", 6)
call addParam("cdDestinatario", cdDestinatario, params)
dsDestinatario = GF_PARAMETROS7("dsDestinatario", "", 6)
call addParam("dsDestinatario", dsDestinatario, params)

cdCoordinado = GF_PARAMETROS7("cdCoordinado", "", 6)
call addParam("cdCoordinado", cdCoordinado, params)
dsCoordinado = GF_PARAMETROS7("dsCoordinado", "", 6)
call addParam("dsCoordinado", dsCoordinado, params)
'---------------------------------------------------------

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
<title><%=GF_TRADUCIR("Puertos - Reporte Recarga")%></title>
<link rel="stylesheet" href="../css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="../css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="../css/iwin.css" type="text/css">
<link rel="stylesheet" href="../css/MagicSearch.css" type="text/css">
<link rel="stylesheet" href="../css/calendar-win2k-2.css" type="text/css">
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
<script type="text/javascript" src="../scripts/formato.js"></script>
<script type="text/javascript" src="../scripts/channel.js"></script>
<script type="text/javascript" src="../scripts/controles.js"></script>
<script type="text/javascript" src="../scripts/Toolbar.js"></script>
<script type="text/javascript" src="../scripts/MagicSearchObj.js"></script>
<script defer type="text/javascript" src="../scripts/pngfix.js"></script>

<script type="text/javascript">	
		
	function bodyOnLoad() {			
		tb = new Toolbar('toolbar', 6,'../images/');
		tb.addButton("DocumentoTexto-16x16.png", "Imprimir PDF", "GenerarPDF()");
		tb.addButton("excel.gif", "Imprimir XLS", "GenerarXLS()");
		tb.draw();		
		pngfix();
		var msCoordinado = new MagicSearch("", "divCliente", 25, 4, "puertosStreamElementos.asp?tipo=clientes&pto=<%=pto%>");
			msCoordinado.setToken(";");
			msCoordinado.minChar = 3			
			msCoordinado.onBlur = seleccionarCoordinado;
			msCoordinado.setValue('<% =dsCoordinado %>');
		var msDestinatario = new MagicSearch("", "divDestinatario", 25, 4, "puertosStreamElementos.asp?tipo=destinatarios&pto=<%=pto%>");
			msDestinatario.setToken(";");
			msDestinatario.minChar = 3
			msDestinatario.onBlur = seleccionarDestinatario;
			msDestinatario.setValue('<% =dsDestinatario %>');
		var msVendedor = new MagicSearch("", "divVendedor", 25, 4, "puertosStreamElementos.asp?tipo=vendedores&pto=<%=pto%>");
			msVendedor.setToken(";");
			msVendedor.minChar = 3
			msVendedor.onBlur = seleccionarVendedor;
			msVendedor.setValue('<% =dsVendedor %>');
	<%	if (flagCall) then 
			if (ty = XLS_REPORT) then	%>
				window.open("reporteCamionesRecargaPrintXLS.asp<% =params %>");				
	<%		else	%>
				window.open("reporteCamionesRecargaPrint.asp<% =params %>");
	<%		end if
		end if %>
	}

	function GenerarPDF() {			
		document.getElementById("frmSel").action="reporteCamionesRecarga.asp";
		document.getElementById("frmSel").target="";
		document.getElementById("ty").value = "<% =PDF_REPORT %>";
		document.getElementById("frmSel").submit();
	}
	
	function GenerarXLS() {			
		document.getElementById("frmSel").action="reporteCamionesRecarga.asp";
		document.getElementById("frmSel").target="";
		document.getElementById("ty").value = "<% =XLS_REPORT %>";
		document.getElementById("frmSel").submit();
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
	function seleccionarDestinatario(ms) {
		var desc = ms.getSelectedItem();
		if (desc.indexOf('|') != -1) {
			var arr = desc.split('|');
			document.getElementById("cdDestinatario").value = arr[0];
			document.getElementById("dsDestinatario").value = arr[1];
			ms.setValue(arr[1]);
		} else {
			if (desc == ""){
				document.getElementById("cdDestinatario").value = "";
				document.getElementById("dsDestinatario").value = "";
			}
		}		
	}		
	function seleccionarCoordinado(ms) {
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
			
</script>
</head>
<body onLoad="bodyOnLoad()">	

<div id="toolbar"></div>
<br>		
<form id="frmSel" name="frmSel" action="reporteCamionesRecarga.asp" method="POST">	
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
			<font class="big"><%=GF_Traducir("Reporte de Camiones: Recarga")%></big>
		</td>
	</tr>
	<tr id="filaFecha">		
		<td width="13%" align="left" class="reg_Header_navdos"><% = GF_TRADUCIR("Desde") %>:</td>
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
		<td class="reg_Header_navdos" width="13%" align="left"><% = GF_TRADUCIR("Destinatario") %>:</td>
		<td width="20%">
			<div id="divDestinatario"></div>																		
			<input type="hidden" id="cdDestinatario" name="cdDestinatario" value="<%=cdDestinatario%>">
			<input type="hidden" id="dsDestinatario" name="dsDestinatario" value="<%=dsDestinatario%>">
		</td>
	</tr>
	<tr>
		<td class="reg_Header_navdos" width="13%" align="left"><% = GF_TRADUCIR("Vendedor") %>:</td>
		<td width="8%">
			<div id="divVendedor"></div>																		
			<input type="hidden" id="cdVendedor" name="cdVendedor" value="<%=cdVendedor%>">
			<input type="hidden" id="dsVendedor" name="dsVendedor" value="<%=dsVendedor%>">
		</td>
		<td class="reg_Header_navdos" width="13%" align="left"><% = GF_TRADUCIR("Coordinado") %>:</td>
		<td width="20%">
			<div id="divCliente"></div>																		
			<input type="hidden" id="cdCoordinado" name="cdCoordinado" value="<%=cdCoordinado%>">
			<input type="hidden" id="dsCoordinado" name="dsCoordinado" value="<%=dsCoordinado%>">
		</td>		
	</tr>		
</table>
<br>
<div align="center"><div id="actionLabel" class="round_border_all TDSUCCESS" style="width:80%;visibility:hidden;"></div></div>

<input type="hidden" id="accion" name="accion" value="<% =ACCION_SUBMITIR %>">	
<input type="hidden" id="pto" name="pto" value="<% =pto %>">	
<input type="hidden" id="ty" name="ty" value="<% =PDF_REPORT %>">
<input type="hidden" id="usr" name="usr" value="<% =session("Usuario") %>">
</form>
</body>
</html>
