<!--#include file="../../Includes/procedimientosUnificador.asp"-->
<!--#include file="../../Includes/procedimientosParametros.asp"-->
<!--#include file="reportePosicionTerminalPrintExcel.asp"-->
<%
Const ARCHIVO_PDF = 0
Const ARCHIVO_XLS = 1
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
Dim RPT_TipoArchivo,fecActualD,fecActualM,fecActualA,accion,pto
Dim RPT_accion, fechaActual, lstClientes

pto = GF_PARAMETROS7("pto", "", 6)
call addParam("pto", pto, params)
RPT_TipoArchivo = GF_Parametros7("tipoArchivo", "", 6)
RPT_TipoArchivo = cInt(RPT_TipoArchivo)
accion = GF_PARAMETROS7("accion", "", 6)
'---------------------------------------------------
fecActualD = GF_PARAMETROS7("fecActualD", "", 6)
if (fecActualD = "") then fecActualD=Day(Now()) 
call addParam("fecActualD", fecActualD, params)

fecActualM = GF_PARAMETROS7("fecActualM", "", 6)
if (fecActualM = "") then fecActualM=Month(Now())
Call addParam("fecActualM", fecActualM, params)

fecActualA = GF_PARAMETROS7("fecActualA", "", 6)
if (fecActualA = "") then fecActualA=Year(Now())
Call addParam("fecActualA", fecActualA, params)

if (accion = ACCION_CONTROLAR) then	
	if(GF_CONTROL_FECHA(fecActualD, fecActualM, fecActualA))then
		accion = ACCION_PROCESAR
		call addParam("accion", accion, params)		
		if (RPT_TipoArchivo = ARCHIVO_XLS) then
		    Call GF_STANDARIZAR_FECHA(fecActualD, fecActualM, fecActualA)            
            fechaActual = fecActualA & fecActualM & fecActualD
            fname = "Terminal_" & pto & "_" & fechaActual    
			lstClientes = ""
			if (not isToepfer(session("KCOrganizacion"))) then 
				Call executeQueryDb(pto, rs, "OPEN", "Select CDCLIENTE from Clientes where NUCUIT = '" & session("CuitOrganizacion") & "'")
				lstClientes = rs.GetString(,,,", ")
				lstClientes = Left(lstClientes, Len(lstClientes)-2)
			end if
            Call armarReporteTerminalXLS(pto, fname, fechaActual, lstClientes, XLS_STREAM_MODE)
            response.end			
		end if		
	else
		Call setError(PERIODO_ERRONEO)
	end if		
end if


%>
<html>
<head>

<title>Reporte Posicion Terminal</title>

<meta http-equiv="x-ua-compatible" content="IE=11">

<link rel="stylesheet" href="../../css/main.css" type="text/css">
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
<script type="text/javascript" src="../../scripts/MagicSearchObj.js"></script>
<script defer type="text/javascript" src="../../scripts/pngfix.js"></script>

<script type="text/javascript">	
			
	function GenerarXLS() {			
		location.href = "reportePosicionTerminalPrintExel.asp<%=params%>";				
		submitInfo();
	}	
	
			
	function submitInfo() {
		document.getElementById("frmSel").submit();		
	}	
	
</script>
</head>
<body onLoad="bodyOnLoad()">	

<br>		
<form id="frmSel" name="frmSel" action="reportePosicionTerminal.asp" method="POST">	
<table id="TAB0" align="center" width="60%" >				
	<tr>
		<td>
			<% Call showErrors() %>
		</td>
	</tr>
</table>	
<table class="reg_Header" id="TAB1" align="center" width="60%" border="0">				
	<tr>
		<td class="reg_Header_nav" align="left" colspan="6">
			<font class="big"><%=GF_Traducir("Reporte Posicion de Terminal")%></font>
			<div class="col26"></div>
		</td>
	</tr>
	<tr>		
		<td class="reg_Header_navdos" width="35%"><% = GF_TRADUCIR("Fecha Busqueda") %></td>
		<td>
			<input type="text" size="1" maxLength="2" value="<% =fecActualD%>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecActualD" id="fecActualD"> /
			<input type="text" size="1" maxLength="2" value="<% =fecActualM %>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecActualM" id="fecActualM"> /
			<input type="text" size="2" maxLength="4" value="<% =fecActualA %>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecActualA" id="fecActualA">			
		</td>		
	</tr>
</table>
<div class="col26"></div>
<span class="btnaction">
	<input type="button" value="Generar" onclick="javascript:GenerarXLS()" ></input>
</span>
<input type="hidden" id="accion" name="accion" value="<% =ACCION_CONTROLAR %>">
<input type="hidden" id="pto" name="pto" value="<% =pto %>">	
<input type="hidden" id="tipoArchivo" name="tipoArchivo" value="<% =ARCHIVO_XLS %>">
</form>
</body>
</html>