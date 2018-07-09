<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<%
Const TIPO_CATEGORIA = 1
Const TIPO_PART_PRES = 2
Const TIPO_TODOS = 3

Const ARCHIVO_PDF = 0
Const ARCHIVO_XLS = 1

'-------------------------------------------------------------------------------
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
'-------------------------------------------------------------------------------
Function existeCierre(division, mes, anio)
	dim rtrn, strSQL, conn, rs
	rtrn = false
	strSQL = "select * " &_
					" from TBLCIERRESCABECERA2 " &_
					" where MES = " & mes & " and ANIO = " & anio &_
					"	and IDDIVISION = " & division
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then rtrn = true
	existeCierre = rtrn
End Function
'-------------------------------------------------------------------------------
Function controlFecha(mes, anio)
	dim rtrn, auxmes, auxanio
	auxmes  = false
	auxanio = false
	if (isnumeric(mes)) then
		if ((mes >= 1) and (mes <= 12)) then auxmes = true
	end if
	if (isnumeric(anio)) then
		if ((anio >= 1990) and (anio <= 2020)) then auxanio = true
	end if
	rtrn = ((auxmes) and (auxanio))
	controlFecha = rtrn
End Function
'-------------------------------------------------------------------------------
'****************************************************************
'*******************  COMIENZO DE PAGINA  ***********************
'****************************************************************
dim RPT_Division, RPT_Month, RPT_Year, rsDivisiones, filtroCat, RPT_TipoArchivo
dim fecha, rsalmacenes, ret, rtrn, RPT_accion, RPT_Filtro, filtroObr

RPT_Division = GF_Parametros7("idDivision", 0, 6)
call addParam("idDivision", RPT_Division, params)
RPT_Month    = GF_Parametros7("month", "", 6)
call addParam("month", RPT_Month, params)
RPT_Year     = GF_Parametros7("year", "", 6)
call addParam("year", RPT_Year, params)
RPT_accion   = GF_Parametros7("accion", "", 6)
filtroCat = GF_Parametros7("filtroCat", "", 6)
filtroObr = GF_Parametros7("filtroObr", "", 6)
RPT_Filtro = 0
if (filtroCat = "on") then RPT_Filtro = TIPO_CATEGORIA
if (filtroObr = "on") then
	if (RPT_Filtro = TIPO_CATEGORIA) then
		RPT_Filtro = TIPO_TODOS
	else
		RPT_Filtro = TIPO_PART_PRES
	end if
end if
call addParam("filtro", RPT_Filtro, params)
RPT_TipoArchivo = GF_Parametros7("tipoArchivo", "", 6)
RPT_TipoArchivo = cInt(RPT_TipoArchivo)
if (RPT_accion = ACCION_CONTROLAR) then
	if (controlFecha(RPT_Month, RPT_Year)) then
		if (RPT_Division > 0) then
			if (RPT_Filtro <> 0) then
				if (existeCierre(RPT_Division, RPT_Month, RPT_Year)) then
					RPT_accion = ACCION_PROCESAR
					call addParam("accion", RPT_accion, params)
				else
					setError(CIERRE_NO_EXISTE)
				end if
			else
				setError(CODIGO_VACIO)
			end if
		else
			setError(DIVISION_NO_EXISTE)
		end if
	else
		setError(PERIODO_ERRONEO)
	end if
end if



if (RPT_accion = ACCION_PROCESAR) then
	if (RPT_TipoArchivo = ARCHIVO_PDF) then
		Response.Redirect "almacenReporteResumenConsumosPrint.asp" & params
	elseif (RPT_TipoArchivo = ARCHIVO_XLS) then
		Response.Redirect "almacenReporteResumenConsumosPrintXLS.asp" & params
	end if
end if

%>
<html>
<head>
<title><%=GF_TRADUCIR("Reporte del Resumen de Consumos")%></title>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/iwin.css" type="text/css">
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
<script type="text/javascript" src="scripts/date.js"></script>
<script type="text/javascript" src="scripts/formato.js"></script>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript" src="scripts/script_fechas.js"></script>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/iwin.js"></script>
<script type="text/javascript">	
	var ch = new channel();
	function bodyOnLoad() {
		tb = new Toolbar('toolbar', 6,'images/almacenes/');
		tb.addButton("../DocumentoTexto-16x16.png", "Imprimir PDF", "GenerarPDF()");
		tb.addButton("../excel3.gif", "Imprimir XLS", "GenerarXLS()");
		tb.addButton("Previous-16x16.png", "Volver", "Volver()");
		tb.draw();
	}

	function submitInfo() {
		document.getElementById("actionLabel").style.visibility='visible';
		document.getElementById("frmSel").submit();
	}

	function GenerarPDF() {
		document.getElementById("actionLabel").innerHTML='<% =GF_TRADUCIR("Generando Reporte en PDF") %>...';
		document.getElementById("tipoArchivo").value = '<% =ARCHIVO_PDF %>';
		submitInfo();
	}

	function GenerarXLS() {
		document.getElementById("actionLabel").innerHTML='<% =GF_TRADUCIR("Generando Reporte en Excel") %>...';
		document.getElementById("tipoArchivo").value = '<% =ARCHIVO_XLS %>';
		submitInfo();
	}

	function Volver() {
		location.href = "almacenReportes.asp";
	}

</script>
</head>
<body onLoad="bodyOnLoad()">	
<% call GF_TITULO2("kogge64.gif","Reporte del Resumen de Consumos") %>
<div id="toolbar"></div>
<br>		
<form id="frmSel" name="frmSel" action="almacenReporteResumenConsumos.asp" method="POST">
<table class="reg_Header" align="center" width="60%" border="0">
	<tr><td colspan="3"><% Call showErrors() %></td></tr>
	<tr>
		<td class="reg_Header_nav" align="left" colspan="3">
			<font class="big"><%=GF_Traducir("Resumen de Consumos")%></big>
		</td>
	</tr>
	<tr>
		<!--Division-->	
		<td class="reg_Header_navdos" align="left"><%=GF_TRADUCIR("Division")%></td>
		<td align="left">
			<select id="idDivision" name="idDivision">
				<% Set rsalmacenes = obtenerListaAlmacenesUsuario() %>
				<% While (not rsalmacenes.eof) %>
					<% ret = ret & getDivisionAlmacen(rsalmacenes("IDALMACEN")) & "," %>
					<% rsalmacenes.MoveNext() %>
				<% Wend %>
				<% ret = left(ret, len(ret)-1) %>
				<% Call executeQueryDB(DBSITE_SQL_INTRA, rsDivisiones, "OPEN", "Select * from TBLDIVISIONES Where IDDIVISION in (" & ret & ")") %>
				<% While (not rsDivisiones.eof) %>
						<option value="<% =rsDivisiones("IDDIVISION") %>" <% if (rsDivisiones("IDDIVISION") = RPT_idDivision) then response.write "selected='true'" %>><% =GF_TRADUCIR(rsDivisiones("DSDIVISION")) %></option>
						<% rsDivisiones.MoveNext() %>
				<% wend %>
			</select>
		</td>
        <td>
			<input id="filtroCat" name="filtroCat" type="checkbox" value="on" checked>
			<%= GF_TRADUCIR("Por Categorias") %>
        </td>
	</tr>
	<tr>
		<!--Fecha-->
		<td class="reg_Header_navdos"><% =GF_TRADUCIR("Fecha - Mes/Año") %></td>
		<td align="left">
			<input type="text" id="month" name="month" value="<% =RPT_Month %>" maxlength="2" size="3" align="right" onKeyPress="return controlDatos(this, event, 'N');">
			/
			<input type="text" id="year" name="year" value="<% =RPT_Year %>" maxlength="4" size="5" align="right" onKeyPress="return controlDatos(this, event, 'N');">
		</td>
        <td>
			<input id="filtroObr" name="filtroObr" type="checkbox" value="on" checked>
			<%= GF_TRADUCIR("Por Part. Pres.") %>
        </td>
	</tr>
</table>
<div align="center"><div id="actionLabel" class="round_border_bottom TDSUCCESS" style="width:60%;visibility:hidden;"><% =GF_TRADUCIR("Generando Reporte") %>...</div></div><br>
<input type="hidden" id="accion" name="accion" value="<% =ACCION_CONTROLAR %>">
<input type="hidden" id="tipoArchivo" name="tipoArchivo" value="<% =ARCHIVO_PDF %>">
</form>
</body>
</html>