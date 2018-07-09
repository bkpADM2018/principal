<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosExcel.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<% Response.Buffer = False
Const TOTAL_COLUMNAS = 12
'************************************************************************************************************
Function writeFilter() %>
	<table style="font-size:12; font-weight:bold; font-family:courier;">
		<tr><td colspan="<% =TOTAL_COLUMNAS %>" align="right" style="font-weight:normal; font-size:10"><% =GF_FN2DTE(session("MmtoSistema")) %><br><% =session("usuario") %></td></tr>
		<tr><td colspan="<% =TOTAL_COLUMNAS %>" align="center" style="font-size:24"><% =GF_TRADUCIR("REPORTE DE PIC CON SALDO PENDIENTE") %></td></tr>
		<tr><td></td></tr>
	</table>
	<table style="font-size:12; font-weight:bold; font-family:courier;">
		<tr><td colspan="2" >Divisi�n:	 </td><td colspan="2" align="left"><% =g_idDivision &"-"& getDivisionDS(g_idDivision) %></td></tr>
		<tr><td colspan="2" >Fecha Desde: </td><td colspan="2" align="left"><% =GF_FN2DTE(g_fechaDesde)	%></td></tr>
		<tr><td colspan="2" >Fecha Hasta: </td><td colspan="2" align="left"><% =GF_FN2DTE(g_fechaHasta)	%></td></tr>
	</table>
<%
End Function
'-------------------------------------------------------------------------------------------------------------
Function drawDetalle()
	sp_parameter = g_idDivision &"||"& g_fechaDesde &"||"& g_fechaHasta &"||1||0$$totalRegistros"
    Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLCTZCABECERA_GET_SALDO_PENDIENTE_BY_PARAMETERS", sp_parameter)
	totalRegistros = sp_ret("totalRegistros")
	if not rs.Eof then 
		while not rs.Eof %>
			<tr style="font-size:10;" class="border">
				<td align="center"><% =GF_FN2DTE(Left(rs("MOMENTO"),8)) %></td>
				<td align="center"><% =rs("IDCOTIZACION") %></td>				
				<td align="left"><% =rs("OBRA") %></td>
				<td align="left"><% =getUserDescription(rs("CDUSUARIO")) %></td>
				<td align="left"><% =rs("IDPROVEEDOR") &"-"& rs("DSEMPRESA")%></td>
				<td align="center"><% =getSimboloMonedaLetras(rs("CDMONEDA")) %></td>
				<td align="right"><%  =getSimboloMoneda(MONEDA_PESO) &" "& GF_EDIT_DECIMALS(Cdbl(rs("IMPORTEPESOS")),2)%></td>
				<td align="right"><%  =getSimboloMoneda(MONEDA_PESO) &" "& GF_EDIT_DECIMALS(Cdbl(rs("PESOSFACTURADO")),2) %></td>
				<td align="right"><%  =getSimboloMoneda(MONEDA_PESO) &" "& GF_EDIT_DECIMALS(Cdbl(rs("SALDOPESOS")),2) %></td>
				<td align="right"><%  =getSimboloMoneda(MONEDA_DOLAR) &" "& GF_EDIT_DECIMALS(Cdbl(rs("IMPORTEDOLARES")),2) %></td>
				<td align="right"><%  =getSimboloMoneda(MONEDA_DOLAR) &" "& GF_EDIT_DECIMALS(Cdbl(rs("DOLARFACTURADO")),2) %></td>
				<td align="right"><%  =getSimboloMoneda(MONEDA_DOLAR) &" "& GF_EDIT_DECIMALS(Cdbl(rs("SALDODOLARES")),2) %></td>
			</tr>			
	<%		rs.MoveNext()
		wend		 
	else %>
		<tr>
			<td colspan="<%=TOTAL_COLUMNAS%>" class="border" align="center"><% =GF_TRADUCIR("No se encontraron resultados") %></td>
		</tr>		
<%	end if
End Function
'-------------------------------------------------------------------------------------------------------------
Dim g_idDivision,g_fechaDesde,g_fechaDesdeD,g_fechaDesdeM,g_fechaDesdeA,g_fechaHastaD,g_fechaHastaM,g_fechaHastaA
Dim g_fechaHasta,fname,rs,sp_parameter

g_idDivision = GF_PARAMETROS7("idDivision", 0, 6)
g_fechaDesdeD = GF_PARAMETROS7("fechaDesdeD", "", 6)
g_fechaDesdeM = GF_PARAMETROS7("fechaDesdeM", "", 6)
g_fechaDesdeA = GF_PARAMETROS7("fechaDesdeA", "", 6)
g_fechaDesde = g_fechaDesdeA & g_fechaDesdeM & g_fechaDesdeD
g_fechaHastaD = GF_PARAMETROS7("fechaHastaD", "", 6)
g_fechaHastaM = GF_PARAMETROS7("fechaHastaM", "", 6)
g_fechaHastaA = GF_PARAMETROS7("fechaHastaA", "", 6)
g_fechaHasta = g_fechaHastaA & g_fechaHastaM & g_fechaHastaD

fname = "REPORTE_PIC_SALDO_PENDIENTE_" & session("MmtoSistema")
Call GF_createXLS(fname)
%>
<html>
<head>
	<style type="text/css">
		.border { 
			border-color:#666666; 
			border-style:solid; 
			border-width:thin;
		}

		.titulos {
			background-color:#D8D8D8;
			font-weight:bold;
		}
	</style>
</head>
<body>	
	<table class="border" style="background-color:#FFFACD; font-weight:bold">
		<tr>
			<td><% Call writeFilter() %></td>
		</tr>
	</table>
	<table >
		<tr class="border" style="font-size:10;">
			<td class="titulos" align="center"><% =GF_TRADUCIR("FECHA") %></td>
			<td class="titulos" align="center"><% =GF_TRADUCIR("PIC") %></td>			
			<td class="titulos" align="center"><% =GF_TRADUCIR("OBRA") %></td>
			<td class="titulos" align="center"><% =GF_TRADUCIR("SOLICITANTE") %></td>
			<td class="titulos" align="center"><% =GF_TRADUCIR("PROVEEDOR") %></td>			
			<td class="titulos" align="center"><% =GF_TRADUCIR("MONEDA") %></td>
			<td class="titulos" align="center"><% =GF_TRADUCIR("IMPORTE PESOS") %></td>
			<td class="titulos" align="center"><% =GF_TRADUCIR("FACTURADO PESOS") %></td>
			<td class="titulos" align="center"><% =GF_TRADUCIR("SALDO PESOS") %></td>
			<td class="titulos" align="center"><% =GF_TRADUCIR("IMPORTE DOLARES") %></td>
			<td class="titulos" align="center"><% =GF_TRADUCIR("FACTURADO DOLARES") %></td>
			<td class="titulos" align="center"><% =GF_TRADUCIR("SALDO DOLARES") %></td>			
		</tr>
		<% Call drawDetalle() %>
	</table>	
</body>
</html>


	