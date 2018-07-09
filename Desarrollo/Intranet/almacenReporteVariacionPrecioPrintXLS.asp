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
Const TOTAL_COLUMNAS = 9
'************************************************************************************************************
Function writeFilter() 
	Dim auxCategoria, auxArticulo, auxMoneda %>
	<table style="font-size:12; font-weight:bold; font-family:courier;">
		<tr><td colspan="<% =TOTAL_COLUMNAS %>" align="right" style="font-weight:normal; font-size:10"><% =GF_FN2DTE(session("MmtoSistema")) %><br><% =session("usuario") %></td></tr>
		<tr><td colspan="<% =TOTAL_COLUMNAS %>" align="center" style="font-size:24"><% =GF_TRADUCIR("REPORTE DE VARIACIÓN DE PRECIO") %></td></tr>
		<tr><td></td></tr>
	</table>	
	<table style="font-size:12; font-weight:bold; font-family:courier;">		
		<tr><td colspan="6" >Fecha Desde: <% =GF_FN2DTE(g_fechaDesde)	%></td></tr>		
		<tr><td colspan="6" >Fecha Hasta: <% =GF_FN2DTE(g_fechaHasta)	%></td></tr>		
		<tr><td colspan="6" >División:	  <% =g_idDivision &"-"& getDivisionDS(g_idDivision) %></td></tr>
	<%	auxCategoria = "Todas"
		if (g_idCategoria <> 0) then  auxCategoria = getCategoriaDS(g_idCategoria) %>
		<tr><td colspan="6" >Categoria:	 <% = auxCategoria %></td></tr>
	<%	auxArticulo = "Todos"
		if (g_idArticulo <> 0) then  
			call getArticuloFull(g_idArticulo, auxArtDS, auxArtAbrev)
			auxArticulo = g_idArticulo & "-" & auxArtDS & "[" & auxArtAbrev & "]"
		end if	%>
		<tr><td colspan="6" >Artículo:	 <% = auxArticulo%></td></tr>
	<%	auxMoneda = getSimboloMonedaLetras(MONEDA_PESO)
		if (g_cdmoneda = MONEDA_DOLAR) then auxMoneda = getSimboloMonedaLetras(MONEDA_DOLAR) %>			
		<tr><td colspan="6" >Moneda:	 <% = auxMoneda %></td></tr>		
	</table>
<%
End Function
'-------------------------------------------------------------------------------------------------------------
Function drawDetalle()
	sp_parameter = g_idDivision &"||"& g_fechaDesde &"||"& g_fechaHasta &"||"& g_idArticulo &"||"& g_idCategoria &"||"& g_cdmoneda &"||1||0$$totalRegistros"	
	Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLARTICULOSPRECIOS_GET_FULL_BY_PARAMETERS", sp_parameter)
	totalRegistros = sp_ret("totalRegistros")
	if not rs.Eof then 
		while not rs.Eof %>
			<tr style="font-size:10;" class="border">
				<td align="left"><% =rs("IDARTICULO") &"-"& rs("DSARTICULO") %></td>
			<%	auxImporteInicio = "Sin Datos"
				if (not isNull(rs("IMPORTEINICIO"))) then auxImporteInicio = getSimboloMoneda(g_cdmoneda) & " " & GF_EDIT_DECIMALS(rs("IMPORTEINICIO"),2) %>
				<td align="right"><% =auxImporteInicio %></td>		
			<%	auxMmtoInicio = "Sin Datos"
				if (rs("MMTOPRECIOINICIO") <> "") then auxMmtoInicio = GF_FN2DTE(Left(rs("MMTOPRECIOINICIO"),8)) %>
				<td align="center"><% =auxMmtoInicio %></td>
				
			<%	auxImporteFinal = "Sin Datos"
				if (not isNull(rs("IMPORTEFINAL"))) then auxImporteFinal = getSimboloMoneda(g_cdmoneda) & " " & GF_EDIT_DECIMALS(rs("IMPORTEFINAL"),2) %>
				<td align="right"><% =auxImporteFinal %></td>		
			<%	auxMmtoFinal = "Sin Datos"
				if (rs("MMTOPRECIOFINAL") <> "") then auxMmtoFinal = GF_FN2DTE(Left(rs("MMTOPRECIOFINAL"),8)) %>
				<td align="center"><% =auxMmtoFinal %></td>
			
			<%	auxVariacion = "-"
				if ((not isNull(rs("IMPORTEINICIO")))and(not isNull(rs("IMPORTEFINAL")))) then
					if ((Cdbl(rs("IMPORTEINICIO")) > 0)and(Cdbl(rs("IMPORTEFINAL")) > 0)) Then auxVariacion = rs("VARIACION") & "%" 
				end if	%>		
				<td align="right"><% =auxVariacion %></td>
			
			<%	auxImporteUltCom = "Sin Datos"
				if (not isNull(rs("VLUMONEDAULTCOMPRA"))) then auxImporteUltCom = getSimboloMoneda(g_cdmoneda) & " " & GF_EDIT_DECIMALS(rs("VLUMONEDAULTCOMPRA"),2) %>
				<td align="right"><% =auxImporteUltCom %></td>
			<%	auxMmtoUltCom = "Sin Datos"
				if (rs("MMTOULTIMACOMPRA") <> "") then auxMmtoUltCom = GF_FN2DTE(Left(rs("MMTOULTIMACOMPRA"),8)) %>
				<td align="center"><% =auxMmtoUltCom %></td>
				<td align="center">
				<%	if (rs("IDPIC") <> "") then Response.Write rs("IDPIC") %>
				</td>			
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
Dim g_FechaDesde,g_FechaHasta,g_idDivision,flagCall,g_accion,g_origen,g_idCategoria,g_idArticulo,g_cdmoneda
dim oPDF, currentY, contador, color, nroPagina
nroPagina = 1

g_idDivision = GF_PARAMETROS7("idDivision", 0, 6)
g_FechaDesde = GF_PARAMETROS7("fechaDesde", "", 6)
g_FechaHasta = GF_PARAMETROS7("fechaHasta", "", 6)
g_idArticulo = GF_PARAMETROS7("idArticulo", 0, 6)
g_idCategoria = GF_PARAMETROS7("idCategoria", 0, 6)
g_cdmoneda = GF_PARAMETROS7("cdMoneda", "", 6)
fname = "REPORTE_VARIACION_PRECIO_" & session("MmtoSistema")

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
			<td rowspan="2" class="titulos" align="center"><% =GF_TRADUCIR("ARTÍCULO") %></td>
			<td colspan="2" class="titulos" align="center"><% =GF_TRADUCIR("INICIO") %></td>			
			<td colspan="2" class="titulos" align="center"><% =GF_TRADUCIR("FIN") %></td>
			<td rowspan="2" class="titulos" align="center"><% =GF_TRADUCIR("VARIACIÓN") %></td>
			<td colspan="2" class="titulos" align="center"><% =GF_TRADUCIR("ÚLTIMA COMPRA") %></td>
			<td rowspan="2" class="titulos" align="center"><% =GF_TRADUCIR("PIC") %></td>			
		</tr>
		<tr class="border" style="font-size:10;">			
			<td class="titulos" align="center"><% =GF_TRADUCIR("PRECIO") %></td>			
			<td class="titulos" align="center"><% =GF_TRADUCIR("FECHA") %></td>
			<td class="titulos" align="center"><% =GF_TRADUCIR("PRECIO") %></td>			
			<td class="titulos" align="center"><% =GF_TRADUCIR("FECHA") %></td>
			<td class="titulos" align="center"><% =GF_TRADUCIR("PRECIO") %></td>			
			<td class="titulos" align="center"><% =GF_TRADUCIR("FECHA") %></td>
		</tr>
		<% Call drawDetalle() %>
	</table>	
</body>
</html>


	