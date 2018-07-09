<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosPuertos.asp"-->
<!--#include file="../Includes/procedimientosSQL.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->
<!--#include file="../Includes/procedimientosExcel.asp"-->
<!--#include file="../Includes/procedimientostraducir.asp"-->
<!--#include file="../Includes/procedimientosfechas.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="reporteCamionesRecargaCommon.asp"-->
<% Response.Buffer = False %>
<%

Function imprimirFiltros(pto,fechaD,fechaH,pcdProducto, pDsProducto,pcdVendedor,pcdDestinatario,pcdCoordinado)	
%>
	<table style="font-size:16; font-weight:bold; font-family:courier">
	<tr>
		<td colspan="4">Puerto......: <% =pto %></td>
		<td></td>
<%
	auxDestinatario = "Todos"
	if(pcdDestinatario > 0)then
		auxDestinatario = Trim(pcdDestinatario)&" - "&Trim(getDsComprador(pcdDestinatario))	
	end if	
%>	
		<td colspan="4">Destinatario.: <% =auxDestinatario %></td>
	</tr>
<%
	myFormatFecha = Replace(fechaD,"-", "") 	
	myFormatFecha = GF_FN2DTE(myFormatFecha)
%>
	<tr>
		<td colspan="4">Fecha Desde.: <% =myFormatFecha %></td>
		<td></td>
<%
	auxVendedor = "Todos"
	if(pcdVendedor > 0)then 
		auxVendedor = Trim(pcdVendedor)&" - "&Trim(getDsVendedor(pcdVendedor))			
	end if
%>
		<td colspan="4">Vendedor.....: <% =auxVendedor %></td>
	</tr>
<%
	myFormatFecha = Replace(fechaH,"-", "") 
	myFormatFecha = GF_FN2DTE(myFormatFecha)
%>
	<tr>
		<td colspan="4">Fecha Hasta.: <% =myFormatFecha %></td>
		<td></td>
<%
	auxCoordinado = "Todos"
	if(pcdCoordinado > 0)then 
		auxCoordinado = Trim(pcdCoordinado)&" - "&Trim(getDsCliente(pcdCoordinado))			
	end if
%>

		<td colspan="4">Coordinado...: <% =auxCoordinado %></td>
	</tr>
<%
	if(pidcamion > 0)then 
		auxCamion = GF_nDigits(pidcamion,10)		
	else
		if (nuCartaPorte <> "") then
			auxCamion = GF_EDIT_CTAPTE(GF_nChars(nuCartaPorte, 16, "0", CHR_AFT))
		else
			auxCamion = "Todos" 
		end if
	end if
%>
	<tr>
		<td colspan="4">Camion......: <% =auxCamion %></td>
		<td></td>
<%
	auxEntregador = "Todos"
	if(pcdEntregador > 0)then 
		auxEntregador = Trim(pcdEntregador)&" - "&Trim(getDsEntregador(pcdEntregador))	
	end if
%>	
		<td colspan="4">Entregador...: <% =auxEntregador %></td>
	</tr>
<%
	auxProducto = "Todos"
	if(pcdProducto > 0)then auxProducto = Trim(pcdProducto)&" - "&Trim(pDsProducto)	
%>
	<tr>
		<td colspan="4">Producto....: <% =auxProducto %></td>
	</tr>
	</table>
<%
End Function
'-----------------------------------------------------------------------------------------
Function imprimirTitulos()
%>	
		<tr style="background-color:#E3F6CE; font-weight:bold">
			<td class="border" align="center"><%=GF_TRADUCIR("REMITO")%></td>
			<td class="border" align="center"><%=GF_TRADUCIR("FECHA")%></td>
			<td class="border" align="center"><%=GF_TRADUCIR("TURNO")%></td>
			<td class="border" align="center"><%=GF_TRADUCIR("ID CAMION ")%></td>
			<td class="border" align="center"><%=GF_TRADUCIR("CTA. PORTE")%></td>			
			<td class="border" align="center"><%=GF_TRADUCIR("COORDINADO")%></td>
			<td class="border" align="center"><%=GF_TRADUCIR("DESTINATARIO")%></td>
			<td class="border" align="center"><%=GF_TRADUCIR("VENDEDOR")%></td>
			<td class="border" align="center"><%=GF_TRADUCIR("PATENTE")%></td>
			<td class="border" align="center"><%=GF_TRADUCIR("BRUTO")%></td>
			<td class="border" align="center"><%=GF_TRADUCIR("TARA")%></td>
			<td class="border" align="center"><%=GF_TRADUCIR("NETO S/MERMA")%></td>
		</tr>	
<%
End Function
'-----------------------------------------------------------------------------------------
Function imprimirDatos()
	Dim cdProducto_old, flagInicio, totalBruto, totalTara, totalMerma, auxSMerma
	flagInicio = true	
	if(Not rsRecarga.Eof)then
		while Not rsRecarga.EoF
			auxSMerma = Cdbl(rsRecarga("Bruto")) - Cdbl(rsRecarga("Tara"))
			if(Cdbl(rsRecarga("cdproducto")) <> cdProducto_old)then
				cdProducto_old = CDbl(rsRecarga("CDPRODUCTO"))
				if(not flagInicio)then
				%>
					<tr style="background-color:#D8D8D8; font-weight:bold">
						<td class="border" colspan=9 align="center"><%=GF_TRADUCIR("TOTAL")%></td>						
						<td class="border" align="right"><%= GF_EDIT_DECIMALS(cdbl(totalBruto)*100,2) %></td>
						<td class="border" align="right"><%= GF_EDIT_DECIMALS(cdbl(totalTara)*100,2) %></td>
						<td class="border" align="right"><%= GF_EDIT_DECIMALS(cdbl(totalMerma)*100,2) %></td>
					</tr>
				<%
				end if
				totalMerma = 0
				totalTara  = 0
				totalBruto = 0				
				flagInicio = false				
				%>				
				<tr></tr>
				<tr style="font-weight:bold">					
					<td colspan=12><b><%= GF_TRADUCIR("PRODUCTO: ") & rsRecarga("CDPRODUCTO") & " - " & rsRecarga("DSPRODUCTO")%></b></td>					
				</tr>					
				<%				
				Call imprimirTitulos()
			end if
			%>	
				<tr style="font-weight:bold">
					<td class="border" align="right"><%=rsRecarga("Remito") %></td>
					<td class="border" align="center"><%=GF_FN2DTE(rsRecarga("Fecha"))%></td>
					<td class="border" align="right"><%=rsRecarga("Turno")  %></td>
					<td class="border" align="right"><% Response.Write("=""" & rsRecarga("IdCamion") & """") %></td>
					<td class="border" align="left"><%=GF_EDIT_CTAPTE(GF_nChars(rsRecarga("CP"), 16, "0", CHR_AFT)) %></td>					
					<td class="border" align="left"><%=rsRecarga("Coordinado") %></td>
					<td class="border" align="left"><%=rsRecarga("Destinatario") %></td>					
					<td class="border" align="left"><%=rsRecarga("Vendedor") %></td>					
					<td class="border" align="left"><%=GF_EDIT_PATENTE(rsRecarga("Chapa")) %></td>
					<td class="border" align="right"><%=GF_EDIT_DECIMALS(cdbl(rsRecarga("Bruto"))*100,2)%></td>
					<td class="border" align="right"><%=GF_EDIT_DECIMALS(cdbl(rsRecarga("Tara"))*100,2)%></td>
					<td class="border" align="right"><%=GF_EDIT_DECIMALS(cdbl(auxSMerma)*100,2)%></td>
				</tr>	
			<%			
			totalMerma = totalMerma + auxSMerma
			totalTara  = totalTara + Cdbl(rsRecarga("Tara"))
			totalBruto = totalBruto + Cdbl(rsRecarga("Bruto"))
			rsRecarga.MoveNext()
		wend
		%>
			<tr style="background-color:#D8D8D8; font-weight:bold">
				<td class="border" colspan=9 align="center"><%=GF_TRADUCIR("TOTAL")%></td>
				<td class="border" align="right"><%= GF_EDIT_DECIMALS(cdbl(totalBruto)*100,2) %></td>
				<td class="border" align="right"><%= GF_EDIT_DECIMALS(cdbl(totalTara)*100,2)  %></td>
				<td class="border" align="right"><%= GF_EDIT_DECIMALS(cdbl(totalMerma)*100,2) %></td>
			</tr>
		<%
	else
		%>
		<tr style="background-color:#D8D8D8; font-weight:bold">
			<td class="border" align="center" colspan="12"><%=GF_TRADUCIR("No se encontraron resultados") %></td>
		</tr>
		<%
	end if
End Function
'******************************************************************************************************
'**************************************** COMIENZO DE PAGINA ******************************************
'******************************************************************************************************
Randomize()
filename = "RECARGA_" & g_Puerto
if (dsProducto <> "") then filename = filename & "_" & dsProducto
filename = filename & "_" & Replace(g_fechaDesde,"-", "_") & "_al_" & Replace(g_fechaHasta,"-", "_") & "_" & Int(100 * Rnd()) & ".xls"
Call GF_createXLS(filename)
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

		.areas {
			background-color:#CECEF6;
			font-weight:bold;
		}
	</style>	
</head>
<body onLoad="bodyOnLoad()">	
	<table class="border" style="background-color:#FFFACD; font-weight:bold">		
		<tr><td colspan=12 align="right" style="font-weight:normal; font-size:10"><% =GF_FN2DTE(session("MmtoSistema")) %><br><% =session("usuario") %></td></tr>
		<tr><td colspan=12 align="center" style="font-size:24"><% =GF_TRADUCIR("REPORTE DE CAMIONES: RECARGA") %></td></tr>		
	</table>
	
<%		
	Call imprimirFiltros(g_Puerto,g_fechaDesde,g_fechaHasta,g_Producto, dsProducto,g_Vendedor,g_Destinatario,g_Coordinado)
%>
	<table class="border">
<%	
	Call imprimirDatos()	
%>		
	</table>
</body>
</html>
