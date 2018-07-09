<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<%

Function getRemitos(idPIC, idArticulo)
	dim strSQL, conn, rs, auxwhere
	auxwhere = ""
	if (idArticulo > 0) then auxwhere = " and R.IDARTICULO = " & idArticulo
	strSQL = "Select R.*, C.* from (Select IDREMITO, IDPIC, IDARTICULO, SUM(CANTIDAD) CANTIDAD from TBLREMPIC group by IDREMITO, IDPIC, IDARTICULO) R"
	strSQL = strSQL & " inner join TBLREMCABECERA C on R.IDREMITO = C.IDREMITO"	
	strSQL = strSQL & " where R.IDPIC = " & idPIC & auxwhere
	strSQL = strSQL & " order by R.IDREMITO, R.IDARTICULO"
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set getRemitos = rs
End Function
'---------------------------------------------------------------------------------------------------
'Funcion que obtiene el total pedido para el PIC indicado del articulo indicado.
Function getTotalArticuloPIC(idPIC, idArticulo)
	getTotalArticuloPIC = 0
	strSQL="Select sum(CANTIDAD) CANTIDAD from TBLCTZDETALLE where IDCOTIZACION=" & idPIC & " and IDARTICULO=" & idArticulo
	'Response.Write strSQL
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then getTotalArticuloPIC = CDbl(rs("CANTIDAD"))
End Function
'---------------------------------------------------------------------------------------------------
Function mostrarRemitos(idCotizacionElegida)
%>	
	<table align="center" width="90%" class="reg_Header">
		<tr>
			<td align="center" colspan="7"><b><% =GF_TRADUCIR("Remitos que cumplen el PIC") %></b></td>
		</tr>
		<tr class="reg_Header_nav">
			<td align="center"><% =GF_TRADUCIR(".") %></td>
			<td align="center"><% =GF_TRADUCIR("Remito") %></td>
			<td align="center"><% =GF_TRADUCIR("Nro.Remito") %></td>
			<td align="center"><% =GF_TRADUCIR("Almacen") %></td>
			<td align="center"><% =GF_TRADUCIR("Proveedor") %></td>
			<td align="center"><% =GF_TRADUCIR("Fecha") %></td>
			<td align="center">.</td>
		</tr>	

	<%	
	Set rsREM = getRemitos(idCotizacionElegida, 0)
	reg=0
	if (not rsREM.eof) then	
		while (not rsREM.eof)
			reg=reg+1
			idremito = CDbl(rsREM("IDREMITO"))
			auxremito = idremito
			trClass = "reg_Header_navdos"
			if (CInt(rsREM("ESTADO")) = ESTADO_BAJA) then trClass = "reg_Header_navdos reg_Header_rejected"
		%>							
		<tr class="<% =trClass %>"  onMouseOver="javascript:lightOn(this, <% =rsREM("ESTADO") %>)" onMouseOut="javascript:lightOff(this, <% =rsREM("ESTADO") %>)">
			<td align="center"><img style="cursor:pointer;" onclick="verDetalle(this,'<%="detalleArt_REM_" & reg%>')" title="<%=GF_TRADUCIR("Detalle Articulos")%>" src="images/Mas.gif"></td>
			<td align="center" onclick="REMprint(<%=idremito%>)"><%=rsREM("CDREMITO") & " " & idremito%></td>
			<td align="center" onclick="REMprint(<%=idremito%>)"><% =rsREM("NROREMITO") %></td>
			<td align="left" onclick="REMprint(<%=idremito%>)">
				<%
				Set rsAlmacenes = obtenerListaAlmacenes(rsREM("IDALMACEN"))
				if (not rsAlmacenes.eof) then 
					response.write rsAlmacenes("CDALMACEN") & " - " & rsAlmacenes("DSALMACEN")
				end if
				%>
			</td>
			<td align="left" onclick="REMprint(<%=idremito%>)"><% =rsREM("IDPROVEEDOR") & " - " & getDescripcionProveedor(cDbl(rsREM("IDPROVEEDOR"))) %></td>
			<td align="center" onclick="REMprint(<%=idremito%>)"><% =GF_FN2DTE(rsREM("FECHA")) %></td>
			<td align="center" width="3%"><img style="cursor:pointer;" onclick="REMprint(<%=idremito%>)" title="<%=GF_TRADUCIR("Ver Remito")%>" src="images/almacenes/REM-16x16.png"></td>
		</tr>
		<tr>
			<td colspan="2"></td>
			<td colspan="4" align="left">
				<div id="detalleArt_REM_<%=reg%>" style="display:none;">
					<table width="100%">
						<tr class="reg_Header_nav">
							<td align="center"><% =GF_TRADUCIR("ID Articulo") %></td>
							<td align="center"><% =GF_TRADUCIR("Descripción") %></td>
							<td align="center"><% =GF_TRADUCIR("Cd. Interno") %></td>
							<td align="center"><% =GF_TRADUCIR("Cantidad") %></td>
							<td align="center">.</td>
						</tr>
				<%	While ((not rsREM.eof) and (idremito = auxremito))%>
						<tr class="reg_Header_navdos">
							<td align="center"><b><% =rsREM("IDARTICULO") %></b></td>
							<% Call getArticuloFull(rsREM("IDARTICULO"), descArticulo, pAbrev) %>
							<td align="left"><% =descArticulo %></td>
							<%
								Call executeQueryDb(DBSITE_SQL_INTRA, rsArt, "OPEN", "Select * from TBLARTICULOSDATOS where IDALMACEN=" & rsREM("IDALMACEN") & " and  idArticulo=" & rsREM("IDARTICULO"))
								if (not rsArt.eof) then cdInterno = rsArt("CDINTERNO")
							%>
							<td align="left"><% =cdInterno %></td>
							<td align="right"><% =rsREM("CANTIDAD") & " " & pAbrev & " " %></td>
							<td align="center"><img style="cursor:pointer" title="Ver Cta Cte Articulo" src="images/compras/Search-16x16.png" onClick="mostrarRemitos(<% =idCotizacionElegida %>, <% =rsREM("IDARTICULO") %>, 1)"></td>
						</tr>
				<%		
						rsREM.MoveNext 
						if (not rsREM.eof) then 
							auxremito = CDbl(rsREM("IDREMITO"))
						else
							auxremito = 0
						end if
					Wend 
				%>
					</table>
				</div>
			</td>
			<td></td>
		</tr>
	<%	wend
	end if %>
	<% if (reg = 0) then %>
		<tr><td class="TDNOHAY" colspan="7"><% =GF_TRADUCIR("No se encontraron datos para mostrar") %></td></tr>
	<% end if %>
	</table>
<%
End Function
'---------------------------------------------------------------------------------------------------
Function mostrarCtaCteArticulo(idCotizacionElegida, idArticulo)
	dim cantidad
%>
	<table align="center" width="70%" class="reg_Header">
<%
	Set rsREM = getRemitos(idCotizacionElegida, idArticulo)
	if (not rsREM.eof) then
		Call getArticuloFull(rsREM("IDARTICULO"), descArticulo, pAbrev)
		saldo = getTotalArticuloPIC(idCotizacionElegida, idArticulo)
%>	
		<tr>
			<td colspan="4"><b><% =GF_TRADUCIR("Cumplimiento del Articulo") %>: <% =rsREM("IDARTICULO") & " - " & descArticulo%></b></td>
			<td align="right"><b><% =GF_TRADUCIR("Saldo Inicial: ") & saldo & " " & pAbrev %></b></td>
		</tr>
		<tr class="reg_Header_nav">
			<td align="center"><% =GF_TRADUCIR("Fecha") %></td>
			<td align="center"><% =GF_TRADUCIR("Remito") %></td>
			<td align="center"><% =GF_TRADUCIR("Nro.Remito") %></td>			
			<td align="center"><% =GF_TRADUCIR("Recibido") %></td>			
			<td align="center"><% =GF_TRADUCIR("Saldo") %></td>
			<td align="center" width="3%">.</td>
		</tr>	

	<%			
		reg=0
		while (not rsREM.eof)
			reg=reg+1
			idremito = CDbl(rsREM("IDREMITO"))
			trClass = "reg_Header_navdos"
			if (CInt(rsREM("ESTADO")) = ESTADO_ANULACION) then trClass = "reg_Header_navdos reg_Header_rejected"
			cantidad = CDbl(rsREM("CANTIDAD"))
			'if (rsREM("CDREMITO") = CODIGO_REM_ANULACION) then
				'if (cantidad > 0) then
				'	cantidad = (cantidad * (-1))
				'end if
			'end if
			saldo = saldo - cantidad
		%>							
		<tr class="<% =trClass %>">			
			<td align="center"><% =GF_FN2DTE(rsREM("FECHA")) %></td>
			<td align="center"><%=rsREM("CDREMITO") & " " & rsREM("IDREMITO")%></td>
			<td align="center"><% =rsREM("NROREMITO") %></td>
			<td align="right"><% =cantidad & " " & pAbrev & " " %></td>
			<td align="right"><% =saldo & " " & pAbrev & " " %></td>
			<td align="center"><img style="cursor:pointer;" onclick="REMprint(<%=idremito%>)" title="<%=GF_TRADUCIR("Ver Remito")%>" src="images/almacenes/REM-16x16.png"></td>
		</tr>
<%				
			rsREM.MoveNext 
		wend
	end if %>
	<% if (reg = 0) then %>
		<tr><td class="TDNOHAY" colspan="7"><% =GF_TRADUCIR("No se encontraron datos para mostrar") %></td></tr>
	<% end if %>
	</table>
<%
End Function
'***********************************************************************************
'*******	                     COMIENZO DE LA PAGINA                      ********
'***********************************************************************************
idCotizacionElegida = GF_PARAMETROS7("id", 0, 6)
tipo = GF_PARAMETROS7("ty", 0, 6)
idArticulo = GF_PARAMETROS7("art", 0, 6)

Select case (tipo)
	case 0:
		Call mostrarRemitos(idCotizacionElegida)
	case 1: 
		Call mostrarCtaCteArticulo(idCotizacionElegida, idArticulo)
End Select
%>
