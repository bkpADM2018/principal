<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosExcel.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<%
dim rs, conn, strSQL, oPDF, rsMovimientos, contador, color, nroPagina, currentY, currentAuxY, filename, RPT_idSector
dim RPT_Almacen, RPT_FechaDesde, RPT_FechaHasta, RPT_idObra, RPT_idArticulo, RPT_dsArticulo, RPT_cdSolicitante, RPT_dsSolicitante, RPT_idBudgetArea, RPT_idBudgetDetalle, RPT_idCategoria
dim SEPARATION, MARGIN, PAGE_HEIGHT_SIZE, PAGE_TOP_INIT, RPT_cdCargo, RPT_idCargo, RPT_dsCargo, RPT_verDetalle, RPT_idVale, RPT_auxIdVale

SEPARATION = 10
MARGIN = 0
PAGE_HEIGHT_SIZE = 820
PAGE_TOP_INIT = 82
nroPagina = 1

RPT_Almacen = GF_Parametros7("idAlmacen", "", 6)
RPT_FechaDesde = GF_Parametros7("issuedate", "", 6)
RPT_FechaHasta = GF_Parametros7("closingdate", "", 6)
RPT_idObra = GF_Parametros7("idObra", "", 6)
RPT_idBudgetArea = GF_Parametros7("idBudgetArea", 0, 6)
RPT_idBudgetDetalle = GF_Parametros7("idBudgetDetalle", 0, 6)
RPT_idCategoria = GF_Parametros7("idCategoria", 0, 6)
RPT_idArticulo = GF_Parametros7("idArticulo", 0, 6)
RPT_cdSolicitante = GF_Parametros7("cdSolicitante", "", 6)
RPT_dsSolicitante = getUserDescription(RPT_cdSolicitante)
RPT_cdCargo = GF_Parametros7("cdCargo", "", 6)
if (RPT_cdCargo <> "") then Call GF_MGKS("SG", RPT_cdCargo, RPT_idCargo, RPT_dsCargo)
RPT_verDetalle = GF_Parametros7("verDetalle", "", 6)
if (RPT_verDetalle = "SI") then RPT_verDetalle = true
RPT_idSector = GF_Parametros7("idSector", 0, 6)

filename = "VALES_EMITIDOS_" & RPT_FechaDesde & "-" & RPT_FechaHasta
Call GF_createXLS(filename)
call armadoXLS()
'-----------------------------------------------------------------------------------------------
Function filtrarMovimientos(ByRef myWhere, idAlmacen, fechaDesde, fechaHasta, idObra, idArticulo, cdSolicitante, cdCargo, idCategoria, idSector, idBudgetArea, idBudgetDetalle)
		
	'Filtro
	if (idAlmacen <> 0)			then Call mkWhere(myWhere, "A.IDALMACEN", idAlmacen, "=", 1)	
	if (fechaDesde <> "")		then Call mkWhere(myWhere, "A.FECHA", GF_DTE2FN(fechaDesde), ">=", 3)
	if (fechaHasta <> "")		then Call mkWhere(myWhere, "A.FECHA", GF_DTE2FN(fechaHasta), "<=", 3)	
	if (idObra <> 0)			then Call mkWhere(myWhere, "A.IDOBRA", idObra, "=", 1)
	if (idBudgetArea <> 0)		then Call mkWhere(myWhere, "A.IDBUDGETAREA", idBudgetArea, "=", 1)
	if (idBudgetDetalle <> 0)	then Call mkWhere(myWhere, "A.IDBUDGETDETALLE", idBudgetDetalle, "=", 1)
	if (idArticulo <> 0)		then Call mkWhere(myWhere, "B.IDARTICULO", idArticulo, "=", 1)
	if (cdSolicitante <> "")	then Call mkWhere(myWhere, "A.CDSOLICITANTE", cdSolicitante, "=", 3)
	if (cdCargo <> "")			then Call mkWhere(myWhere, "A.CDUSUARIO", cdCargo, "=", 3)
	if (idCategoria <> 0)		then Call mkWhere(myWhere, "D.IDCATEGORIA", idCategoria, "=", 1)
	if (idSector <> 0)			then Call mkWhere(myWhere, "A.IDSECTOR", idSector, "=", 1)
	filtrarMovimientos = myWhere	
End Function

'-----------------------------------------------------------------------------------------------
Function obtenerListaMovimientos(idAlmacen, fechaDesde, fechaHasta, idObra, idArticulo, cdSolicitante, cdCargo, idCategoria, idSector, idBudgetArea, idBudgetDetalle)
	Dim strSQL, rs, myWhere, conn
	Call filtrarMovimientos(myWhere, idAlmacen, fechaDesde, fechaHasta, idObra, idArticulo, cdSolicitante, cdCargo, idCategoria, idSector, idBudgetArea, idBudgetDetalle)	
	strSQL = "Select A.FECHA, A.IDVALE, A.CDVALE, A.NRVALE, A.CDSOLICITANTE, A.IDOBRA, A.ESTADO, A.IDBUDGETAREA, A.IDBUDGETDETALLE, A.IDSECTOR, B.IDARTICULO, B.CANTIDAD, B.VLUPESOS, B.VLUDOLARES, D.IDCATEGORIA from TBLVALESCABECERA A inner join TBLVALESDETALLE B on A.IDVALE=B.IDVALE" 
	strSQL = strSQL & " inner join TBLARTICULOS C on B.IDARTICULO = C. IDARTICULO inner join TBLARTCATEGORIAS D on C.IDCATEGORIA=D.IDCATEGORIA " & myWhere
	strSQL = strSQL & " group by A.FECHA, A.IDVALE, A.CDVALE, A.NRVALE, A.CDSOLICITANTE, A.IDOBRA, A.ESTADO, A.IDBUDGETAREA, A.IDBUDGETDETALLE, A.IDSECTOR, B.IDARTICULO, B.CANTIDAD, B.VLUPESOS, B.VLUDOLARES, D.IDCATEGORIA order by A.FECHA, A.IDVALE asc"
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set obtenerListaMovimientos = rs
End Function
'------------------------------------------------------------------------------------------
Function armadoXLS()
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

			.art {
				background-color:#F5F6CE;
				font-weight:bold;
			}

			.vales {
				background-color:#CECEF6;
				border-color:#666666; 
				border-style:solid; 
				border-width:thin;
			}

			.titArt {
				background-color:#F2F5A9;
				font-weight:bold;
			}
		</style>
	</head>
	<body>
	<%
	Call armadoCabecera(GF_TRADUCIR("VALES EMITIDOS"), RPT_Almacen, RPT_FechaDesde, RPT_FechaHasta, RPT_idObra, RPT_idBudgetArea, RPT_idBudgetDetalle, RPT_idArticulo, RPT_cdSolicitante, RPT_idCargo, RPT_idCategoria, RPT_verDetalle, RPT_idSector)
	call dibujarTitulosDetalle()
	set rsMovimientos = obtenerListaMovimientos(RPT_Almacen, RPT_FechaDesde, RPT_FechaHasta, RPT_idObra, RPT_idArticulo, RPT_cdSolicitante, RPT_cdCargo, RPT_idCategoria, RPT_idSector, RPT_idBudgetArea, RPT_idBudgetDetalle)
	%><table class="border"><%
	if (rsMovimientos.eof) then
		%>
		<tr><td align="center" colspan="8"><% =GF_TRADUCIR(ucase("No se han encontrado resultados"))%></td></tr>
		<%
	else
		while (not rsMovimientos.eof)
				RPT_idVale = rsMovimientos("IdVale")
				if ((RPT_idVale <> 0) and (RPT_idVale <> RPT_auxIdVale)) then
					call dibujarVale(rsMovimientos("FECHA"), RPT_idVale,rsMovimientos("CdVale"),rsMovimientos("nrVale"), rsMovimientos("cdSolicitante"), rsMovimientos("IDOBRA"), CInt(rsMovimientos("IDBUDGETAREA")), rsMovimientos("IDBUDGETDETALLE"),rsMovimientos("ESTADO"), rsMovimientos("IDSECTOR"))
					if (RPT_verDetalle) then Call dibujarTitularBoxDetalle(MONEDA_PESO)
				end if
				if (RPT_verDetalle) then
					Call dibujarDetalleVale(rsMovimientos("IDARTICULO"), rsMovimientos("CANTIDAD"), rsMovimientos("IDCATEGORIA"), rsMovimientos("VLUPESOS"))
				end if
				RPT_auxIdVale = RPT_idVale
			rsMovimientos.moveNext
		wend
		%>
		<tr><td colspan="5"></td></tr>
		<tr><td colspan="5"></td></tr>
		<tr><td align="center" colspan="8"><% =GF_TRADUCIR("FIN DEL REPORTE")%></td></tr>
		<%
	end if
	%>
	</table>
	</body>
	</html>
	<%
end Function
'------------------------------------------------------------------------------------------
function dibujarTitulosDetalle()
	%>
	<table class="border" style="background-color:#E3F6CE; font-weight:bold">
		<tr>
			<td class="border" align="center"><% =GF_TRADUCIR("FECHA") %></td>
			<td class="border" align="center"><% =GF_TRADUCIR("VALE") %></td>
			<td class="border" align="center"><% =GF_TRADUCIR("TIPO") %></td>
			<td class="border" align="center" colspan="3"><% =GF_TRADUCIR("SOLICITANTE") %></td>
			<td class="border" align="center" colspan="2"><% =GF_TRADUCIR("PTDA. PRES. / SECTOR") %></td>
			<td class="border" align="center"><% =GF_TRADUCIR("ESTADO") %></td>
		</tr>
	</table>
	<%
end function
'------------------------------------------------------------------------------------------
function dibujarVale(fechaVale, idVale, cdVale, nrVale, cdSolicitante, idObra, idBudgetArea, idBudgetDetalle, estado, idsector)
	Dim auxValeRelacionado, dsPartSector

	auxDsSolicitante = getUserDescription(cdSolicitante)
	call loadDatosObra(idObra, RPT_cdObra, RPT_dsObra, 0, "", 0, "", 0, "", "", "", "", "")
	if (RPT_cdObra <> "") then
		dsPartSector = ucase(RPT_cdObra)
		if (idBudgetArea <> 0) then
			set rsBudget = obtenerListaBudgetObra(idObra, idBudgetArea, idBudgetDetalle)
			if not rsBudget.eof then dsPartSector = dsPartSector & " (Detalle Partida: " & idBudgetArea & " - " & idBudgetDetalle & ")"
		end if
	elseif (idsector <> 0) then
		dsPartSector = uCase(getSectorDS(idsector))
	end if
	if (estado <> ESTADO_ACTIVO) then
		auxValeRelacionado = getValeRelacionado(idVale)
		Select case estado 
			Case ESTADO_BAJA
				auxEstado = "ANULADO POR "	& auxValeRelacionado
			Case ESTADO_ANULACION
				auxEstado = "ANULA AL "	& auxValeRelacionado
		End Select	
	end if
	%>
	<tr>
		<td class="vales" align="center"><% =GF_FN2DTE(fechaVale) %></td>
		<td class="vales" align="left"><% =nrVale %></td>
		<td class="vales" align="center"><% =cdVale %></td>
		<td class="vales" align="left" colspan="3"><% =left(ucase(auxDsSolicitante),31) %></td>
		<td class="vales" align="left" colspan="2"><% =ucase(dsPartSector) %></td>
		<td class="vales" align="left"><% =auxEstado %></td>
	</tr>
	<%
end function
'------------------------------------------------------------------------------------------
Function armadoCabecera(p_Titulo, idAlmacen, fechaDesde, fechaHasta, idObra, idBudgetArea, idBudgetDetalle, idArticulo, cdSolicitante, idCargo, idCategoria, verDetalle, idSector)
	dim auxCargo, auxCatDS, auxDetalle, auxSector

	auxAlmacen = GF_Traducir("Todos")
	if (idAlmacen <> 0) then 
		set rsAlmacen = obtenerListaAlmacenes(idAlmacen)
		if not rsAlmacen.eof then auxAlmacen = rsAlmacen("CDALMACEN") & " - " & rsAlmacen("DSALMACEN")
	end if
	auxFechaDesde = GF_Traducir("Todas")
	if (fechaDesde <> "") then auxFechaDesde = fechaDesde
	auxFechaHasta = GF_Traducir("Todas")
	if (fechaHasta <> "") then auxFechaHasta = fechaHasta
	auxObra = GF_Traducir("Todas")
	if (idObra <> 0) then
		call loadDatosObra(idObra, auxObraCD, auxObraDS, 0, "", 0, "", 0, "", "", "", "", "")
		auxObra = auxObraCD & " - " & auxObraDS
		if (idBudgetArea <> 0) then
			set rsBudget = obtenerListaBudgetObra(idObra, idBudgetArea, idBudgetDetalle)
			if not rsBudget.eof then auxObra = auxObra & " (Detalle Partida: " & idBudgetArea & " - " & idBudgetDetalle & ")"
		end if
	end if
	auxArt = GF_Traducir("Todos")
	if (idArticulo <> 0) then 	
		call getArticuloFull(idArticulo, auxArtDS, auxArtAbrev)
		auxArt = idArticulo & " - " & auxArtDS & "[" & auxArtAbrev & "]"
	end if
	auxCatDS = GF_Traducir("Todas")
	if (idCategoria <> 0) then 	auxCatDS = getCategoriaFull(idCategoria)
	auxSolicitante = GF_Traducir("Todos")
	if (cdSolicitante <> "") then auxSolicitante = cdSolicitante & " - " & getUserDescription(cdSolicitante)
	auxCargo = GF_Traducir("Todos")
	if (idCargo <> "")	then auxCargo = RPT_cdCargo & " - " & RPT_dsCargo
	auxDetalle = GF_Traducir("NO")
	if (verDetalle)	then auxDetalle = GF_Traducir("SI")
	auxSector = GF_Traducir("Todos")
	if (idSector <> 0) then auxSector = getSectorDS(idSector)
	%>
	<table class="border" style="background-color:#FFFACD; font-weight:bold">
		<tr>
			<td>
				<table class="border">
					<tr><td colspan="9" align="right" style="font-weight:normal; font-size:10"><% =GF_FN2DTE(session("MmtoSistema")) %><br><% =session("usuario") %></td></tr>
					<tr><td colspan="9" align="center" style="font-size:24"><% =GF_TRADUCIR(p_Titulo) %></td></tr>
				</table>
				<table style="font-size:16; font-weight:bold; font-family:courier">
					<tr><td></td></tr>
					<tr><td><% =GF_TRADUCIR("Almacen......:") %></td><td align="left"><% =ucase(auxAlmacen)		%></td></tr>
					<tr><td><% =GF_TRADUCIR("Fecha Desde..:") %></td><td align="left"><% =ucase(auxFechaDesde)	%></td></tr>
					<tr><td><% =GF_TRADUCIR("Fecha Hasta..:") %></td><td align="left"><% =ucase(auxFechaHasta)	%></td></tr>
					<tr><td><% =GF_TRADUCIR("Ptda. Presup.:") %></td><td align="left"><% =ucase(auxObra)		%></td></tr>
					<tr><td><% =GF_TRADUCIR("Sector.......:") %></td><td align="left"><% =ucase(auxSector)		%></td></tr>
					<tr><td><% =GF_TRADUCIR("Articulo.....:") %></td><td align="left"><% =ucase(auxArt)			%></td></tr>
					<tr><td><% =GF_TRADUCIR("Categoria....:") %></td><td align="left"><% =ucase(auxCatDS)		%></td></tr>
					<tr><td><% =GF_TRADUCIR("Solicitante..:") %></td><td align="left"><% =ucase(auxSolicitante)	%></td></tr>
					<tr><td><% =GF_TRADUCIR("Cargado por..:") %></td><td align="left"><% =ucase(auxCargo)		%></td></tr>
					<tr><td><% =GF_TRADUCIR("Ver Detalle..:") %></td><td align="left"><% =auxDetalle			%></td></tr>
					<tr><td></td></tr>
				</table>
			</td>
		</tr>
	</table>
	<%
end Function
'------------------------------------------------------------------------------------------	
Function dibujarDetalleVale(idArticulo, Cantidad, idCategoria, vlu)
	Dim pAbrev, pDS
	
	if (idArticulo <> 0) then
		Call getArticuloFull(idArticulo, pDS, pAbrev)
		%>
		<tr>
			<td></td>
			<td></td>
			<td class="art" align="center"><% =idCategoria %></td>
			<td class="art" width="600px"><% =GF_TRADUCIR(uCase(getCategoriaDS(idCategoria))) %></td>
			<td class="art"><% =idArticulo %></td>
			<td class="art" width="600px"><% =GF_TRADUCIR(uCase(pDS)) %></td>
			<td class="art" align="right"><% =Cantidad & " " & pAbrev %></td>						
			<td class="art" align="right"><% =GF_EDIT_DECIMALS(CDbl(vlu), 2) %></td>
			<td class="art" align="right"><% =GF_EDIT_DECIMALS(CDbl(vlu)* CDbl(existencia), 2) %></td>
		</tr>
		<%
	end if
end Function
'------------------------------------------------------------------------------------------	
Function dibujarTitularBoxDetalle(cdMoneda)
	%>
	<tr>
		<td></td>
		<td></td>
		<td class="titArt" align="center"><% =GF_TRADUCIR("CATEGORIA") %></td>
		<td class="titArt"><% =GF_TRADUCIR("DESCRIPCIÓN CATEGORIA") %></td>
		<td class="titArt"><% =GF_TRADUCIR("ARTICULO") %></td>
		<td class="titArt"><% =GF_TRADUCIR("DESCRIPCIÓN ARTICULO") %></td>
		<td class="titArt" align="center"><% =GF_TRADUCIR("CANTIDAD")%></td>
		<td class="titArt" align="center"><% =GF_TRADUCIR("VALOR U.") %> (<% =getSimboloMoneda(cdMoneda)%>)</td>
		<td class="titArt" align="center"><% =GF_TRADUCIR("IMPORTE")%></td>
		<td></td>
	</tr>
	<%
end Function
'------------------------------------------------------------------------------------------	
%>