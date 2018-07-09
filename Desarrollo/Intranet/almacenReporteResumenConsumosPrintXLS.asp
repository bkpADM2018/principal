<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosExcel.asp"-->
<!--#include file="Includes/procedimientosLog.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<%
dim RPT_Division, RPT_Month, RPT_Year, RPT_Filtro, RPT_accion
dim rs, conn, strSQL, filename, ultimaLinea

Const LINEA_POST_ENCABEZADO = 10
Const TOTAL_COLUMNAS    = 5

Const TIPO_CATEGORIA = 1
Const TIPO_PART_PRES = 2
Const TIPO_TODOS = 3

'-----------------------------------------------------------------------------------------
Function completarEspacios(p_palabra,p_len, p_agregado)
	Dim rtrn
	rtrn = p_palabra
	for i = len(p_palabra) to p_len
		if (p_agregado = ".") then
			rtrn = rtrn & p_agregado
		else
			rtrn = p_agregado & rtrn
		end if
	next
	completarEspacios = rtrn
End Function
'-----------------------------------------------------------------------------------------
Function writeDatosCategorias()
	dim cdcategoria, descripcion, ccostos, fcolor
	dim strSQL, conn, rsCategorias, contRs, pesos, dolares, TotalPesos, TotalDolares

	TotalPesos = 0
	TotalDolares = 0
	
	Call getSQLCategorias(strSQL)
	Call executeQueryDB(DBSITE_SQL_INTRA, rsCategorias, "OPEN", strSQL)
	if (not rsCategorias.eof) then
		While (not rsCategorias.eof)
			contRs = contRs + 1
			cdcategoria = rsCategorias("CDCATEGORIA")
			descripcion = rsCategorias("DSCATEGORIA")
			pesos = cDbl(rsCategorias("TOTALPESOS"))
			dolares = cDbl(rsCategorias("TOTALDOLARES"))
			%>
			<tr>
				<td class="border" align="center"><% =cdcategoria %></td>
				<td class="border" align="left" colspan="2"><% =uCase(descripcion) %></td>
				<td class="border" align="right"><% ="$ " & GF_EDIT_DECIMALS(round(pesos,0),2) %></td>
				<td class="border" align="right"><% ="U$S " & GF_EDIT_DECIMALS(round(dolares,0),2) %></td>
			</tr>
			<%
			TotalPesos = TotalPesos + pesos
			TotalDolares = TotalDolares + dolares
			rsCategorias.movenext
		Wend
		%>
		<tr class="titulos">
			<td class="border" align="right" colspan="3"><% =GF_TRADUCIR("TOTAL DEL RESUMEN DE CONSUMOS POR CATEGORIAS: ") %></td>
			<td class="border"><% ="$ " & GF_EDIT_DECIMALS(round(TotalPesos,0),2) %></td>
			<td class="border"><% ="U$S " & GF_EDIT_DECIMALS(round(TotalDolares,0),2) %></td>
		</tr>
		
	<%else%>
		<tr><td align="center" colspan="5"><% GF_TRADUCIR("NO SE HAN ENCONTRADO DATOS EN LA BUSQUEDA") %></td></tr>
	<%end if%>
	<tr><td></td></tr>
	<tr><td></td></tr>
	<tr><td></td></tr>
	<%
End Function
'-----------------------------------------------------------------------------------------
Function writeDatosObras()
	dim cuenta, descripcion, ccostos, pesos, dolares
	dim strSQL, conn, rsObras, contRs, TotalPesos, TotalDolares
	dim idObra, actualObra, idArea, actualArea, dsdetalle
	dim TotalObraPesos, TotalObraDolar, cambioObra

	Call getSQLObras(strSQL)
	Call executeQueryDB(DBSITE_SQL_INTRA, rsObras, "OPEN", strSQL)
	if (not rsObras.eof) then
		While (not rsObras.eof)
			cambioObra=false
			contRs = contRs + 1
			idObra = rsObras("IDOBRA")
			idArea = rsObras("IDAREA")
			if ((idObra <> actualObra) or ((idObra = 0) and (contRs = 1))) then
				if (contRs > 1) then
					%>
					<tr class="titulos">
						<td class="border" align="right" colspan="3"><% =GF_TRADUCIR("TOTAL DE LA PART. PRES.: ") %></td>
						<td class="border"><% ="$ " & GF_EDIT_DECIMALS(round(TotalObraPesos,0),2) %></td>
						<td class="border"><% ="U$S " & GF_EDIT_DECIMALS(round(TotalObraDolar,0),2) %></td>
					</tr>
					<%
				end if
				Call WriteCabeceraObra(idObra)
				cambioObra=true
			end if
			if ((idArea <> actualArea) or ((idArea = 0) and (cambioObra))) then
				Call WriteAreaObra(idArea)
			end if
			if (cambioObra) then
				TotalObraPesos = 0
				TotalObraDolar = 0
			end if
			cuenta = rsObras("CDCUENTA")
			descripcion = completarEspacios(rsObras("IDDETALLE"),3," ") & " - "
			dsdetalle = rsObras("DSBUDGET")
			if ((IsNull(dsdetalle)) or (dsdetalle = "")) then dsdetalle = "SIN DETALLE"
			descripcion = descripcion & dsdetalle
			ccostos = " "
			if not (isInversion(idObra)) then	ccostos = rsObras("CCOSTOS")
			pesos = cDbl(rsObras("TOTALPESOS"))
			dolares = cDbl(rsObras("TOTALDOLARES"))
			%>
			<tr>
				<td align="center" class="border"><% =cuenta %></td>
				<td align="center" class="border"><% =ccostos %></td>
				<td align="left" class="border"><% =uCase(descripcion) %></td>
				<td align="right" class="border"><% ="$ " & GF_EDIT_DECIMALS(round(pesos,0),2) %></td>
				<td align="right" class="border"><% ="U$S " & GF_EDIT_DECIMALS(round(dolares,0),2) %></td>
			</tr>
			<%
			actualObra = idObra
			actualArea = idArea
			TotalObraPesos = TotalObraPesos + pesos
			TotalObraDolar = TotalObraDolar + dolares
			TotalPesos = TotalPesos + pesos
			TotalDolares = TotalDolares + dolares
			rsObras.MoveNext
		Wend
		%>
		<tr class="titulos">
			<td class="border" align="right" colspan="3"><% =GF_TRADUCIR("TOTAL DE LA PART. PRES.: ") %></td>
			<td class="border"><% ="$ " & GF_EDIT_DECIMALS(round(TotalObraPesos,0),2) %></td>
			<td class="border"><% ="U$S " & GF_EDIT_DECIMALS(round(TotalObraDolar,0),2) %></td>
		</tr>
		<%
	else
		%>
		<tr><td align="center" colspan="5"><% GF_TRADUCIR("NO SE HAN ENCONTRADO DATOS EN LA BUSQUEDA") %></td></tr>
		<%
	end if
	if (contRs > 0) then
		%>
		<tr><td></td></tr>
		<tr><td></td></tr>
		<tr class="titulos">
			<td class="border" align="right" colspan="3"><% =GF_TRADUCIR("TOTAL DEL RESUMEN DE CONSUMOS POR PART. PRES.: ") %></td>
			<td class="border"><% ="$ " & GF_EDIT_DECIMALS(round(TotalPesos,0),2) %></td>
			<td class="border"><% ="U$S " & GF_EDIT_DECIMALS(round(TotalDolares,0),2) %></td>
		</tr>
		<%
	end if
End Function
'-----------------------------------------------------------------------------------------
Function getSQLCategorias(ByRef strSQL)
	dim almacenes, fecha, fechaDesde, fechaHasta

	if (RPT_Month < 10) then RPT_Month = "0" & right(RPT_Month,1)
	fecha = RPT_Year & RPT_Month
	fechaDesde = cDbl(fecha & "01" & "000000")
	fechaHasta = cDbl(fecha & "31" & "235959")
	almacenes = getAlmacenesPorDivision(RPT_Division)

	strSQL = ""
	strSQL = "SELECT " &_
			 "	       cat.cdcategoria, " &_
			 "	       cat.dscategoria, " &_
			 "	       cat.ccostos, " &_
			 "	       SUM(vd.existencia * vd.vlupesos)   AS totalpesos, " &_
			 "	       SUM(vd.existencia * vd.vludolares) AS totaldolares " &_
			 "	FROM   tblvalescabecera vc " &_
			 "	       INNER JOIN tblvalesdetalle vd ON vc.idvale = vd.idvale " &_
			 "	       INNER JOIN tblarticulos art " &_
			 "	         ON vd.idarticulo = art.idarticulo " &_
			 "	       INNER JOIN tblartcategorias cat " &_
			 "	         ON art.idcategoria = cat.idcategoria " &_
			 "	WHERE      vc.idalmacen IN ( " & almacenes & " ) " &_
			 "	       AND vc.fecha LIKE '" & fecha & "%' " &_
			 "	       AND vc.estado = 1 " &_
			 "	       AND vc.cdvale IN ('" & CODIGO_VS_SALIDA & "','" & CODIGO_VS_AJUSTE_VALE & "','" & CODIGO_VS_AJUSTE_STOCK & "') " &_
			 "	       AND vd.existencia <> 0 " &_
			 "	GROUP  BY cat.cdcategoria, cat.dscategoria, cat.ccostos, cat.cdcuenta " &_
			 "	ORDER  BY cat.cdcategoria"
End Function
'-----------------------------------------------------------------------------------------
Function getSQLObras(ByRef strSQL)
	dim almacenes, fecha, fechaDesde, fechaHastas

	if (RPT_Month < 10) then RPT_Month = "0" & right(RPT_Month,1)
	fecha = RPT_Year & RPT_Month
	fechaDesde = cDbl(fecha & "01" & "000000")
	fechaHasta = cDbl(fecha & "31" & "235959")
	almacenes = getAlmacenesPorDivision(RPT_Division)

	strSQL = ""
	strSQL = "SELECT " &_
			 "	       tg.idobra, " &_
			 "	       tg.dsbudget, " &_
			 "	       tg.idbudgetarea as idarea, " &_
			 "	       tg.idbudgetdetalle as iddetalle, " &_
			 "	       tg.cdcuenta, " &_
			 "	       tg.ccostos, " &_
			 "	       tg.totalpesos, " &_
			 "	       tg.totaldolares " &_
			 "	FROM   (SELECT t1.idobra, " &_
			 "	               t1.idbudgetarea, " &_
			 "	               t1.idbudgetdetalle, " &_
			 "	               t1.dsbudget, " &_
			 "	               t1.cdcuenta, " &_
			 "	               t1.ccostos, " &_
			 "	               SUM(totalpesos)   AS totalpesos, " &_
			 "	               SUM(totaldolares) AS totaldolares " &_
			 "	        FROM   (SELECT vc.idobra, " &_
			 "	                       vc.idbudgetarea, " &_
			 "	                       vc.idbudgetdetalle, " &_
			 "	                       bo.dsbudget, " &_
			 "	                       bo.cdcuenta, " &_
			 "	                       bo.ccostos, " &_
			 "	                       SUM(vd.existencia * vd.vlupesos)   AS totalpesos, " &_
			 "	                       SUM(vd.existencia * vd.vludolares) AS totaldolares " &_
			 "	                FROM   tblvalescabecera vc " &_
			 "	                       INNER JOIN tblvalesdetalle vd " &_
			 "	                         ON vc.idvale = vd.idvale " &_
			 "	                       LEFT JOIN tblbudgetobras bo " &_
			 "	                         ON bo.idobra = vc.idobra " &_
			 "	                            AND bo.idarea = vc.idbudgetarea " &_
			 "	                            AND bo.iddetalle = vc.idbudgetdetalle " &_
			 "	                WHERE      vc.idalmacen IN ( " & almacenes & " ) " &_
			 "	                       AND vc.fecha LIKE '" & fecha & "%' " &_
			 "	                       AND vc.estado = 1 " &_
			 "	                       AND vc.cdvale IN ('" & CODIGO_VS_SALIDA & "','" & CODIGO_VS_AJUSTE_VALE & "','" & CODIGO_VS_AJUSTE_STOCK & "') " &_
			 "	                       AND vd.existencia <> 0 " &_
			 "	                GROUP  BY vc.idobra, " &_
			 "	                          vc.idbudgetarea, " &_
			 "	                          vc.idbudgetdetalle, " &_
			 "	                          bo.dsbudget, " &_
			 "	                          bo.cdcuenta, " &_
			 "	                          bo.ccostos)t1 " &_
			 "	        GROUP  BY t1.idobra, " &_
			 "	                  t1.idbudgetarea, " &_
			 "	                  t1.idbudgetdetalle, " &_
			 "	                  t1.dsbudget, " &_
			 "	                  t1.cdcuenta, " &_
			 "	                  t1.ccostos " &_
			 "	        )tg " &_
			 "	ORDER  BY tg.idobra, idarea, iddetalle"
End Function
'-----------------------------------------------------------------------------------------
Function WriteCabeceraObra(id)
	dim tituloObra
	if (id = 0) then
		tituloObra = "SIN PARTIDA"
	else
		if (isInversion(id)) then
			tituloObra = "OBRA DE INVERSION: "
		else
			tituloObra = "OBRA DE MANTENIMIENTO: "
		end if
	end if
	tituloObra = tituloObra & getDescripcionObra(id)
	%>
	<tr class="titulos">
		<td class="border" align="left" colspan="5"><% =GF_TRADUCIR(tituloObra) %></td>
	</tr>
	<%
End Function
'-----------------------------------------------------------------------------------------
Function WriteAreaObra(idArea)
	dim tituloArea, rs, strSQL, conn
	p_y = p_y + SEPARACION_OBRAS
	tituloArea = "SIN AREA"
	strSQL = "select DSAREA from TBLBUDGETAREAS where IDAREA="&idArea
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then tituloArea = idArea & " " & rs("DSAREA")
	%>
	<tr class="areas">
		<td class="border" align="left" colspan="5"><% =GF_TRADUCIR(tituloArea) %></td>
	</tr>
	<%
End Function
'-----------------------------------------------------------------------------------------
Function dibujarEncabezado(titulo)
	dim division, conn, rsDivision, strSQL
	%>
	<table class="border">
		<tr><td colspan="<% =TOTAL_COLUMNAS %>" align="right" style="font-weight:normal; font-size:10"><% =GF_FN2DTE(session("MmtoSistema")) %><br><% =session("usuario") %></td></tr>
		<tr><td colspan="<% =TOTAL_COLUMNAS %>" align="center" style="font-size:24"><% =GF_TRADUCIR(titulo) %></td></tr>
	</table>
	<%
	if (RPT_Division <> 0) then 
		strSQL = "Select * from TBLDIVISIONES where IDDIVISION=" & RPT_Division
		Call executeQueryDB(DBSITE_SQL_INTRA, rsDivision, "OPEN", strSQL)
		if (not rsDivision.eof) then division = rsDivision("DSDIVISION")
	end if
	%>
	<table style="font-size:16; font-weight:bold; font-family:courier">
		<tr><td></td></tr>
		<tr><td>División.:	</td><td align="left"><% =division				%></td></tr>
		<tr><td>Mes de...:	</td><td align="left"><% =GF_INT2MES(RPT_Month)	%></td></tr>
		<tr><td>Año......:	</td><td align="left"><% =RPT_Year				%></td></tr>
		<tr><td></td></tr>
	</table>
	<%
End Function
'-----------------------------------------------------------------------------------------
'**************************************************************************
'**************************** INICIO PAGINA *******************************
'**************************************************************************

RPT_Division = GF_Parametros7("idDivision", 0, 6)
RPT_Month    = GF_Parametros7("month", "", 6)
RPT_Year     = GF_Parametros7("year", "", 6)
RPT_Filtro     = GF_Parametros7("filtro", 0, 6)
RPT_accion   = GF_Parametros7("accion", "", 6)

if (RPT_Month < 10) then RPT_Month = "0" & right(RPT_Month,1)
filename = "RESUMEN_DE_CONSUMOS" & RPT_Month  & "-" & RPT_Year

if (RPT_accion = ACCION_PROCESAR) then
	'Call GF_createXLS(filename)
else
	Response.Redirect "comprasAccesoDenegado.asp"
end if
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
<body>
	<%if ((RPT_Filtro = TIPO_CATEGORIA) or (RPT_Filtro = TIPO_TODOS)) then%>
		<table class="border" style="background-color:#FFFACD; font-weight:bold">
			<tr><td><% Call dibujarEncabezado("RESUMEN DE CONSUMOS POR CATEGORIAS") %></td></tr>
		</table>
		<table class="border" style="background-color:#E3F6CE; font-weight:bold">
			<tr>
				<td class="border" align="center"><% =GF_TRADUCIR("CATEGORIA") %></td>
				<td class="border" colspan="2"><% =GF_TRADUCIR("DESCRIPCIÓN") %></td>
				<td class="border" align="center"><% =GF_TRADUCIR("TOTAL  $") %></td>
				<td class="border" align="center"><% =GF_TRADUCIR("TOTAL  U$S") %></td>
			</tr>
		</table>
		<table class="border">
			<% Call writeDatosCategorias() %>
		</table>
	<%end if%>
	<%if ((RPT_Filtro = TIPO_PART_PRES) or (RPT_Filtro = TIPO_TODOS)) then%>
		<table class="border" style="background-color:#FFFACD; font-weight:bold">
			<tr><td><% Call dibujarEncabezado("RESUMEN DE CONSUMOS POR PART. PRESUPUESTARIA") %></td></tr>
		</table>
		<table class="border" style="background-color:#E3F6CE; font-weight:bold">
			<tr>
				<td class="border" align="center"><% =GF_TRADUCIR("CUENTA") %></td>
				<td class="border" align="center"><% =GF_TRADUCIR("C.C.") %></td>
				<td class="border" align="left"  ><% =GF_TRADUCIR("DESCRIPCIÓN") %></td>
				<td class="border" align="center"><% =GF_TRADUCIR("TOTAL  $") %></td>
				<td class="border" align="center"><% =GF_TRADUCIR("TOTAL  U$S") %></td>
			</tr>
		</table>
		<table class="border">
			<% Call writeDatosObras()%>
		</table>
	<%end if%>
</body>
</html>