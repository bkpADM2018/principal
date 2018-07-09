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
dim rs, conn, strSQL, filename, ultimaLinea, totalValuacion, totalValuacionTotal

Const LINEA_POST_ENCABEZADO = 5
Const TOTAL_COLUMNAS    = 6


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
Function writeDatosMovimientos()
	dim cdcategoria, descripcion, ccostos, fcolor
	dim strSQL, conn, rsCategorias, cuentaAnterior, pesos, dolares, TotalPesos, TotalDolares

	TotalPesos = 0
	TotalDolares = 0
	
	Call getSQLMovimientos(strSQL)
	Call executeQueryDB(DBSITE_SQL_INTRA, rsCategorias, "OPEN", strSQL)
	if (not rsCategorias.eof) then
		While (not rsCategorias.eof)
			cont = cont + 1	
			if cont = 1 then cuentaAnterior = rsCategorias("CUENTA")
			if cuentaAnterior <> rsCategorias("CUENTA") then
				call lineasDeTotales("TOTAL", totalValuacion)
				totalValuacion = 0
				cuentaAnterior = rsCategorias("CUENTA")
			end if
			idArticulo = rsCategorias("IDARTICULO")
			dsArticulo = rsCategorias("DSARTICULO")
			cuenta = rsCategorias("CUENTA")
			importe = cDbl(rsCategorias("IMPORTE"))
			stockDisponible = cDbl(rsCategorias("STOCKDISPONIBLE"))
			unitario = cDbl(rsCategorias("IMPORTE"))/stockDisponible
			totalValuacion = totalValuacion + importe
			totalValuacionTotal = totalValuacionTotal + cdbl(importe)
			%>
			<tr>
				<td class="border" align="center"><% =idArticulo %></td>
				<td class="border" align="left"><% =dsArticulo %></td>
				<td class="border" align="right"><% =GF_EDIT_DECIMALS(importe*100, 2) %></td>
				<td class="border" align="right"><% =GF_EDIT_DECIMALS(unitario*100, 2) %></td>
				<td class="border" align="right"><% =stockDisponible %></td>
				<td class="border" align="center"><% =cuenta %></td>				
			</tr>
			<%
			rsCategorias.movenext
		Wend
		call lineasDeTotales("TOTAL",totalValuacion)
		%><tr><td></td></tr><%
		call lineasDeTotales("TOTAL GENERAL",totalValuacionTotal)

	else%>
		<tr><td align="center" colspan="5"><% GF_TRADUCIR("NO SE HAN ENCONTRADO DATOS EN LA BUSQUEDA") %></td></tr>
	<%end if%>
	<%
End Function
'-----------------------------------------------------------------------------------------
function lineasDeTotales(pTexto, pImporte)
%>
	<tr style="background-color:#E3F6CE; font-weight:bold">
		<td class="border" align="right" colspan="2"><%=pTexto%></td>
		<td class="border" align="right"><% =GF_EDIT_DECIMALS(pImporte*100, 2) %></td>
		<td class="border" align="right" colspan="3">&nbsp;</td>
	</tr>
<%	
end function
'-----------------------------------------------------------------------------------------
Function getSQLMovimientos(ByRef strSQL)
	strSQL = ""
	strSQL = "SELECT idarticulo, DSARTICULO, (importe/10000) as Importe, stockdisponible, TG.cuenta " & _
			 " from " & _
			"	( " & _
			"	SELECT ART.IDARTICULO, ART.DSARTICULO, CASE WHEN (ART.CDCUENTA='' OR ART.CDCUENTA IS NULL) THEN CAT.CDCUENTA ELSE ART.CDCUENTA END AS CUENTA, STOCKDISPONIBLE, (STOCKDISPONIBLE*VLUPESOS) AS IMPORTE " & _
			"		FROM TBLARTVALUACION VAL " & _
			"			INNER JOIN TBLARTICULOS ART ON VAL.IDARTICULO=ART.IDARTICULO AND (ART.CDCUENTA LIKE '114%' OR  ART.CDCUENTA = '') " & _
			"			INNER JOIN TBLARTCATEGORIAS CAT ON ART.IDCATEGORIA=CAT.IDCATEGORIA AND (CAT.CDCUENTA LIKE '114%' OR CAT.CDCUENTA = '')" & _
			"				WHERE FECHACIERRE=" & pFechaCierre & " AND IDDIVISION=" & pIdDivision & " AND VAL.STOCKDISPONIBLE<>0 " & _
			"	) TG " & _
			"	LEFT JOIN [Database].[dbo].[CGT020A] CGT ON LEFT(TG.CUENTA,9) COLLATE Modern_Spanish_CI_AS  = CONVERT(VARCHAR(9),CGT.CUENTA) " & _			
			"ORDER BY TG.CUENTA " 
getSQLMovimientos = strSQL
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
	if (pIdDivision <> 0) then 
		strSQL = "Select * from TBLDIVISIONES where IDDIVISION=" & pIdDivision
		Call executeQueryDB(DBSITE_SQL_INTRA, rsDivision, "OPEN", strSQL)
		if (not rsDivision.eof) then division = rsDivision("DSDIVISION")
	end if
	%>
	<table style="font-size:16; font-weight:bold; font-family:courier">
		<tr><td></td></tr>
		<tr><td>División.....:	</td><td align="left"><% =division				%></td></tr>
		<tr><td>Fecha Cierre.:	</td><td align="left"><% =pFechaCierre	%></td></tr>
		<tr><td>Cuenta.......:	</td><td align="left"><% =GF_Traducir("Todas")				%></td></tr>
		<tr><td></td></tr>
	</table>
	<%
End Function
'-----------------------------------------------------------------------------------------
'**************************************************************************
'**************************** INICIO PAGINA *******************************
'**************************************************************************
pIdDivision = GF_Parametros7("idDivision", 0, 6)
pFechaCierre = GF_Parametros7("fecCierre", "", 6)
pCuenta = GF_Parametros7("cuentaContable", "", 6) 
filename = "RPT_CIERRES_CONTABLES_" & pIdDivision  & "_" & pFechaCierre & "_TODAS"

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
<body>
		<table class="border" style="background-color:#FFFACD; font-weight:bold">
			<tr><td><% Call dibujarEncabezado("REPORTE DE MOVIMIENTOS DE ARTICULOS EN CUENTA CONTABLES") %></td></tr>
		</table>
		<table class="border" style="background-color:#E3F6CE; font-weight:bold">
			<tr>
				<td class="border" align="center"><% =GF_TRADUCIR("ID ART") %></td>
				<td class="border" align="center"><% =GF_TRADUCIR("DS ART") %></td>	
				<td class="border" align="center"><% =GF_TRADUCIR("TOTAL") %></td>			
				<td class="border" align="center"><% =GF_TRADUCIR("PRECIO UNI") %></td>
				<td class="border" align="center"><% =GF_TRADUCIR("STOCK") %></td>
				<td class="border" align="center"><% =GF_TRADUCIR("CUENTA") %></td>
			</tr>
		</table>
		<table class="border">
			<% Call writeDatosMovimientos() %>
		</table>

</body>
</html>