<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosPDF.asp"-->
<!--#include file="Includes/procedimientossql.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosformato.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientoslog.asp"-->
<%
Const SEPARATION = 10
Const MARGIN = 0
Const PAGE_HEIGHT_SIZE = 800
Const PAGE_TOP_INIT = 82
Const PAGE_SPACE_DETAIL = 20
Const PAGE_SPACE_SALDOS = 16
Const HHMMSS_FECHA_CIERRE = "235959"
Const CODIGO_REM_REMITO = "REM"
Const CODIGO_REM_ANULACION = "XEM"

'constantes de posicion
Const PAGE_INIT			 = 10
Const PAGE_INIT_CDCUENTA = 10 
Const PAGE_INIT_DSCUENTA = 80 
Const PAGE_INIT_VLUPESOS = 280 
Const PAGE_INIT_FILTROS	 = 100

'constantes de ancho
Const PAGE_WIDTH			= 570
Const PAGE_WIDTH_CDCUENTA	= 70 
Const PAGE_WIDTH_DSCUENTA   = 200 
Const PAGE_WIDTH_VLUPESOS   = 80

Const PAGE_WIDTH_CABECERA	= 580
Const PAGE_WIDTH_TIT_FILTROS= 50
Const PAGE_WIDTH_FILTROS	= 200
'------------------------------------------------------------------------------------------
Function armadoPDF()
	currentY = armadoCabecera(GF_TRADUCIR("Reporte Saldo Cuentas Contables"), RPT_IdDivision, RPT_FechaCierre)
	Call armadoDetalleMovimientos(RPT_IdDivision, RPT_FechaCierre)
end Function
'------------------------------------------------------------------------------------------
Function armadoCabecera(p_Titulo, idDivision, fechaCierre)
	dim yInicio, ySeparation
	yInicio = PAGE_TOP_INIT
	ySeparation = 0

	call dibujarTitulo(p_Titulo)
	call GF_setFont(oPDF,"COURIER",8,0)

	Call GF_writeTextAlign(oPDF, PAGE_INIT, yInicio, GF_TRADUCIR("Division.........:"), PAGE_WIDTH_TIT_FILTROS, PDF_ALIGN_LEFT)
	dsDivision = getDivisionDS(idDivision)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_FILTROS, yInicio, ucase(dsDivision), PAGE_WIDTH_FILTROS, PDF_ALIGN_LEFT)
	ySeparation = ySeparation + SEPARATION
	
	Call GF_writeTextAlign(oPDF, PAGE_INIT, yInicio + ySeparation, GF_TRADUCIR("Fecha Cierre.....:"), PAGE_WIDTH_TIT_FILTROS, PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_FILTROS, yInicio + ySeparation, GF_FN2DTE(fechaCierre), PAGE_WIDTH_FILTROS, PDF_ALIGN_LEFT)
	ySeparation = ySeparation + SEPARATION

	yFinal = yInicio + (ySeparation + SEPARATION)

	armadoCabecera = yFinal
end Function
'------------------------------------------------------------------------------------------
Function dibujarTitulo(pTitulo)
	Call GF_squareBox(oPDF, 2, 2, PAGE_WIDTH_CABECERA + 10, 833, 0, "", "#0B3B0B", 2, PDF_SQUARE_ROUND)
	Call GF_writeImage(oPDF, Server.MapPath("images\kogge64.gif"), PAGE_INIT, 10, 48, 48, 0)
	call GF_setFont(oPDF,"ARIAL",16,0)
	Call GF_writeTextAlign(oPDF,2, 25, GF_TRADUCIR(pTitulo), PAGE_WIDTH_CABECERA + 10, PDF_ALIGN_CENTER)
	Call GF_horizontalLine(oPDF,2,65,590)
	call GF_setFont(oPDF,"ARIAL",8,0)
	Call GF_writeTextAlign(oPDF, 10 , 840, "Pagina  " & nroPagina, PAGE_WIDTH_CABECERA , PDF_ALIGN_RIGHT)
	Call GF_setFont(oPDF,"COURIER",8,0)
	GP_CONFIGURARMOMENTOS
	Call GF_writeTextAlign(oPDF,5,5,GF_FN2DTE(session("MmtoSistema")), PAGE_WIDTH_CABECERA , PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,5,5+SEPARATION,session("Usuario"), PAGE_WIDTH_CABECERA , PDF_ALIGN_RIGHT)
end Function
'------------------------------------------------------------------------------------------
Function nuevaPagina()
	Call GF_newPage(oPDF)
	currentY = PAGE_TOP_INIT
	nroPagina = nroPagina + 1
	call dibujarTitulo(GF_TRADUCIR("Reporte de Cuenta Corriente del Articulo"))
	call dibujarTitulosDetalle(currentY)
	currentY = PAGE_TOP_INIT + PAGE_SPACE_DETAIL
end Function
'------------------------------------------------------------------------------------------
Function dibujarSeparadores(yInicio, yTotal)
	Call GF_verticalLine(oPDF, PAGE_INIT_CDCUENTA, yInicio, yTotal - yInicio)
	Call GF_verticalLine(oPDF, PAGE_INIT_DSCUENTA, yInicio, yTotal - yInicio)
	Call GF_verticalLine(oPDF, PAGE_INIT_VLUPESOS, yInicio, yTotal - yInicio)
end Function
'------------------------------------------------------------------------------------------
Function dibujarTitulosDetalle(pY)
	Call dibujarTitulosDetalleBox(pY)
	call GF_setFont(oPDF,"ARIAL",8,8)
	call GF_setFontColor("FFFFFF")
	Call GF_writeTextAlign(oPDF, PAGE_INIT_CDCUENTA, pY+5, GF_TRADUCIR("CUENTA"), PAGE_WIDTH_CDCUENTA, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_DSCUENTA, pY+5, GF_TRADUCIR("DESCRIPCION"), PAGE_WIDTH_DSCUENTA, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_VLUPESOS, pY+5, GF_TRADUCIR("SALDO"), PAGE_WIDTH_VLUPESOS, PDF_ALIGN_CENTER)
	call GF_setFontColor("000000")
end Function			
'------------------------------------------------------------------------------------------
Function dibujarTitulosDetalleBox(pY)
	Call GF_squareBox(oPDF, PAGE_INIT, pY, PAGE_WIDTH, 18, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_verticalLine(oPDF, PAGE_INIT_CDCUENTA+PAGE_WIDTH_CDCUENTA, pY, 18)
	Call GF_verticalLine(oPDF, PAGE_INIT_DSCUENTA+PAGE_WIDTH_DSCUENTA, pY, 18)
	Call GF_verticalLine(oPDF, PAGE_INIT_VLUPESOS+PAGE_WIDTH_VLUPESOS, pY, 18)
end Function
'------------------------------------------------------------------------------------------
Function dibujarSaldosFinales(p_currentY, p_totalSaldoAcumulado)
	Call GF_squareBox(oPDF, PAGE_INIT, p_currentY + 3, PAGE_WIDTH, SEPARATION + 3, 0, "#CECEF6", "", 0, PDF_SQUARE_NORMAL)
	Call GF_horizontalLine(oPDF,PAGE_INIT,p_currentY + 2,PAGE_WIDTH)
	call GF_setFont(oPDF,"COURIER",8,8)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_DSCUENTA, p_currentY + 4, GF_TRADUCIR("TOTAL INVENTARIO"), PAGE_WIDTH_DSCUENTA, PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_VLUPESOS, p_currentY + 4, getSimboloMoneda(RPT_Moneda) & " " & GF_EDIT_DECIMALS(cdbl(p_totalSaldoAcumulado)/100, 2), PAGE_WIDTH_VLUPESOS, PDF_ALIGN_RIGHT)
	Call GF_horizontalLine(oPDF,PAGE_INIT, p_currentY + SEPARATION + 5, PAGE_WIDTH)
end Function
'------------------------------------------------------------------------------------------
Function dibujarMovimiento(pY, cdCuenta, dsCuenta, vlPesos)
	dim auxDsCta
	call GF_setFont(oPDF,"COURIER",8,0)
	contador = contador + 1
	if contador mod 2 then
		color = "#FFFFFF"
	else
		color = "#E0E0F8"
	end if
	Call GF_squareBox(oPDF, PAGE_INIT, pY, PAGE_WIDTH, 10, 0, color, color, 1, PDF_SQUARE_NORMAL)
	auxDsCta = dsCuenta
	if isnull(auxDsCta) then auxDsCta = "Falta descripción de la cuenta " & TRIM(cdCuenta)
	Call GF_writeTextAlign(oPDF,PAGE_INIT_CDCUENTA, pY, TRIM(cdCuenta), PAGE_WIDTH_CDCUENTA, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,PAGE_INIT_DSCUENTA , pY, auxDsCta, PAGE_WIDTH_DSCUENTA, PDF_ALIGN_LEFT)			
	Call GF_writeTextAlign(oPDF,PAGE_INIT_VLUPESOS-2, pY, getSimboloMoneda(RPT_Moneda) & " " & GF_EDIT_DECIMALS(cDbl(cdbl(vlPesos)/100),2), PAGE_WIDTH_VLUPESOS, PDF_ALIGN_RIGHT)	
end Function
'------------------------------------------------------------------------------------------
'realiza los controles y arma el listado de movimientos
Function armadoDetalleMovimientos(idDivision, fechaCierre)
	dim rsMovimientos, totalReg, i, isInicio, aux_Y, totalAcumulado
	
	'Tomo los movimientos a partir del siguiente al ultimo cierre.
	Set rsMovimientos = obtenerSaldoCuentas(fechaCierre, idDivision)	
	call dibujarTitulosDetalle(currentY)
	currentY = currentY + PAGE_SPACE_DETAIL
	aux_Y = currentY
	isInicio = true
	flagNuevaPagina = false
	
	while not rsMovimientos.eof 
			if (flagNuevaPagina) then
				Call dibujarSeparadores(currentY, aux_Y)
				Call GF_horizontalLine(oPDF,PAGE_INIT, currentY, PAGE_WIDTH)
				Call nuevaPagina()
				Call GF_horizontalLine(oPDF,PAGE_INIT, currentY - 1, PAGE_WIDTH)
				aux_Y = currentY
				flagNuevaPagina = false
			end if
			call dibujarMovimiento(currentY, rsMovimientos("CDCUENTA"), rsMovimientos("DSCUENTA"), rsMovimientos("TOTAL_CUENTA"))
			currentY = currentY + SEPARATION
			if (currentY > PAGE_HEIGHT_SIZE) then flagNuevaPagina = true
			totalAcumulado = totalAcumulado + CDBL(rsMovimientos("TOTAL_CUENTA"))
		rsMovimientos.MoveNext()
	wend
	Call dibujarSaldosFinales(currentY, totalAcumulado)
end Function
'------------------------------------------------------------------------------------------
Function obtenerSaldoCuentas(fechaCierre, idDivision)
	Dim strSQL, rs, conn
	strSQL= "SELECT TG.CUENTA AS CDCUENTA, CGT.NOMCUE AS DSCUENTA, SUM(IMPORTE) AS TOTAL_CUENTA FROM " & _
			"	( " & _
			"	SELECT ART.IDARTICULO, ART.DSARTICULO, CASE WHEN (ART.CDCUENTA='' OR ART.CDCUENTA IS NULL) THEN CAT.CDCUENTA ELSE ART.CDCUENTA END AS CUENTA, (STOCKDISPONIBLE*VLUPESOS) AS IMPORTE " & _
			"		FROM TBLARTVALUACION VAL " & _ 
			"			INNER JOIN TBLARTICULOS ART ON VAL.IDARTICULO=ART.IDARTICULO AND (ART.CDCUENTA LIKE '1141%' OR  ART.CDCUENTA = '') " & _
			"			INNER JOIN TBLARTCATEGORIAS CAT ON ART.IDCATEGORIA=CAT.IDCATEGORIA AND (CAT.CDCUENTA LIKE '114%' OR CAT.CDCUENTA = '')" & _
			"				WHERE FECHACIERRE=" & fechaCierre & " AND IDDIVISION=" & idDivision & " AND VAL.STOCKDISPONIBLE<>0 " & _
			"	) TG " & _
			"	LEFT JOIN [Database].[dbo].[CGT020A] CGT ON LEFT(TG.CUENTA,9) COLLATE Modern_Spanish_CI_AS  = CONVERT(VARCHAR(9),CGT.CUENTA) " & _			
			"	and CIA = '" & getCIADivision(idDivision) & "'" & _
			"	WHERE TG.CUENTA <> '' " & _
			"	GROUP BY TG.CUENTA, CGT.NOMCUE " & _
			"ORDER BY TG.CUENTA"
	Call logdebug(strSQL)
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set obtenerSaldoCuentas = rs
End Function
'-----------------------------------------------------------------------------------------------

'*************************************************************************************
'***************************** COMIENZO DE LA PAGINA ***********************************
'*************************************************************************************
dim RPT_FechaCierre, RPT_IdDivision, RPT_Moneda
dim oPDF, currentY, contador, color, nroPagina

RPT_FechaCierre = GF_Parametros7("fechaCierre", "", 6)
RPT_IdDivision = GF_Parametros7("idDivision", 0, 6)
nroPagina = 1
RPT_Moneda = MONEDA_PESO


'------------------------------------------------------------------------------------------

Set oPDF = GF_createPDF("PDFTemp")
Call GF_setPDFMODE(PDF_STREAM_MODE)
call armadoPDF()
'Response.End 
Call GF_closePDF(oPDF)
'------------------------------------------------------------------------------------------
%>