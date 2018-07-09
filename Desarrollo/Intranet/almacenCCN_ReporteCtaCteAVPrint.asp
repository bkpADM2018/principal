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
Const PAGE_SPACE_DETAIL = 18
Const PAGE_SPACE_SALDOS = 16

'constantes de posicion
Const PAGE_INIT				= 10
Const PAGE_INIT_FECHA	    = 10
Const PAGE_INIT_IDVALE		= 70
Const PAGE_INIT_CDVALE		= 110
Const PAGE_INIT_CANTIDAD	= 150
Const PAGE_INIT_PRECIO		= 200
Const PAGE_INIT_IMPORTE		= 270
Const PAGE_INIT_TIPO_VAL	= 340
Const PAGE_INIT_CTA_INV		= 420
Const PAGE_INIT_CTA_GTOS	= 480
Const PAGE_INIT_CCOSTOS		= 540

Const PAGE_INIT_FILTROS		= 100

'constantes de ancho
Const PAGE_WIDTH			= 570
Const PAGE_WIDTH_FECHA		= 60
Const PAGE_WIDTH_IDVALE		= 40 
Const PAGE_WIDTH_CDVALE		= 40
Const PAGE_WIDTH_CANTIDAD	= 50
Const PAGE_WIDTH_PRECIO		= 70
Const PAGE_WIDTH_IMPORTE	= 70
Const PAGE_WIDTH_TIPO_VAL	= 80
Const PAGE_WIDTH_CTA_INV	= 60
Const PAGE_WIDTH_CTA_GTOS	= 60
Const PAGE_WIDTH_CCOSTOS	= 40 

Const PAGE_WIDTH_CABECERA	= 580
Const PAGE_WIDTH_TIT_FILTROS= 50
Const PAGE_WIDTH_FILTROS	= 200

'------------------------------------------------------------------------------------------
Function armadoPDF()
	currentY = armadoCabecera(GF_TRADUCIR("Reporte de Cuenta Corriente Contable"), RPT_IdDivision, RPT_IdArticulo, RPT_IdVale, RPT_FechaDesde, RPT_FechaHasta)
	Call armadoDetalleMovimientos(RPT_idArticulo, RPT_Almacen, RPT_FechaDesde, RPT_FechaHasta)
end Function
'------------------------------------------------------------------------------------------
Function armadoCabecera(p_Titulo, pIdDivision, pIdArticulo, pIdVale, pFechaDesde, pFechaHasta)
	dim yInicio, ySeparation
	yInicio = PAGE_TOP_INIT
	ySeparation = 0

	call dibujarTitulo(p_Titulo)
	call GF_setFont(oPDF,"COURIER",8,0)
	if isVale then
		Call GF_writeTextAlign(oPDF, PAGE_INIT, yInicio, GF_TRADUCIR("Vale............:")		, PAGE_WIDTH_TIT_FILTROS, PDF_ALIGN_LEFT)
		Call GF_writeTextAlign(oPDF, PAGE_INIT_FILTROS, yInicio, getNroVale(pIdVale) & " (" & pIdVale & ")", PAGE_WIDTH_FILTROS, PDF_ALIGN_LEFT)
	else
		Call GF_writeTextAlign(oPDF, PAGE_INIT, yInicio, GF_TRADUCIR("Articulo........:")		, PAGE_WIDTH_TIT_FILTROS, PDF_ALIGN_LEFT)
		call getArticuloFull(pIdArticulo, auxArtDS, auxArtAbrev)
		auxArt = pIdArticulo & " - " & auxArtDS
		Call GF_writeTextAlign(oPDF, PAGE_INIT_FILTROS, yInicio, ucase(auxArt), PAGE_WIDTH_FILTROS, PDF_ALIGN_LEFT)
	end if
	ySeparation = ySeparation + SEPARATION
	
	Call GF_writeTextAlign(oPDF, PAGE_INIT, yInicio + ySeparation, GF_TRADUCIR("Division........:")		, PAGE_WIDTH_TIT_FILTROS, PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_FILTROS, yInicio + ySeparation, ucase(getDivisionDS(pIdDivision)), PAGE_WIDTH_FILTROS, PDF_ALIGN_LEFT)
	ySeparation = ySeparation + SEPARATION

	Call GF_writeTextAlign(oPDF, PAGE_INIT, yInicio + ySeparation, GF_TRADUCIR("Fecha Desde.....:")	, PAGE_WIDTH_TIT_FILTROS, PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_FILTROS, yInicio + ySeparation, GF_FN2DTE(pFechaDesde), PAGE_WIDTH_FILTROS, PDF_ALIGN_LEFT)
	ySeparation = ySeparation + SEPARATION

	Call GF_writeTextAlign(oPDF, PAGE_INIT, yInicio + ySeparation, GF_TRADUCIR("Fecha Hasta.....:")	, PAGE_WIDTH_TIT_FILTROS, PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_FILTROS, yInicio + ySeparation, GF_FN2DTE(pFechaHasta), PAGE_WIDTH_FILTROS, PDF_ALIGN_LEFT)
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
Function dibujarTitulosDetalle(pY)
	Call dibujarTitulosDetalleBox(pY)
	call GF_setFont(oPDF,"ARIAL",8,8)
	call GF_setFontColor("FFFFFF")
	Call GF_writeTextAlign(oPDF, PAGE_INIT_FECHA, pY+5, GF_TRADUCIR("FECHA"), PAGE_WIDTH_FECHA, PDF_ALIGN_CENTER)
	if isVale then
		Call GF_writeTextAlign(oPDF, PAGE_INIT_IDVALE, pY+5, GF_TRADUCIR("ARTICULO"), PAGE_WIDTH_IDVALE+PAGE_WIDTH_CDVALE, PDF_ALIGN_CENTER)
	else
		Call GF_writeTextAlign(oPDF, PAGE_INIT_IDVALE, pY+5, GF_TRADUCIR("IDVALE"), PAGE_WIDTH_IDVALE, PDF_ALIGN_CENTER)
		Call GF_writeTextAlign(oPDF, PAGE_INIT_CDVALE, pY+5, GF_TRADUCIR("CDVALE"), PAGE_WIDTH_CDVALE, PDF_ALIGN_CENTER)
	end if
	Call GF_writeTextAlign(oPDF, PAGE_INIT_CANTIDAD, pY+5, GF_TRADUCIR("CANTIDAD"), PAGE_WIDTH_CANTIDAD, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_PRECIO, pY+5, GF_TRADUCIR("PRECIO"), PAGE_WIDTH_PRECIO, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_IMPORTE, pY+5, GF_TRADUCIR("IMPORTE"), PAGE_WIDTH_IMPORTE, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_TIPO_VAL, pY+5, GF_TRADUCIR("TIPO"), PAGE_WIDTH_TIPO_VAL, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_CTA_INV, pY+5, GF_TRADUCIR("CTA INV"), PAGE_WIDTH_CTA_INV, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_CTA_GTOS, pY+5, GF_TRADUCIR("CTA GTOS"), PAGE_WIDTH_CTA_GTOS, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_CCOSTOS, pY+5, GF_TRADUCIR("COSTOS"), PAGE_WIDTH_CCOSTOS, PDF_ALIGN_CENTER)
	call GF_setFontColor("000000")
end Function			
'------------------------------------------------------------------------------------------
Function dibujarTitulosDetalleBox(pY)
	Call GF_squareBox(oPDF, PAGE_INIT, pY, PAGE_WIDTH, 18, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_verticalLine(oPDF, PAGE_INIT_IDVALE, pY, 18)
	if not isVale then Call GF_verticalLine(oPDF, PAGE_INIT_CDVALE, pY, 18)
	Call GF_verticalLine(oPDF, PAGE_INIT_CANTIDAD, pY, 18)
	Call GF_verticalLine(oPDF, PAGE_INIT_PRECIO, pY, 18)
	Call GF_verticalLine(oPDF, PAGE_INIT_IMPORTE, pY, 18)
	Call GF_verticalLine(oPDF, PAGE_INIT_TIPO_VAL, pY, 18)
	Call GF_verticalLine(oPDF, PAGE_INIT_CTA_INV, pY, 18)
	Call GF_verticalLine(oPDF, PAGE_INIT_CTA_GTOS, pY, 18)
	Call GF_verticalLine(oPDF, PAGE_INIT_CCOSTOS, pY, 18)
end Function
'------------------------------------------------------------------------------------------
Function dibujarLineaFinal(p_currentY)
	Call GF_squareBox(oPDF, PAGE_INIT, p_currentY + 1, PAGE_WIDTH, SEPARATION + 3, 0, "#CECEF6", "", 0, PDF_SQUARE_NORMAL)
	Call GF_horizontalLine(oPDF,PAGE_INIT,p_currentY,PAGE_WIDTH)
	call GF_setFont(oPDF,"COURIER",8,8)
	Call GF_writeTextAlign(oPDF, PAGE_INIT, p_currentY + SEPARATION/3, GF_TRADUCIR("FIN DEL REPORTE"), PAGE_WIDTH, PDF_ALIGN_CENTER)
	Call GF_horizontalLine(oPDF,PAGE_INIT, p_currentY + SEPARATION + 5, PAGE_WIDTH)
end Function
'------------------------------------------------------------------------------------------
Function dibujarMovimiento(pY, pFecha, pIdVale, pCdVale, pCantidad, pPrecio, pImporte, pTipoValuacion, pCtaInventario, pCtaGastos, pCostos)
	call GF_setFont(oPDF,"COURIER",8,0)
	contador = contador + 1
	
	if cdbl(pFecha) <> cdbl(myFechaAnterior) then 
		if myFechaAnterior <> "" then Call GF_horizontalLine(oPDF,PAGE_INIT,pY,PAGE_WIDTH)
		myFechaAnterior = pFecha
	end if	
	if contador mod 2 then
		color = "#FFFFFF"
	else
		color = "#E0E0F8"
	end if
	Call GF_squareBox(oPDF, PAGE_INIT, pY+1 , PAGE_WIDTH, 10, 0, color, color, 1, PDF_SQUARE_NORMAL)
	Call GF_writeTextAlign(oPDF,PAGE_INIT_FECHA, pY, GF_FN2DTE(pFecha), PAGE_WIDTH_FECHA, PDF_ALIGN_CENTER)
	if isVale then
		Call GF_writeTextAlign(oPDF,PAGE_INIT_IDVALE, pY, pIdVale & "-" & left(pCdVale,11), PAGE_WIDTH_IDVALE+PAGE_WIDTH_CDVALE, PDF_ALIGN_LEFT)	
	else
		Call GF_writeTextAlign(oPDF,PAGE_INIT_IDVALE, pY, pIdVale, PAGE_WIDTH_IDVALE, PDF_ALIGN_CENTER)	
		Call GF_writeTextAlign(oPDF,PAGE_INIT_CDVALE, pY, pCdVale, PAGE_WIDTH_CDVALE, PDF_ALIGN_CENTER)
	end if
	Call GF_writeTextAlign(oPDF,PAGE_INIT_CANTIDAD, pY, GF_EDIT_DECIMALS(cdbl(pCantidad)*1000,3), PAGE_WIDTH_CANTIDAD-2, PDF_ALIGN_RIGHT)	
	Call GF_writeTextAlign(oPDF,PAGE_INIT_PRECIO, pY, getSimboloMoneda(RPT_Moneda) & " " & GF_EDIT_DECIMALS(cDbl(cdbl(pPrecio)/100),2), PAGE_WIDTH_PRECIO-2, PDF_ALIGN_RIGHT)	
	Call GF_writeTextAlign(oPDF,PAGE_INIT_IMPORTE, pY, getSimboloMoneda(RPT_Moneda) & " " & GF_EDIT_DECIMALS(cDbl(cdbl(pImporte)/100),2), PAGE_WIDTH_IMPORTE-2, PDF_ALIGN_RIGHT)	
	Call GF_writeTextAlign(oPDF,PAGE_INIT_TIPO_VAL, pY, getLeyendaTipoValuacion(pTipoValuacion), PAGE_WIDTH_TIPO_VAL, PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(oPDF,PAGE_INIT_CTA_INV, pY, pCtaInventario, PAGE_WIDTH_CTA_INV, PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(oPDF,PAGE_INIT_CTA_GTOS, pY, pCtaGastos, PAGE_WIDTH_CTA_GTOS, PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(oPDF,PAGE_INIT_CCOSTOS, pY, pCostos, PAGE_WIDTH_CCOSTOS, PDF_ALIGN_CENTER)	
end Function
'------------------------------------------------------------------------------------------
'realiza los controles y arma el listado de movimientos
Function armadoDetalleMovimientos(idArticulo, idAlmacen, fechaDesde, fechaHasta)
	dim rsMovimientos, totalReg, i, isInicio, aux_Y, idDivision, auxId, auxCd
	
	'Tomo los movimientos a partir del siguiente al ultimo cierre.
	Set rsMovimientos = obtenerListaMovimientos(RPT_IdDivision, RPT_IdArticulo, RPT_IdVale, RPT_FechaDesde, RPT_FechaHasta)	
	call dibujarTitulosDetalle(currentY)
	currentY = currentY + PAGE_SPACE_DETAIL
	aux_Y = currentY
	isInicio = true
	flagNuevaPagina = false
	
	while not rsMovimientos.eof
		if (flagNuevaPagina) then
			Call nuevaPagina()
			aux_Y = currentY
			flagNuevaPagina = false
		end if
		if isVale then 
			auxId = rsMovimientos("IDARTICULO")
			auxCd = rsMovimientos("DSARTICULO")
		else
			auxId = rsMovimientos("IDVALE")
			auxCd = rsMovimientos("CDVALE")
		end if
		call dibujarMovimiento(currentY, rsMovimientos("FECHACIERRE"), auxId, auxCd, rsMovimientos("CANTIDAD"), rsMovimientos("VLUPESOS"), CDBL(rsMovimientos("CANTIDAD"))*CDBL(rsMovimientos("VLUPESOS")), rsMovimientos("TIPOVALUACION"), TRIM(rsMovimientos("CUENTAINVENTARIO")), TRIM(rsMovimientos("CUENTAGASTOS")), TRIM(rsMovimientos("CCOSTOS")))
		currentY = currentY + SEPARATION
		if (currentY > PAGE_HEIGHT_SIZE) then flagNuevaPagina = true
		rsMovimientos.MoveNext()
	wend
	Call dibujarLineaFinal(currentY)
end Function

'------------------------------------------------------------------------------------------
Function obtenerListaMovimientos(idDivision, idArticulo, idVale, fechaDesde, fechaHasta)
	Dim strSQL, rs, conn
	if isVale then
		strSQL= "SELECT CC.*, ART.IDARTICULO, ART.DSARTICULO FROM " & _
			"	TBLARTCTACTE CC INNER JOIN TBLARTICULOS ART ON CC.IDARTICULO=ART.IDARTICULO " & _
			" WHERE IDDIVISION=" & idDivision & " AND FECHACIERRE BETWEEN " & fechaDesde & " AND " & fechaHasta & _
			" AND CC.IDVALE=" & idVale & " ORDER BY FECHACIERRE DESC" 
	else
		strSQL= "SELECT CC.*, VC.CDVALE FROM " & _
			"	TBLARTCTACTE CC INNER JOIN TBLVALESCABECERA VC ON CC.IDVALE=VC.IDVALE " & _
			" WHERE IDDIVISION=" & idDivision & " AND FECHACIERRE BETWEEN " & fechaDesde & " AND " & fechaHasta & _
			" AND CC.IDARTICULO=" & idArticulo & " ORDER BY FECHACIERRE DESC, CC.IDVALE" 		
	end if	
	Call logdebug(strSQL)
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set obtenerListaMovimientos = rs
End Function

'-----------------------------------------------------------------------------------------------

'*************************************************************************************
'***************************** COMIENZO DE LA PAGINA ***********************************
'*************************************************************************************
dim RPT_IdDivision, RPT_IdVale, RPT_FechaDesde, RPT_FechaHasta, RPT_IdArticulo, RPT_Moneda
dim oPDF, currentY, contador, color, nroPagina, myFechaAnterior
dim isVale
isvale = false


RPT_IdDivision = GF_Parametros7("idDivision", 0, 6)
RPT_IdArticulo = GF_Parametros7("idArticulo", 0, 6)
RPT_IdVale = GF_Parametros7("idVale", 0, 6)
RPT_FechaDesde = GF_Parametros7("issuedate", "", 6)
RPT_FechaHasta = GF_Parametros7("closingdate", "", 6)
if RPT_IdVale <> 0 then isVale=true
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