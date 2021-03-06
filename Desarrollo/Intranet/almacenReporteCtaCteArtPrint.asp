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
Const PAGE_SPACE_DETAIL = 30
Const PAGE_SPACE_SALDOS = 16
Const HHMMSS_FECHA_CIERRE = "235959"
Const CODIGO_REM_REMITO = "REM"
Const CODIGO_REM_ANULACION = "XEM"

'constantes de posicion
Const PAGE_INIT				= 10
Const PAGE_INIT_COMPROBANTE = 60 '-10
Const PAGE_INIT_NUMERO		= 75 '-10
Const PAGE_INIT_FILTROS		= 100
Const PAGE_INIT_MOVIMIENTO	= 133 '-10
Const PAGE_INIT_MOV_E		= 174 '-10
Const PAGE_INIT_MOV_SALDO	= 215 '-10
Const PAGE_INIT_MOV_SALDO_E	= 252 '-10
Const PAGE_INIT_MOV_SALDO_S = 289 '-10
Const PAGE_INIT_PRECIO_COMPRA= 330
Const PAGE_INIT_PRECIO_U	= 385
Const PAGE_INIT_IMPORTE		= 440
Const PAGE_INIT_IMP_SALDO	= 510


'constantes de ancho
Const PAGE_WIDTH			= 570
Const PAGE_WIDTH_FECHA		= 50 '-10
Const PAGE_WIDTH_COMPROBANTE= 73 '-10
Const PAGE_WIDTH_MOVIMIENTO	= 82
Const PAGE_WIDTH_PRECIO_U	= 55
Const PAGE_WIDTH_PRECIO_COMPRA= 55
Const PAGE_WIDTH_IMPORTE	= 140
Const PAGE_WIDTH_SALDO_TOTAL= 200
Const PAGE_WIDTH_CD			= 15
Const PAGE_WIDTH_NUMERO		= 58 '-10
Const PAGE_WIDTH_MOV_SALDO	= 115
'Const PAGE_WIDTH_MOV_ING_EGR= 100
'Const PAGE_WIDTH_MOV_SALDO	= 100
Const PAGE_WIDTH_MOV_I= 41
Const PAGE_WIDTH_MOV_E= 41
'Const PAGE_WIDTH_MOV_S= 30
Const PAGE_WIDTH_MOV_S_I= 37
Const PAGE_WIDTH_MOV_S_E= 37
Const PAGE_WIDTH_MOV_S_S= 41

Const PAGE_WIDTH_IMP_ING_EGR= 70
Const PAGE_WIDTH_IMP_SALDO	= 70
Const PAGE_WIDTH_CABECERA	= 580
Const PAGE_WIDTH_TIT_FILTROS= 50
Const PAGE_WIDTH_FILTROS	= 200

'------------------------------------------------------------------------------------------
Function armadoPDF()
	currentY = armadoCabecera(GF_TRADUCIR("Reporte de Cuenta Corriente por Articulo"), RPT_Almacen, RPT_FechaDesde, RPT_FechaHasta, RPT_idArticulo)
	Call armadoDetalleMovimientos(RPT_idArticulo, RPT_Almacen, RPT_FechaDesde, RPT_FechaHasta)
end Function
'------------------------------------------------------------------------------------------
Function armadoCabecera(p_Titulo, idAlmacen, fechaDesde, fechaHasta, idArticulo)
	dim yInicio, ySeparation
	yInicio = PAGE_TOP_INIT
	ySeparation = 0

	call dibujarTitulo(p_Titulo)
	call GF_setFont(oPDF,"COURIER",8,0)

	auxArt = GF_Traducir("Todos")
	Call GF_writeTextAlign(oPDF, PAGE_INIT, yInicio, GF_TRADUCIR("Articulo........:")		, PAGE_WIDTH_TIT_FILTROS, PDF_ALIGN_LEFT)
	if (idArticulo <> 0) then 	
		call getArticuloFull(idArticulo, auxArtDS, auxArtAbrev)
		auxArt = idArticulo & " - " & auxArtDS
	end if
	Call GF_writeTextAlign(oPDF, PAGE_INIT_FILTROS, yInicio, ucase(auxArt), PAGE_WIDTH_FILTROS, PDF_ALIGN_LEFT)
	ySeparation = ySeparation + SEPARATION
	
	auxAlmacen = GF_Traducir("Todos")
	Call GF_writeTextAlign(oPDF, PAGE_INIT, yInicio + ySeparation, GF_TRADUCIR("Almacen.........:")		, PAGE_WIDTH_TIT_FILTROS, PDF_ALIGN_LEFT)
	if (idAlmacen <> 0) then 
		set rsAlmacen = obtenerListaAlmacenes(idAlmacen)
		if not rsAlmacen.eof then auxAlmacen = rsAlmacen("CDALMACEN") & " - " & rsAlmacen("DSALMACEN")
	end if
	Call GF_writeTextAlign(oPDF, PAGE_INIT_FILTROS, yInicio + ySeparation, ucase(auxAlmacen), PAGE_WIDTH_FILTROS, PDF_ALIGN_LEFT)
	ySeparation = ySeparation + SEPARATION

	auxFechaDesde = GF_Traducir("Todas") 
	Call GF_writeTextAlign(oPDF, PAGE_INIT, yInicio + ySeparation, GF_TRADUCIR("Fecha Desde.....:")	, PAGE_WIDTH_TIT_FILTROS, PDF_ALIGN_LEFT)
	if (fechaDesde <> "") then auxFechaDesde = fechaDesde 
	Call GF_writeTextAlign(oPDF, PAGE_INIT_FILTROS, yInicio + ySeparation, ucase(auxFechaDesde), PAGE_WIDTH_FILTROS, PDF_ALIGN_LEFT)
	ySeparation = ySeparation + SEPARATION

	auxFechaHasta = GF_Traducir("Todas") 
	Call GF_writeTextAlign(oPDF, PAGE_INIT, yInicio + ySeparation, GF_TRADUCIR("Fecha Hasta.....:")	, PAGE_WIDTH_TIT_FILTROS, PDF_ALIGN_LEFT)
	if (fechaHasta <> "") then auxFechaHasta = fechaHasta  
	Call GF_writeTextAlign(oPDF, PAGE_INIT_FILTROS, yInicio + ySeparation, ucase(auxFechaHasta), PAGE_WIDTH_FILTROS, PDF_ALIGN_LEFT)
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
	Call GF_verticalLine(oPDF, PAGE_INIT_COMPROBANTE, yInicio, yTotal - yInicio)
	Call GF_verticalLine(oPDF, PAGE_INIT_MOVIMIENTO, yInicio, yTotal - yInicio)
	Call GF_verticalLine(oPDF, PAGE_INIT_PRECIO_COMPRA, yInicio, yTotal - yInicio)
	Call GF_verticalLine(oPDF, PAGE_INIT_IMPORTE, yInicio, yTotal - yInicio)
	Call GF_verticalLine(oPDF, PAGE_INIT_MOV_SALDO, yInicio, yTotal - yInicio)
end Function
'------------------------------------------------------------------------------------------
Function dibujarTitulosDetalle(pY)
	Call dibujarTitulosDetalleBox(pY)
	call GF_setFont(oPDF,"ARIAL",8,8)
	call GF_setFontColor("FFFFFF")
	Call GF_writeTextAlign(oPDF, PAGE_INIT, pY+9, GF_TRADUCIR("FECHA"), PAGE_WIDTH_FECHA, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_COMPROBANTE, pY+2, GF_TRADUCIR("COMPROBANTE"), PAGE_WIDTH_COMPROBANTE, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_MOVIMIENTO, pY+2, GF_TRADUCIR("MOVIMIENTO"), PAGE_WIDTH_MOVIMIENTO, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_MOV_SALDO, pY+2, GF_TRADUCIR("SALDO"), PAGE_WIDTH_MOV_SALDO, PDF_ALIGN_CENTER)
	
	'Call GF_writeTextAlign(oPDF, PAGE_INIT_PRECIO_U, pY+9, GF_TRADUCIR("P/UNIDAD"), PAGE_WIDTH_PRECIO_U, PDF_ALIGN_CENTER)
	'Call GF_writeTextAlign(oPDF, PAGE_INIT_PRECIO_COMPRA, pY+9, GF_TRADUCIR("P/COMPRA"), PAGE_WIDTH_PRECIO_COMPRA, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_PRECIO_COMPRA, pY+2, GF_TRADUCIR("PRECIO POR UNIDAD"), PAGE_WIDTH_PRECIO_U*2, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_PRECIO_COMPRA, pY+15, GF_TRADUCIR("COMPRA"), PAGE_WIDTH_PRECIO_COMPRA, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_PRECIO_U, pY+15, GF_TRADUCIR("PA�OL"), PAGE_WIDTH_PRECIO_U, PDF_ALIGN_CENTER)
	
	
	Call GF_writeTextAlign(oPDF, PAGE_INIT_IMPORTE, pY+2, GF_TRADUCIR("IMPORTE"), PAGE_WIDTH_IMPORTE, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_COMPROBANTE, pY+15, GF_TRADUCIR("CD"), PAGE_WIDTH_CD, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_NUMERO, pY+15, GF_TRADUCIR("NRO"), PAGE_WIDTH_NUMERO, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_MOVIMIENTO , pY+15, GF_TRADUCIR("EXI"), PAGE_WIDTH_MOV_I, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_MOV_E, pY+15, GF_TRADUCIR("SOB"), PAGE_WIDTH_MOV_E, PDF_ALIGN_CENTER)
	
	Call GF_writeTextAlign(oPDF, PAGE_INIT_MOV_SALDO, pY+15, GF_TRADUCIR("EXI"), PAGE_WIDTH_MOV_S_I, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_MOV_SALDO_E, pY+15, GF_TRADUCIR("SOB"), PAGE_WIDTH_MOV_S_E, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_MOV_SALDO_S, pY+15, GF_TRADUCIR("TOTAL"), PAGE_WIDTH_MOV_S_S, PDF_ALIGN_CENTER)
	
	Call GF_writeTextAlign(oPDF, PAGE_INIT_IMPORTE, pY+15, GF_TRADUCIR("ING/EGR (+/-)"), PAGE_WIDTH_IMP_ING_EGR, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_IMP_SALDO, pY+15, GF_TRADUCIR("SALDO"), PAGE_WIDTH_IMP_SALDO, PDF_ALIGN_CENTER)
	call GF_setFontColor("000000")
end Function			
'------------------------------------------------------------------------------------------
Function dibujarTitulosDetalleBox(pY)
	Call GF_squareBox(oPDF, PAGE_INIT, pY, PAGE_WIDTH, 26, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_verticalLine(oPDF, PAGE_INIT_COMPROBANTE, pY, 26)
	Call GF_verticalLine(oPDF, PAGE_INIT_MOVIMIENTO, pY, 26)
	Call GF_verticalLine(oPDF, PAGE_INIT_PRECIO_COMPRA, pY, 26)
	Call GF_verticalLine(oPDF, PAGE_INIT_IMPORTE, pY, 26)
	Call GF_horizontalLine(oPDF, PAGE_INIT_COMPROBANTE, pY + 13, PAGE_WIDTH_COMPROBANTE)
	Call GF_horizontalLine(oPDF, PAGE_INIT_MOVIMIENTO, pY + 13, PAGE_WIDTH_MOVIMIENTO)
	Call GF_horizontalLine(oPDF, PAGE_INIT_MOV_SALDO, pY + 13, PAGE_WIDTH_MOV_SALDO)
	Call GF_horizontalLine(oPDF, PAGE_INIT_IMPORTE, pY + 13, PAGE_WIDTH_IMPORTE)
	Call GF_verticalLine(oPDF, PAGE_INIT_NUMERO, pY + 13, 13)
	Call GF_verticalLine(oPDF, PAGE_INIT_MOV_SALDO, pY, 26)
	Call GF_verticalLine(oPDF, PAGE_INIT_IMP_SALDO, pY + 13, 13)
	Call GF_horizontalLine(oPDF, PAGE_INIT_PRECIO_COMPRA, pY + 13, PAGE_WIDTH_PRECIO_COMPRA*2)
	
	Call GF_verticalLine(oPDF, PAGE_INIT_PRECIO_U, pY+13, 13)
	Call GF_verticalLine(oPDF, PAGE_INIT_MOV_E, pY+13, 13)
	Call GF_verticalLine(oPDF, PAGE_INIT_MOV_SALDO_E, pY+13, 13)
	Call GF_verticalLine(oPDF, PAGE_INIT_MOV_SALDO_S, pY+13, 13)

end Function
'------------------------------------------------------------------------------------------
Function dibujarSaldosIniciales(p_currentY, p_totalSaldoMovExi, p_totalSaldoMovSob, p_totalSaldoMovSal, p_totalSaldoIMP, pvalor)
	Call GF_horizontalLine(oPDF,PAGE_INIT,p_currentY,PAGE_WIDTH)
	Call GF_squareBox(oPDF, PAGE_INIT, p_currentY + 1, PAGE_WIDTH, SEPARATION + 3, 0, "#CECEF6", "", 0, PDF_SQUARE_NORMAL)
	call GF_setFont(oPDF,"COURIER",8,8)
	Call GF_writeTextAlign(oPDF, PAGE_INIT, p_currentY + SEPARATION/3, GF_TRADUCIR("SALDO INICIAL DEL ARTICULO"), PAGE_WIDTH_SALDO_TOTAL, PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_MOV_SALDO, p_currentY + SEPARATION/3, getFormatSaldo(p_totalSaldoMovExi), PAGE_WIDTH_MOV_S_I -2, PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_MOV_SALDO_E, p_currentY + SEPARATION/3, getFormatSaldo(p_totalSaldoMovSob), PAGE_WIDTH_MOV_S_E-2, PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_MOV_SALDO_S, p_currentY + SEPARATION/3, getFormatSaldo(p_totalSaldoMovSal), PAGE_WIDTH_MOV_S_S-2, PDF_ALIGN_RIGHT)		
	Call GF_writeTextAlign(oPDF,PAGE_INIT_PRECIO_U, p_currentY + SEPARATION/3, getSimboloMoneda(RPT_Moneda) & " " & GF_EDIT_DECIMALS(pvalor,2), PAGE_WIDTH_PRECIO_U-2, PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_IMP_SALDO, p_currentY + SEPARATION/3, getSimboloMoneda(RPT_Moneda) & " " & GF_EDIT_DECIMALS(p_totalSaldoIMP, 2), PAGE_WIDTH_IMP_SALDO-2, PDF_ALIGN_RIGHT)
	Call GF_horizontalLine(oPDF,PAGE_INIT, p_currentY + SEPARATION + 5, PAGE_WIDTH)
end Function
'------------------------------------------------------------------------------------------
Function dibujarSaldosFinales(p_currentY, p_totalSaldoMovExi, p_totalSaldoMovSob, p_totalSaldoMovSal, p_totalSaldoIMP, pPrecioFinal)
	Call GF_squareBox(oPDF, PAGE_INIT, p_currentY + 1, PAGE_WIDTH, SEPARATION + 3, 0, "#CECEF6", "", 0, PDF_SQUARE_NORMAL)
	Call GF_horizontalLine(oPDF,PAGE_INIT,p_currentY,PAGE_WIDTH)
	call GF_setFont(oPDF,"COURIER",8,8)
	Call GF_writeTextAlign(oPDF, PAGE_INIT, p_currentY + SEPARATION/3, GF_TRADUCIR("SALDO FINAL DEL ARTICULO"), PAGE_WIDTH_SALDO_TOTAL, PDF_ALIGN_LEFT)

	Call GF_writeTextAlign(oPDF, PAGE_INIT_MOV_SALDO, p_currentY + SEPARATION/3, getFormatSaldo(p_totalSaldoMovExi), PAGE_WIDTH_MOV_S_I -2, PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_MOV_SALDO_E, p_currentY + SEPARATION/3, getFormatSaldo(p_totalSaldoMovSob), PAGE_WIDTH_MOV_S_E-2, PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_MOV_SALDO_S, p_currentY + SEPARATION/3, getFormatSaldo(p_totalSaldoMovSal), PAGE_WIDTH_MOV_S_S-2, PDF_ALIGN_RIGHT)		
	Call GF_writeTextAlign(oPDF,PAGE_INIT_PRECIO_U, p_currentY + SEPARATION/3, getSimboloMoneda(RPT_Moneda) & " " & GF_EDIT_DECIMALS(pPrecioFinal,2), PAGE_WIDTH_PRECIO_U-2, PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_IMP_SALDO, p_currentY + SEPARATION/3, getSimboloMoneda(RPT_Moneda) & " " & GF_EDIT_DECIMALS(p_totalSaldoIMP, 2), PAGE_WIDTH_IMP_SALDO-2, PDF_ALIGN_RIGHT)
	Call GF_horizontalLine(oPDF,PAGE_INIT, p_currentY + SEPARATION + 5, PAGE_WIDTH)
end Function
'------------------------------------------------------------------------------------------
Function dibujarMovimiento(pY, fecha, cd, nro, existencia, sobrante, importeNuevo, precioUltimaCompra, pTotalSaldoMovExi, pTotalSaldoMovSob, pTotalSaldoMovSal, pTotalSaldoIMP, valor, ptipoMovimiento)
	dim aux
	call GF_setFont(oPDF,"COURIER",8,0)
	Call GF_setFontColor("#000000")
	g_contador = g_contador + 1
	if (g_contador mod 2) then
		color = "#FFFFFF"
	else
		color = "#E0E0F8"
	end if
	Call GF_squareBox(oPDF, PAGE_INIT, pY, PAGE_WIDTH, 10, 0, color, color, 1, PDF_SQUARE_NORMAL)	

	if (fecha <> "") then Call GF_writeTextAlign(oPDF,PAGE_INIT, pY, GF_FN2DTE(fecha), PAGE_WIDTH_FECHA, PDF_ALIGN_CENTER)
	if (cd <> "") then Call GF_writeTextAlign(oPDF,PAGE_INIT_COMPROBANTE, pY, cd, PAGE_WIDTH_CD, PDF_ALIGN_CENTER)	
	if (nro <> "") then 
		aux = nro
		if ((cd <> CODIGO_REM_REMITO) and (cd <> CODIGO_REM_ANULACION)) then aux = getNroVale(nro)
		Call GF_writeTextAlign(oPDF,PAGE_INIT_NUMERO, pY, aux, PAGE_WIDTH_NUMERO - 2, PDF_ALIGN_RIGHT)
	end if
	
	if (cd = CODIGO_REM_ANULACION) then 
		sobrante = sobrante * (-1)	
		existencia = existencia * (-1)
		importeNuevo = importeNuevo * -1
	end if	
	Call GF_writeTextAlign(oPDF,PAGE_INIT_MOV_E , pY, getFormatSaldo(sobrante), PAGE_WIDTH_MOV_E-2, PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,PAGE_INIT_MOVIMIENTO, pY, getFormatSaldo(existencia), PAGE_WIDTH_MOV_I-2, PDF_ALIGN_RIGHT)			
	
	if cdbl(precioUltimaCompra) <> 0 then
		Call GF_writeTextAlign(oPDF,PAGE_INIT_PRECIO_COMPRA, pY, getSimboloMoneda(RPT_Moneda) & " " & GF_EDIT_DECIMALS(precioUltimaCompra, 2), PAGE_WIDTH_PRECIO_COMPRA-2, PDF_ALIGN_RIGHT)
	end if	
	Call GF_writeTextAlign(oPDF,PAGE_INIT_PRECIO_U, pY, getSimboloMoneda(RPT_Moneda) & " " & GF_EDIT_DECIMALS(valor,2), PAGE_WIDTH_PRECIO_U-2, PDF_ALIGN_RIGHT)
		
	Call GF_writeTextAlign(oPDF,PAGE_INIT_MOV_SALDO, pY, getFormatSaldo(pTotalSaldoMovExi), PAGE_WIDTH_MOV_S_I-2, PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,PAGE_INIT_MOV_SALDO_E, pY, getFormatSaldo(pTotalSaldoMovSob), PAGE_WIDTH_MOV_S_E-2, PDF_ALIGN_RIGHT)	
	Call GF_writeTextAlign(oPDF,PAGE_INIT_MOV_SALDO_S, pY, getFormatSaldo(pTotalSaldoMovSal), PAGE_WIDTH_MOV_S_S-2, PDF_ALIGN_RIGHT)
	
	Call GF_writeTextAlign(oPDF,PAGE_INIT_IMPORTE, pY, getSimboloMoneda(RPT_Moneda) & " " & GF_EDIT_DECIMALS(importeNuevo, 2), PAGE_WIDTH_IMP_ING_EGR-2, PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,PAGE_INIT_IMP_SALDO, pY, getSimboloMoneda(RPT_Moneda) & " " & GF_EDIT_DECIMALS(pTotalSaldoIMP, 2), PAGE_WIDTH_IMP_SALDO-2, PDF_ALIGN_RIGHT)	
end Function
'------------------------------------------------------------------------------------------
'realiza los controles y arma el listado de movimientos
Function armadoDetalleMovimientos(idArticulo, idAlmacen, fechaDesde, fechaHasta)
	dim rsMovimientos, totalReg, i, isInicio, aux_Y, idDivision, tipoMovimiento, amdDesde, flagMovimientos
	dim totalSaldoIMP, precioCalculo
	
	amdDesde = GF_DTE2FN(fechaDesde)	
	
	Call obtenerDatosUltimoCierre(idAlmacen, amdDesde, idArticulo, ultimoCierreExistencia, ultimoCierreSobrante, fecha)				
	idDivision = getDivisionAlmacen(idAlmacen)
	
	
	totalSaldoMovExi = ultimoCierreExistencia
	totalSaldoMovSob = ultimoCierreSobrante 
	totalSaldoMovSal = ultimoCierreExistencia + ultimoCierreSobrante
	
	if (fecha <> "") then
		'Tomo el precio al final del d�a del cierre.
		valor = getUltimoPrecio(idDivision, idArticulo, MONEDA_PESOS, fecha & HHMMSS_FECHA_CIERRE)
		totalSaldoIMP = ultimoCierreExistencia * valor
	else
		valor=0
		totalSaldoIMP = 0
	end if
	'Tomo los movimientos a partir del siguiente al ultimo cierre.
	Set rsMovimientos = obtenerListaMovimientos(idArticulo, idAlmacen, fecha, fechaHasta)	
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
		tipoMovimiento = tipoDeMovimiento(rsMovimientos("CD"))
		if (tipoMovimiento <> NO_MODIFICA_STOCK) then			
			existencia = (cDbl(rsMovimientos("EXISTENCIA")))
			sobrante = (cDbl(rsMovimientos("SOBRANTE")))						
			if (cstr(rsMovimientos("FECHA")) < amdDesde) then				
				precioCalculo = rsMovimientos("VLUPESOS")				
				if (rsMovimientos("CD") = CODIGO_REM_REMITO or rsMovimientos("CD") = CODIGO_REM_ANULACION) then  precioCalculo = rsMovimientos("PRECIOCOMPRA")				
				Call calcularTotalSaldos(tipoMovimiento, precioCalculo, existencia, sobrante, importeNuevo, totalSaldoMovExi, totalSaldoMovSob, totalSaldoMovSal, totalSaldoIMP)
			else													
				if (isInicio) then			
					Call dibujarSaldosIniciales(currentY, totalSaldoMovExi,totalSaldoMovSob,totalSaldoMovSal, totalSaldoIMP, valor)
					currentY = currentY + PAGE_SPACE_SALDOS
					aux_Y = currentY
					isInicio = false					
				end if
				if ((tipoMovimiento = DISMINUYE_STOCK) or ((tipoMovimiento = MODIFICA_STOCK) and (left(rsMovimientos("CD"),1) = "X"))) then
				'if (tipoMovimiento = DISMINUYE_STOCK) then
					existencia = existencia * -1
					sobrante = sobrante * -1				
				end if
				valor = rsMovimientos("VLUPESOS")
				precioCalculo = valor
				if (rsMovimientos("CD") = CODIGO_REM_REMITO or rsMovimientos("CD") = CODIGO_REM_ANULACION) then 
					valor = getUltimoPrecio(idDivision, idArticulo, MONEDA_PESOS, rsMovimientos("FECHA") & HHMMSS_FECHA_CIERRE)
					if (isnull(valor)) then valor = 0 else valor = cDbl(valor) end if
					precioCompra = rsMovimientos("PRECIOCOMPRA")
					precioCalculo = precioCompra
				elseif (rsMovimientos("CD") = CODIGO_VS_RECEPCION or rsMovimientos("CD") = CODIGO_VS_RECEPCION_X) then
					valor = getUltimoPrecio(idDivision, idArticulo, MONEDA_PESOS, rsMovimientos("FECHA") & HHMMSS_FECHA_CIERRE)
					if (isnull(valor)) then valor = 0 else valor = cDbl(valor) end if
					precioCompra = rsMovimientos("VLUPESOS")
				else
					precioCompra = 0
				end if	
				call logdebug("existencia:"&rsMovimientos("EXISTENCIA")&" sobrante:"&rsMovimientos("SOBRANTE"))
				Call calcularTotalSaldos(tipoMovimiento, precioCalculo, existencia, sobrante, importeNuevo, totalSaldoMovExi, totalSaldoMovSob, totalSaldoMovSal, totalSaldoIMP)
				call dibujarMovimiento(currentY, rsMovimientos("FECHA"), rsMovimientos("CD"), rsMovimientos("NRO"), existencia, sobrante, importeNuevo, precioCompra, totalSaldoMovExi, totalSaldoMovSob, totalSaldoMovSal, totalSaldoIMP, valor, tipoMovimiento)
				currentY = currentY + SEPARATION
				if (currentY > PAGE_HEIGHT_SIZE) then flagNuevaPagina = true
			end if
		end if
		rsMovimientos.MoveNext()
	wend
	if (not isInicio) then
		Call dibujarSeparadores(currentY, aux_Y)
	else
		Call dibujarSaldosIniciales(currentY, totalSaldoMovExi,totalSaldoMovSob,totalSaldoMovSal, totalSaldoIMP, valor)	
		currentY = currentY + PAGE_SPACE_SALDOS + SEPARATION
	end if	
	
	valor = getUltimoPrecio(idDivision, idArticulo, MONEDA_PESOS, fechaHasta & HHMMSS_FECHA_CIERRE)

	if (isnull(valor)) then valor = 0 else valor = cDbl(valor) end if
	Call dibujarSaldosFinales(currentY, totalSaldoMovExi,totalSaldoMovSob,totalSaldoMovSal, totalSaldoIMP, valor)
end Function
'------------------------------------------------------------------------------------------
'Devuleve la fecha, el total de stock y el valor del articulo, del cierre mensual pasado mas proximo a la fechas desde indicada
Function obtenerDatosUltimoCierre(idAlmacen, amdDesde, idArticulo, Byref ucExistencia, Byref ucSobrante, Byref fecha)
	dim rsCierre, strSQL, conn, existencia, sobrante
	ucExistencia = 0
	ucSobrante = 0	
	fecha = ""
	strSQL= "select FECHACIERRE, EXISTENCIA, SOBRANTE "
	strSQL= strSQL & "from TBLCIERRESARTICULOS2 where FECHACIERRE = "
	strSQL= strSQL & "( select max(FECHACIERRE) from TBLCIERRESARTICULOS2 "
	strSQL= strSQL & "where FECHACIERRE < " & amdDesde & " AND IDALMACEN = " & idAlmacen & " AND IDARTICULO= " & idArticulo & ") "
	strSQL= strSQL & "AND IDALMACEN = " & idAlmacen & " AND IDARTICULO= " & idArticulo		
	Call executeQueryDB(DBSITE_SQL_INTRA, rsCierre, "OPEN", strSQL)
	if (not rsCierre.eof) then
		ucExistencia = CDbl(rsCierre("EXISTENCIA"))
		ucSobrante = CDbl(rsCierre("SOBRANTE"))
		fecha = rsCierre("FECHACIERRE")
	end if
end Function
'------------------------------------------------------------------------------------------
Function obtenerListaMovimientos(idArticulo, idAlmacen, fecha, fechaHasta)
	Dim strSQL, rs, conn
	fechaHasta = GF_DTE2FN(fechaHasta)
	strSQL = "(select R.NROREMITO NRO, R.CDREMITO CD, R.FECHA FECHA, R.MOMENTO MOMENTO, PRECIOS.CANTIDAD CANTIDAD, "
	strSQL = strSQL & "(case when E.EXISTENCIA <> 0 then PRECIOS.CANTIDAD else 0 end )EXISTENCIA, "
	strSQL = strSQL & "(case when E.SOBRANTE   <> 0 then PRECIOS.CANTIDAD else 0 end )SOBRANTE, "
	strSQL = strSQL & "0 VLUPESOS, CAST(PRECIOS.PRECIOUNITARIO as bigint) PRECIOCOMPRA "
	strSQL = strSQL & "from (Select * from TBLREMCABECERA where "
	if (fecha <> "") then strSQL = strSQL & "FECHA >" & fecha & " and "
	strSQL = strSQL & " FECHA <= " & fechaHasta & "  and IDALMACEN = " & idAlmacen
	strSQL = strSQL & "	)R inner join (Select * from TBLREMDETALLE where IDARTICULO = " & idArticulo & ") E on R.IDREMITO = E.IDREMITO "
	strSQL = strSQL & " inner join (Select IDREMITO, REL.IDARTICULO, PRECIOUNITARIO, sum(CANTIDAD) CANTIDAD "
	strSQL = strSQL & "					from tblrempic REL"
	strSQL = strSQL & "					INNER JOIN (SELECT idcotizacion, idarticulo, ( Sum(importepesos) / Sum(cantidad) ) PRECIOUNITARIO "
	strSQL = strSQL & "									FROM   tblctzdetalle"
	strSQL = strSQL & "									WHERE  idarticulo = " & idArticulo & " and cantidad > 0 "
	strSQL = strSQL & "									GROUP  BY idcotizacion, idarticulo) CTZ ON REL.idpic = CTZ.idcotizacion	AND REL.idarticulo = CTZ.idarticulo "
    strSQL = strSQL & "              group by IDREMITO, REL.IDARTICULO, PRECIOUNITARIO) PRECIOS on PRECIOS.IDREMITO=R.IDREMITO and PRECIOS.IDARTICULO=E.IDARTICULO)"
	strSQL = strSQL & "union "
	strSQL = strSQL & "(select V.IDVALE NRO, V.CDVALE CD, V.FECHA FECHA, V.MOMENTO MOMENTO, A.CANTIDAD CANTIDAD, A.EXISTENCIA, A.SOBRANTE, A.VLUPESOS, 0 PRECIOCOMPRA "
	strSQL = strSQL & "from TBLVALESCABECERA V inner join TBLVALESDETALLE A on V.IDVALE = A.IDVALE "
	strSQL = strSQL & "where "
	if (fecha <> "") then strSQL = strSQL & "V.FECHA >" & fecha & " and " 
	strSQL = strSQL & "  V.FECHA <= " & fechaHasta & "  and V.IDALMACEN = " & idAlmacen & " and A.IDARTICULO = " & idArticulo & ") "
	strSQL = strSQL & "order by FECHA, MOMENTO"
	'Response.write strSQL
	'Response.End 
	Call logdebug(strSQL)
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set obtenerListaMovimientos = rs
End Function
'-----------------------------------------------------------------------------------------------
'devuelve true si el movimiento es un ingreso
Function tipoDeMovimiento(p_CD)
	dim rs, strSQl, cn, rtrn
	strSQL = "select STOCKFISICO from TBLVALESMAESTRO WHERE CDTIPO = '" & p_CD & "' "
	call logdebug(strSQL)
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then rtrn = cDbl(rs("STOCKFISICO"))
	tipoDeMovimiento = rtrn	
end Function
'-----------------------------------------------------------------------------------------------
Function calcularTotalSaldos(pTipoMovimiento, pPrecio, pExistencia, pSobrante, Byref pImporte, byref pSaldoExistencia, byref pSaldoSobrante, byref pSaldoSaldo, byref pImporteSaldo)
	dim aux	
	'Response.Write "<BR>EXI(" & pExistencia & ")"
	'Response.Write "PRE(" & pPrecio & ")"	
	pImporte = cdbl(pExistencia) * cdbl(pPrecio)
	if (pTipoMovimiento = AUMENTA_STOCK) then
		pImporteSaldo		= Round(pImporteSaldo + abs(cdbl(pImporte)), 3)
		pSaldoExistencia	= Round(CDBL(pSaldoExistencia + abs(cdbl(pExistencia))), 3)
		pSaldoSobrante		= Round(pSaldoSobrante + abs(cdbl(pSobrante)), 3)
		pSaldoSaldo			= Round(pSaldoSaldo + abs(cdbl(pExistencia)) + abs(cdbl(pSobrante)), 3)
	elseif (pTipoMovimiento = DISMINUYE_STOCK) then
		pImporteSaldo		= Round(pImporteSaldo - abs(cdbl(pImporte)), 3)
		pSaldoExistencia	= Round(CDBL(pSaldoExistencia - abs(cdbl(pExistencia))), 3)
		pSaldoSobrante		= Round(totalSaldoMovSob - abs(cdbl(pSobrante)), 3)
		pSaldoSaldo			= Round(pSaldoSaldo - abs(cdbl(pExistencia)) - abs(cdbl(pSobrante)), 3)
	else
		pImporteSaldo		= Round(pImporteSaldo + cdbl(pImporte), 3)
		pSaldoExistencia	= Round(CDBL(pSaldoExistencia) + cdbl(pExistencia), 3)
		pSaldoSobrante		= Round(totalSaldoMovSob + cdbl(pSobrante), 3)
		pSaldoSaldo			= Round(pSaldoSaldo + cdbl(pExistencia) + cdbl(pSobrante), 3)
	end if
	'Response.Write "SALDO(" & pImporteSaldo & ")"	
end Function
'-----------------------------------------------------------------------------------------------
'Function encargada de devolver el formato apropiado para los Saldos(Existencia,Sobrante y Total).
'Si el Saldo pasado por parametro es entero, se mustra sin los decimales. 
'Si el Saldo pasado por parametro tiene decimales, se mustra 3 digitos del decimal.
Function getFormatSaldo(pSaldo)	
	if isNumeric(pSaldo) then
		if cDbl(pSaldo) = pSaldo then			
			getFormatSaldo = pSaldo
		else
			getFormatSaldo = GF_EDIT_DECIMALS(Cdbl(pSaldo)*1000,3)
	   end if
	end if
End Function
'*************************************************************************************
'***************************** COMIENZO DE LA PAGINA ***********************************
'*************************************************************************************
dim RPT_Almacen, RPT_FechaDesde, RPT_FechaHasta, RPT_idArticulo, RPT_cdArticulo, RPT_dsArticulo, RPT_Moneda
dim oPDF, currentY, g_contador, color, nroPagina
dim stock, fecha, valor, totalSaldoIMP, ultimoCierreExistencia, ultimoCierreSobrante
dim totalSaldoMovExi, totalSaldoMovSob, totalSaldoMovSal


RPT_idArticulo = GF_Parametros7("idArticulo", 0, 6)
RPT_Almacen = GF_Parametros7("idAlmacen", "", 6)
RPT_FechaDesde = GF_Parametros7("issuedate", "", 6)
RPT_FechaHasta = GF_Parametros7("closingdate", "", 6)

nroPagina = 1

stock = 0
valor = 0
fecha = ""
totalSaldoMovExi = 0 
totalSaldoMovSob = 0 
totalSaldoMovSal = 0
totalSaldoIMP = 0
RPT_Moneda = MONEDA_PESO

'------------------------------------------------------------------------------------------
Set oPDF = GF_createPDF("PDFTemp")
Call GF_setPDFMODE(PDF_STREAM_MODE)
call armadoPDF()
'Response.End 
Call GF_closePDF(oPDF)
'------------------------------------------------------------------------------------------
%>