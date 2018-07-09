<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosPDF.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<% Response.Buffer = False
Const SEPARATION = 14
Const MARGIN = 0
Const PAGE_HEIGHT_SIZE = 800
Const PAGE_TOP_INIT = 92
Const PAGE_SPACE_DETAIL = 27
Const PAGE_SPACE_SALDOS = 16
Const HHMMSS_FECHA_CIERRE = "235959"
Const CODIGO_REM_REMITO = "REM"
Const CODIGO_REM_ANULACION = "XEM"

'constantes de posicion
Const PAGE_INIT						= 10
Const PAGE_INIT_ARTICULO			= 10
Const PAGE_INIT_PRECIO_INICIO		= 200
Const PAGE_INIT_PRECIO_FINAL		= 275
Const PAGE_INIT_VARIACION			= 350
Const PAGE_INIT_PRECIO_ULT			= 390
Const PAGE_INIT_FECHA_ULT			= 460
Const PAGE_INIT_PIC					= 530
Const PAGE_INIT_FILTROS				= 120

'constantes de ancho
Const PAGE_WIDTH				   = 560
Const PAGE_WIDTH_ARTICULO		   = 190
Const PAGE_WIDTH_PRECIO_INICIO	   = 75
Const PAGE_WIDTH_PRECIO_FINAL	   = 75
Const PAGE_WIDTH_VARIACION		   = 40
Const PAGE_WIDTH_PRECIO_ULT		   = 70
Const PAGE_WIDTH_FECHA_ULT		   = 70
Const PAGE_WIDTH_PIC			   = 40

Const PAGE_WIDTH_CABECERA	= 570
Const PAGE_WIDTH_TIT_FILTROS= 50
Const PAGE_WIDTH_FILTROS	= 200
'-------------------------------------------------------------------------------------------------------------
Function armadoPDF()
	Call drawTitle()
	currentY = drawFilter()	
	Call armadoDetalleMovimientos()
	
End Function
'-------------------------------------------------------------------------------------------------------------
Function dibujarTitulosDetalle(pY)
	Call dibujarTitulosDetalleBox(pY)
	call GF_setFont(oPDF,"ARIAL",6,8)
	call GF_setFontColor("FFFFFF")
	Call GF_writeTextAlign(oPDF, PAGE_INIT_ARTICULO, pY+7, GF_TRADUCIR("ARTICULO"), PAGE_WIDTH_ARTICULO, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_PRECIO_INICIO, pY+7, GF_TRADUCIR("PRECIO INICIO"), PAGE_WIDTH_PRECIO_INICIO, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_PRECIO_FINAL, pY+7, GF_TRADUCIR("PRECIO FINAL"), PAGE_WIDTH_PRECIO_FINAL, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_VARIACION, pY+7, GF_TRADUCIR("VAR."), PAGE_WIDTH_VARIACION, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_PRECIO_ULT, pY + 3, GF_TRADUCIR("ULTIMA COMPRA"), PAGE_WIDTH_PRECIO_ULT + PAGE_WIDTH_FECHA_ULT, PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(oPDF, PAGE_INIT_PRECIO_ULT, pY+15, GF_TRADUCIR("IMPORTE"), PAGE_WIDTH_PRECIO_ULT, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_FECHA_ULT, pY+15, GF_TRADUCIR("FECHA"), PAGE_WIDTH_FECHA_ULT, PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(oPDF, PAGE_INIT_PIC, pY+7, GF_TRADUCIR("PIC"), PAGE_WIDTH_PIC, PDF_ALIGN_CENTER)
	call GF_setFontColor("000000")
End Function
'-------------------------------------------------------------------------------------------------------------
Function dibujarTitulosDetalleBox(pY)
	Call GF_squareBox(oPDF, PAGE_INIT, pY, PAGE_WIDTH, 25, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_verticalLine(oPDF, PAGE_WIDTH_ARTICULO+PAGE_INIT_ARTICULO, pY, 25)
	Call GF_verticalLine(oPDF, PAGE_WIDTH_PRECIO_INICIO+PAGE_INIT_PRECIO_INICIO, pY, 25)
	Call GF_verticalLine(oPDF, PAGE_WIDTH_PRECIO_FINAL+PAGE_INIT_PRECIO_FINAL, pY, 25)
	Call GF_verticalLine(oPDF, PAGE_WIDTH_VARIACION+PAGE_INIT_VARIACION, pY, 25)
	Call GF_verticalLine(oPDF, PAGE_WIDTH_PRECIO_ULT+PAGE_INIT_PRECIO_ULT+PAGE_WIDTH_FECHA_ULT, pY, 25)
	Call GF_verticalLine(oPDF, PAGE_WIDTH_PRECIO_ULT+PAGE_INIT_PRECIO_ULT, pY+12, 12)
	Call GF_horizontalLine(oPDF, PAGE_INIT_PRECIO_ULT, pY+12, PAGE_WIDTH_FECHA_ULT + PAGE_WIDTH_PRECIO_ULT)	
	Call GF_verticalLine(oPDF, PAGE_WIDTH_PIC+PAGE_INIT_PIC, pY, 25)
End Function
'-------------------------------------------------------------------------------------------------------------
Function nuevaPagina()
	Call GF_newPage(oPDF)
	currentY = PAGE_TOP_INIT
	nroPagina = nroPagina + 1
	call drawTitle()
	call dibujarTitulosDetalle(currentY)
	currentY = PAGE_TOP_INIT + PAGE_SPACE_DETAIL
end Function
'-------------------------------------------------------------------------------------------------------------
Function armadoDetalleMovimientos()
	dim rs,totalReg,i,isInicio,aux_Y,sp_parameter
	sp_parameter = g_idDivision &"||"& g_fechaDesde &"||"& g_fechaHasta &"||"& g_idArticulo &"||"& g_idCategoria &"||"& g_cdmoneda &"||1||0$$totalRegistros"		
	Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLARTICULOSPRECIOS_GET_FULL_BY_PARAMETERS", sp_parameter)
	totalRegistros = sp_ret("totalRegistros")
	call dibujarTitulosDetalle(currentY)	 
	currentY = currentY + 30
	flagNuevaPagina = false
	if not rs.Eof then
		while not rs.eof
			if (flagNuevaPagina) then
				Call nuevaPagina()				
				aux_Y = currentY
				flagNuevaPagina = false
			end if			
			call dibujarMovimiento(currentY, rs("IDARTICULO"),rs("DSARTICULO"),rs("MMTOPRECIOINICIO"),rs("IMPORTEINICIO"),rs("MMTOPRECIOFINAL"),rs("IMPORTEFINAL"),rs("VARIACION"),rs("MMTOULTIMACOMPRA"),rs("VLUMONEDAULTCOMPRA"),rs("IDPIC"))						
			currentY = currentY + SEPARATION
			if (currentY > PAGE_HEIGHT_SIZE) then flagNuevaPagina = true			
		rs.MoveNext()
		wend
	else
		Call GF_writeTextAlign(oPDF,PAGE_INIT , currentY, "No se encontraron resultados", PAGE_WIDTH, PDF_ALIGN_CENTER)
	end if
end Function
'-------------------------------------------------------------------------------------------------------------
Function dibujarMovimiento(pY,pIdArticulo,pDsArticulo,pMmtoPrecioIni,pPesosIni,pMmtoPrecioFin,pPesosFin,pVariacion,pMmtoUltCompIni,pPesosUltCompIni,pPic)

	call GF_setFont(oPDF,"COURIER",8,0)
	contador = contador + 1
	if contador mod 2 then
		color = "#FFFFFF"
	else
		color = "#E0E0F8"
	end if
	Call GF_squareBox(oPDF, PAGE_INIT, pY, PAGE_WIDTH, 14, 0, color, color, 1, PDF_SQUARE_NORMAL)
	auxArticulo = pIdArticulo &"-"& pDsArticulo	
	if (Len(auxArticulo) > 39) then auxArticulo = left(auxArticulo,37) & ".."	
	Call GF_writeTextAlign(oPDF,PAGE_INIT_ARTICULO, pY, auxArticulo, PAGE_WIDTH_ARTICULO, PDF_ALIGN_LEFT)	
	if (not isNull(pPesosIni)) then 
		Call GF_writeTextAlign(oPDF,PAGE_INIT_PRECIO_INICIO , pY, getSimboloMoneda(g_cdmoneda) & " " & GF_EDIT_DECIMALS(pPesosIni,2), PAGE_WIDTH_PRECIO_INICIO, PDF_ALIGN_RIGHT)	
	else
		Call GF_writeTextAlign(oPDF,PAGE_INIT_PRECIO_INICIO , pY, "Sin Datos", PAGE_WIDTH_PRECIO_INICIO, PDF_ALIGN_CENTER)
	end if
	if (not isNull(pPesosFin)) then 
		Call GF_writeTextAlign(oPDF,PAGE_INIT_PRECIO_FINAL , pY, getSimboloMoneda(g_cdmoneda) & " " & GF_EDIT_DECIMALS(pPesosFin,2), PAGE_WIDTH_PRECIO_FINAL, PDF_ALIGN_RIGHT)
	else
		Call GF_writeTextAlign(oPDF,PAGE_INIT_PRECIO_FINAL , pY, "Sin Datos", PAGE_WIDTH_PRECIO_FINAL, PDF_ALIGN_CENTER)
	end if
	auxVariacion = "-"
	if ((not isNull(pPesosIni))and(not isNull(pPesosFin))) then	
		if ((Cdbl(pPesosFin) > 0)and(Cdbl(pPesosIni) > 0)) Then auxVariacion = pVariacion & "%"
	end if	
	Call GF_writeTextAlign(oPDF,PAGE_INIT_VARIACION, pY, auxVariacion , PAGE_WIDTH_VARIACION, PDF_ALIGN_RIGHT)
	if (not isNull(pPesosUltCompIni)) then 
		Call GF_writeTextAlign(oPDF,PAGE_INIT_PRECIO_ULT, pY, getSimboloMoneda(g_cdmoneda) & " " & GF_EDIT_DECIMALS(pPesosUltCompIni,2), PAGE_WIDTH_PRECIO_ULT, PDF_ALIGN_RIGHT)		
	else
		Call GF_writeTextAlign(oPDF,PAGE_INIT_PRECIO_ULT, pY, "Sin Datos", PAGE_WIDTH_PRECIO_ULT, PDF_ALIGN_CENTER)		
	end if		
	if (not isNull(pMmtoUltCompIni)) then
		Call GF_writeTextAlign(oPDF,PAGE_INIT_FECHA_ULT, pY, GF_FN2DTE(Left(pMmtoUltCompIni,8)), PAGE_WIDTH_FECHA_ULT, PDF_ALIGN_RIGHT)
	else
		Call GF_writeTextAlign(oPDF,PAGE_INIT_FECHA_ULT, pY, "Sin Datos", PAGE_WIDTH_FECHA_ULT, PDF_ALIGN_CENTER)
	end if	
	if (pPic <> "") then Call GF_writeTextAlign(oPDF,PAGE_INIT_PIC, pY, pPic, PAGE_WIDTH_PIC, PDF_ALIGN_CENTER)
	
	call GF_setFont(oPDF,"COURIER",5,0)
	if (pMmtoPrecioIni <> "") then Call GF_writeTextAlign(oPDF,PAGE_INIT_PRECIO_INICIO , pY + 8, GF_FN2DTE(Left(pMmtoPrecioIni,8)), PAGE_WIDTH_PRECIO_INICIO, PDF_ALIGN_LEFT)
	if (pMmtoPrecioFin <> "") then Call GF_writeTextAlign(oPDF,PAGE_INIT_PRECIO_FINAL, pY + 8, GF_FN2DTE(Left(pMmtoPrecioFin,8)), PAGE_WIDTH_PRECIO_FINAL, PDF_ALIGN_LEFT)
	
end Function
'------------------------------------------------------------------------------------------
Function drawFilter()
	dim yInicio, ySeparation,auxCategoria,auxArticulo,dsArticulo
	yInicio = PAGE_TOP_INIT
	ySeparation = 0	
	call GF_setFont(oPDF,"COURIER",8,0)
	Call GF_writeTextAlign(oPDF, PAGE_INIT, yInicio, GF_TRADUCIR("Fecha Desde.........:"), PAGE_WIDTH_TIT_FILTROS, PDF_ALIGN_LEFT)		
	Call GF_writeTextAlign(oPDF, PAGE_INIT_FILTROS, yInicio, GF_FN2DTE(g_fechaDesde), PAGE_WIDTH_FILTROS, PDF_ALIGN_LEFT)
	ySeparation = yInicio + SEPARATION
	Call GF_writeTextAlign(oPDF, PAGE_INIT, ySeparation, GF_TRADUCIR("Fecha Hasta.........:"), PAGE_WIDTH_TIT_FILTROS, PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_FILTROS, ySeparation, GF_FN2DTE(g_fechaHasta), PAGE_WIDTH_FILTROS, PDF_ALIGN_LEFT)
	ySeparation = ySeparation + SEPARATION
	Call GF_writeTextAlign(oPDF, PAGE_INIT, ySeparation, GF_TRADUCIR("Division............:"), PAGE_WIDTH_TIT_FILTROS, PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_FILTROS, ySeparation, g_idDivision &"-"& UCASE(getDivisionDS(g_idDivision)), PAGE_WIDTH_FILTROS, PDF_ALIGN_LEFT)
	ySeparation = ySeparation + SEPARATION
	auxCategoria = "Todas"
	if (g_idCategoria <> 0) Then auxCategoria = g_idCategoria &"-"& getCategoriaDS(g_idCategoria)
	Call GF_writeTextAlign(oPDF, PAGE_INIT, ySeparation, GF_TRADUCIR("Categoria...........:"), PAGE_WIDTH_TIT_FILTROS, PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_FILTROS, ySeparation, auxCategoria, PAGE_WIDTH_FILTROS, PDF_ALIGN_LEFT)
	ySeparation = ySeparation + SEPARATION
	auxArticulo = "Todos"
	if (g_idArticulo <> 0) Then 
		Call getArticuloFull(g_idArticulo,dsArticulo,"")
		auxArticulo = g_idArticulo &"-"& dsArticulo
	end if		 
	Call GF_writeTextAlign(oPDF, PAGE_INIT, ySeparation, GF_TRADUCIR("Articulo............:"), PAGE_WIDTH_TIT_FILTROS, PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_FILTROS, ySeparation, auxArticulo, PAGE_WIDTH_FILTROS, PDF_ALIGN_LEFT)
	ySeparation = ySeparation + SEPARATION
	Call GF_writeTextAlign(oPDF, PAGE_INIT, ySeparation, GF_TRADUCIR("Moneda..............:"), PAGE_WIDTH_TIT_FILTROS, PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF, PAGE_INIT_FILTROS, ySeparation, getSimboloMonedaLetras(g_cdmoneda), PAGE_WIDTH_FILTROS, PDF_ALIGN_LEFT)
	ySeparation = ySeparation + SEPARATION
	drawFilter = ySeparation
end Function
'------------------------------------------------------------------------------------------
Function drawTitle()
	Call GF_squareBox(oPDF, 2, 2, PAGE_WIDTH_CABECERA + 10, 833, 0, "", "#0B3B0B", 2, PDF_SQUARE_ROUND)
	Call GF_writeImage(oPDF, Server.MapPath("images\kogge64.gif"), PAGE_INIT, 10, 48, 48, 0)
	call GF_setFont(oPDF,"ARIAL",16,0)
	Call GF_writeTextAlign(oPDF,2, 25, GF_TRADUCIR("REPORTE VARIACIÓN DE PRECIO"), PAGE_WIDTH_CABECERA + 10, PDF_ALIGN_CENTER)
	Call GF_horizontalLine(oPDF,2,65,590)
	call GF_setFont(oPDF,"ARIAL",8,0)
	Call GF_writeTextAlign(oPDF, 10 , 840, "Pagina  " & nroPagina, PAGE_WIDTH_CABECERA , PDF_ALIGN_RIGHT)
	Call GF_setFont(oPDF,"COURIER",8,0)
	GP_CONFIGURARMOMENTOS
	Call GF_writeTextAlign(oPDF,5,5,GF_FN2DTE(session("MmtoSistema")), PAGE_WIDTH_CABECERA , PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,5,5+SEPARATION,session("Usuario"), PAGE_WIDTH_CABECERA , PDF_ALIGN_RIGHT)
end Function
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

Set oPDF = GF_createPDF("PDFTemp")
Call GF_setPDFMODE(PDF_STREAM_MODE)
call armadoPDF()
Call GF_closePDF(oPDF)

%>



	