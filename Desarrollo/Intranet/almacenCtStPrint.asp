<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosPDF.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<%
function nuevaPagina()
	Call GF_newPage(oPDF)
	currentY = PAGE_TOP_INIT + (SEPARATION) + 4		
	nroPagina = nroPagina + 1
	Call dibujarTitulo("CONTROL DE STOCK")
	currentAuxY = PAGE_TOP_INIT 
end function
'-----------------------------------------------------------------------------------------
function dibujarTitulo(pTitulo)
	Call GF_squareBox(oPDF, 2, 2, 590, 833, 0, "", "#0B3B0B", 2, PDF_SQUARE_ROUND)
	Call GF_writeImage(oPDF, Server.MapPath("images\kogge64.gif"), 20, 10, 48, 48, 0)
	call GF_setFont(oPDF,"ARIAL",16,8)
	Call GF_writeTextAlign(oPDF,2, 25, pTitulo, 590, PDF_ALIGN_CENTER)
	Call GF_horizontalLine(oPDF,2,65,590)
	call GF_setFont(oPDF,"ARIAL",8,0)
	Call GF_writeTextAlign(oPDF, 10 , 840, "Pagina  " & nroPagina		 , 580 , PDF_ALIGN_RIGHT)		
	Call GF_setFont(oPDF,"COURIER",8,0)	
	Call GF_writeTextAlign(oPDF,5,5,GF_FN2DTE(session("MmtoSistema")), 580 , PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,5,15,session("Usuario"), 580 , PDF_ALIGN_RIGHT)
end function
'-----------------------------------------------------------------------------------------
'Graba los stocks del sistema y el precio en el detalle del control para que al cargar el resultado no se tengan en cuenta 
'los movimientos que ocurren entre la ultima impresion del reporte y la carga de los resultados.
Function actualizarStocksSistema(idCtrl, rs)
    Dim strSQL, rsAux
    if Not rs.EoF then
		while (not rs.eof)
	        strSQL= "Update TBLCSTKDETALLE set STOCKSISTEMA=" & rs("STOCKSISTEMA") & " , VLUPESOS = " & rs("VLUPESOS") & " where IDCONTROL=" & idCtrl & " and IDARTICULO=" & rs("IDARTICULO")
	        call executeQueryDb(DBSITE_SQL_INTRA, rsAux, "EXEC", strSQL)
			rs.MoveNext()
		wend		
		rs.MoveFirst()
    end if
End Function
'-----------------------------------------------------------------------------------------
function dibujarCabecera(pY)	
	Call GF_squareBox(oPDF, 20	,pY , 250	, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 270	,pY , 85	, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 355	,pY , 55	, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 410	,pY , 75	, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 485	,pY , 75	, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)	
	call GF_setFont(oPDF,"ARIAL",8,8)
	call GF_setFontColor("FFFFFF")
	Call GF_writeTextAlign(oPDF, 20, pY+1 	, GF_TRADUCIR("ARTICULO"), 250	, PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(oPDF, 270, pY+1	, GF_TRADUCIR("COD. INTERNO"), 85	, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, 355, pY+1	, GF_TRADUCIR("PECIO U."), 55 , PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, 410, pY+1	, GF_TRADUCIR("STOCK SISTEMA"), 75 , PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, 485, pY+1	, GF_TRADUCIR("STOCK FISICO"), 75 , PDF_ALIGN_CENTER)	
	call GF_setFontColor("000000")
end function
'-----------------------------------------------------------------------------------------
Function getDetalle(pY,pIdArticulo,pDsArticulo,pCdInterno,pStock,pUnidad,pStockFisico,pPrecioUPeso,pflagDatosCargados)
	Dim auxCdInterno,auxStock,auxArticulo,auxPrecioUPeso
	auxArticulo = Trim(pIdArticulo)&" - "&Trim(pDsArticulo)
	if (len(auxArticulo) > 53) then auxArticulo = left(auxArticulo,50) & "..."
	Call GF_writeTextAlign(oPDF, 20, pY+1 	, auxArticulo,250, PDF_ALIGN_LEFT)
	auxCdInterno = pCdInterno
	if isNull(pCdInterno)then auxCdInterno = ""	
	Call GF_writeTextAlign(oPDF, 270, pY+1	, auxCdInterno, 85	, PDF_ALIGN_LEFT)	
	auxPrecioUPeso = 0
	if Not isNull(pPrecioUPeso)then auxPrecioUPeso = pPrecioUPeso	
	Call GF_writeTextAlign(oPDF, 355, pY+1	, "$ " & GF_EDIT_DECIMALS(auxPrecioUPeso,2), 55	, PDF_ALIGN_RIGHT)	
	auxStock = pStock
	if isNull(pStock)then auxStock = 0
	Call GF_writeTextAlign(oPDF, 410, pY+1	, auxStock&" "&pUnidad, 73 , PDF_ALIGN_RIGHT)
	if(pflagDatosCargados)then
		Call GF_writeTextAlign(oPDF, 485, pY+1	, pStockFisico&" "&pUnidad, 73 , PDF_ALIGN_RIGHT)	
	else	
		Call GF_squareBox(oPDF, 495	,pY+1 , 45	, 10, 0, color, "#000000", 1, PDF_SQUARE_RIGHT)		
		Call GF_writeTextAlign(oPDF, 510, pY+1	, pUnidad, 50 , PDF_ALIGN_RIGHT)	
	end if	
	getDetalle = pY
End function
'-----------------------------------------------------------------------------------------
Function dibujarFiltros(pY, pIdControl, pcdResponsable, pIdAlmacen,flagDatosCargados,pCantArt, pPrecioMin, pArtConStock,pTipoReporte,pPrecioMax)
	Dim auxCantArt, auxPy, auxArtConStock, auxTipoReporte, auxPrecioMinimo, auxPrecioMaximo,auxAlmacen,rsAlmacen
	Call GF_writeText(oPDF, 20, pY, GF_TRADUCIR("ID Control") & ".........: ", 0)
	Call GF_writeText(oPDF, 20, pY + 8, GF_TRADUCIR("Responsable") & "........: ", 0)
	Call GF_writeText(oPDF, 20, pY + 16, GF_TRADUCIR("Almacen") & "............: ", 0)
	Call GF_writeText(oPDF, 120, pY, pIdControl, 0)
	Call GF_writeText(oPDF, 120, pY + 8, getUserDescription(pcdResponsable), 0)	
	auxAlmacen = GF_Traducir("Todos")
	if (pIdAlmacen <> 0) then 
		set rsAlmacen = obtenerListaAlmacenes(pIdAlmacen)
		if not rsAlmacen.eof then auxAlmacen = rsAlmacen("CDALMACEN") & " - " & rsAlmacen("DSALMACEN")
	end if	
	Call GF_writeText(oPDF, 120, pY + 16, ucase(auxAlmacen), 0)
	Call GF_writeText(oPDF, 20, pY + 24, GF_TRADUCIR("Seleccion") & "..........: ", 0)	
	auxTipoReporte = "Automatico"
	if(pTipoReporte = CTST_SELECCION_MANUAL)then auxTipoReporte = "Manual"
	Call GF_writeText(oPDF, 120, pY + 24, auxTipoReporte, 0)
	auxPy = pY + 32
	if(flagDatosCargados)then
		Call GF_writeText(oPDF, 20, auxPy, GF_TRADUCIR("Fecha carga") & "........:", 0)
		Call GF_writeText(oPDF, 120, auxPy, GF_FN2DTE(Left(VS_momento,8)), 0)
		auxPy = auxPy + 8		
	end if	
	if(pTipoReporte = CTST_SELECCION_AUTOMATICA)then		
		Call GF_writeText(oPDF, 20, auxPy, GF_TRADUCIR("Cantidad articulos") & ".: ", 0)
		Call GF_writeText(oPDF, 120,auxPy, pCantArt, 0)
		auxPy = auxPy + 8
		Call GF_writeText(oPDF, 20, auxPy, GF_TRADUCIR("Precio Minimo") & "......: ", 0)
		auxPrecioMinimo = "-"
		if(Cdbl(pPrecioMin) > 0)then auxPrecioMinimo = "u$s " & Cdbl(pPrecioMin)
		Call GF_writeText(oPDF, 120, auxPy, auxPrecioMinimo, 0)
		auxPy = auxPy + 8
		Call GF_writeText(oPDF, 20, auxPy, GF_TRADUCIR("Precio Maximo") & "......: ", 0)
		auxPrecioMaximo = "-"		
		if(Cdbl(pPrecioMax) > 0)then auxPrecioMaximo = "u$s " & Cdbl(pPrecioMax)
		Call GF_writeText(oPDF, 120, auxPy, auxPrecioMaximo, 0)		
		auxPy = auxPy + 8
		Call GF_writeText(oPDF, 20, auxPy, GF_TRADUCIR("Articulo c/Stock") & "...: ", 0)
		auxArtConStock = "Si"
		if(pArtConStock = CTST_ARTICULO_SIN_STOCK)then auxArtConStock = "No"
		Call GF_writeText(oPDF, 120, auxPy, auxArtConStock, 0)
		auxPy = auxPy + 8
	end if	
	dibujarFiltros = auxPy + 5
End Function
'-----------------------------------------------------------------------------------------
Function DibujarLeyenda(pY)
	Call GF_squareBox(oPDF, 20, pY , 540, 12, 0, "#F2F5A9", "", 0, PDF_SQUARE_NORMAL)
	Call GF_writeText(oPDF, 25, pY + 3, GF_TRADUCIR("Los datos del stock del sistema se obtienen al momento de generar el reporte"), 0)
	DibujarLeyenda = pY + 18
End Function
'-----------------------------------------------------------------------------------------
Function armarPDFGeneral(pIdControl,pIdAlmacen,pcdResponsable,flagDatosCargados,pArtConStock,pCantArt,pPrecioMinimo,pTipoReporte,pPrecioMax)
Dim i, cambioPagina,rsCab,myFecha, rs,v_Categoria
v_Categoria = 0
cambioPagina = false
Call dibujarTitulo("CONTROL DE STOCK")
Set rs = leerDetallesCtSt(pIdControl, pIdAlmacen, flagDatosCargados,pTipoReporte)         
currentY = 80
currentY = dibujarFiltros(80,pIdControl,pcdResponsable,pIdAlmacen,flagDatosCargados,pCantArt, pPrecioMinimo, pArtConStock, pTipoReporte,pPrecioMax)
Call dibujarCabecera(currentY)
'Si aún no se grabaron resultados, grabo la foto de los stocks de sistema y del precio unitario.
if (not flagDatosCargados) then Call actualizarStocksSistema(pIdControl, rs)
while not rs.eof	
	if currentY > PAGE_HEIGHT_SIZE  then	
		Call nuevaPagina()					 
		cambioPagina = true
		currentY = 85
		Call dibujarCabecera(currentY)
	end if
	if(rs("CDCATEGORIA") <> v_Categoria)then
		currentY = currentY + SEPARATION + 10
		call GF_setFont(oPDF,"ARIAL",10,8)
		Call GF_writeTextAlign(oPDF, 20, currentY, rs("CDCATEGORIA")&"-"&rs("DSCATEGORIA"), 545, PDF_ALIGN_LEFT)
		currentY = currentY + SEPARATION + 3
		call GF_horizontalLine(oPDF, 20, currentY ,545)		
		call GF_setFont(oPDF,"ARIAL",8,0)
		v_Categoria = rs("CDCATEGORIA")
	end if
	if currentY > PAGE_HEIGHT_SIZE  then	
		Call nuevaPagina()					 
		cambioPagina = true
		currentY = 85
		Call dibujarCabecera(currentY)
	end if
	currentY = currentY + SEPARATION + 3	
	call GF_setFont(oPDF,"ARIAL",8,0)
	currentY = getDetalle(currentY,rs("IDARTICULO"),rs("DSARTICULO"),rs("CDINTERNO"),rs("STOCKSISTEMA"),rs("ABREVIATURA"),rs("STOCKFISICO"),rs("VLUPESOS"),flagDatosCargados)
	rs.moveNext		
wend	
currentY = currentY + SEPARATION + 1
Call GF_horizontalLine(oPDF,20,currentY + SEPARATION/2 ,545)
Call GF_setFont(oPDF,"ARIAL",5,0)
if(not flagDatosCargados)then Call GF_writeTextAlign(oPDF, 190, currentY + SEPARATION, GF_TRADUCIR("Tanto el Stock de Sistema como el Precio Unitario se calcularon al momento de imprimir el reporte."), 550, PDF_ALIGN_CENTER)
Call GF_setFont(oPDF,"ARIAL",8,0)
Call GF_writeTextAlign(oPDF, 20, currentY + SEPARATION/2 + 10, GF_TRADUCIR("FIN DEL REPORTE"), 550, PDF_ALIGN_CENTER)		
end function
'**************************************************************************
'**************************** INICIO PAGINA *******************************
'**************************************************************************
Dim IdControl, flagDatosCargados, rs

SEPARATION = 10
MARGIN = 0
PAGE_HEIGHT_SIZE = 800
PAGE_TOP_INIT = 82
nroPagina = 1

flagDatosCargados = false
IdControl = GF_Parametros7("idControl", 0, 6)
Set rs = leerCabeceraCtSt(IdControl)
Set oPDF = GF_createPDF("")
Call GF_setPDFMODE(PDF_STREAM_MODE)
if( not rs.EoF)then
	if(rs("IDRESULTADO") <> 0)then 
		flagDatosCargados = true
		Call initHeaderValeDB(rs("IDRESULTADO"))
	end if	
	Call armarPDFGeneral(IdControl,rs("IDALMACEN"),rs("CDRESPONSABLE"),flagDatosCargados,rs("ARTCONSTOCK"),rs("CANTIDAD"),rs("PRECIOMINIMO"),rs("TIPO"),rs("PRECIOMAXIMO"))
end if
Call GF_closePDF(oPDF)	


%>