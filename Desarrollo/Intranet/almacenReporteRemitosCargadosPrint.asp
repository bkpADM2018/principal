<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosPDF.asp"-->
<%
dim rs, conn, strSQL, oPDF, rsMovimientos, contador, color, nroPagina, currentY, currentAuxY
dim RPT_Almacen, RPT_FechaDesde, RPT_FechaHasta, RPT_idArticulo, RPT_cdArticulo, RPT_dsArticulo, RPT_idCategoria, RPT_idProveedor, RPT_dsProveedor
dim SEPARATION, MARGIN, PAGE_HEIGHT_SIZE, PAGE_TOP_INIT, RPT_cdCargo, RPT_idCargo, RPT_dsCargo, RPT_verDetalle

SEPARATION = 10
MARGIN = 0
PAGE_HEIGHT_SIZE = 820
PAGE_TOP_INIT = 82
nroPagina = 1

RPT_Almacen = GF_Parametros7("idAlmacen", "", 6)
RPT_FechaDesde = GF_Parametros7("issuedate", "", 6)
RPT_FechaHasta = GF_Parametros7("closingdate", "", 6)
RPT_idCategoria = GF_Parametros7("idCategoria", 0, 6)
RPT_idArticulo = GF_Parametros7("idArticulo", 0, 6)
RPT_idProveedor = GF_Parametros7("idProveedor", 0, 6)
RPT_dsProveedor = GF_Parametros7("dsProveedor", 0, 6)
RPT_cdCargo = GF_Parametros7("cdCargo", "", 6)
if (RPT_cdCargo <> "") then Call GF_MGKS("SG", RPT_cdCargo, RPT_idCargo, RPT_dsCargo)
RPT_verDetalle = GF_Parametros7("verDetalle", "", 6)



Set oPDF = GF_createPDF("PDFTemp")
'Call PDFGirarHoja(90)
Call GF_setPDFMODE(PDF_STREAM_MODE)
call armadoPDF()
Call GF_closePDF(oPDF)
'-----------------------------------------------------------------------------------------------
Function filtrarMovimientos(ByRef myWhere, idAlmacen, fechaDesde, fechaHasta, idArticulo, idProveedor, cdCargo, idCategoria)
		
	'Filtro
	if (idAlmacen <> 0)			then Call mkWhere(myWhere, "A.IDALMACEN", idAlmacen, "=", 1)	
	if (fechaDesde <> "")		then Call mkWhere(myWhere, "A.FECHA", GF_DTE2FN(fechaDesde), ">=", 3)
	if (fechaHasta <> "")		then Call mkWhere(myWhere, "A.FECHA", GF_DTE2FN(fechaHasta), "<=", 3)
	if (idArticulo <> 0)		then Call mkWhere(myWhere, "B.IDARTICULO", idArticulo, "=", 1)
	if (idProveedor <> 0)		then Call mkWhere(myWhere, "A.IDPROVEEDOR", idProveedor, "=", 1)
	if (cdCargo <> "")			then Call mkWhere(myWhere, "A.CDUSUARIO", cdCargo, "=", 3)
	if (idCategoria <> 0)		then Call mkWhere(myWhere, "D.IDCATEGORIA", idCategoria, "=", 1)
	filtrarMovimientos = myWhere	
End Function

'-----------------------------------------------------------------------------------------------
Function obtenerListaMovimientos(idAlmacen, fechaDesde, fechaHasta, idArticulo, idProveedor, cdCargo, idCategoria)
	Dim strSQL, rs, myWhere, conn
	Call filtrarMovimientos(myWhere, idAlmacen, fechaDesde, fechaHasta, idArticulo, idProveedor, cdCargo, idCategoria)	
	strSQL = "Select A.FECHA, A.IDREMITO, A.NROREMITO, A.IDPROVEEDOR from TBLREMCABECERA A inner join TBLREMDETALLE B on A.IDREMITO=B.IDREMITO" 
	strSQL = strSQL & " inner join TBLARTICULOS C on B.IDARTICULO = C. IDARTICULO inner join TBLARTCATEGORIAS D on C.IDCATEGORIA=D.IDCATEGORIA " & myWhere
	strSQL = strSQL & " group by A.FECHA, A.IDREMITO, A.NROREMITO, A.IDPROVEEDOR order by A.FECHA, A.IDREMITO asc"
	'response.Write strSQL
	'response.End
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL) 
	Set obtenerListaMovimientos = rs
End Function
'------------------------------------------------------------------------------------------
Function armadoPDF()
dim totalReg, i
	currentY = armadoCabecera(GF_TRADUCIR("REMITOS CARGADOS"), RPT_Almacen, RPT_FechaDesde, RPT_FechaHasta, RPT_idArticulo, RPT_idProveedor, RPT_idCargo, RPT_idCategoria, RPT_verDetalle)
	set rsMovimientos = obtenerListaMovimientos(RPT_Almacen, RPT_FechaDesde, RPT_FechaHasta, RPT_idArticulo, RPT_idProveedor, RPT_cdCargo, RPT_idCategoria)
	totalReg = rsMovimientos.RecordCount
	if rsMovimientos.eof then
		Call GF_horizontalLine(oPDF,20,currentY ,558)
		Call GF_writeTextAlign(oPDF,20, currentY + SEPARATION/2, ucase("No se han encontrado resultados"), 590, PDF_ALIGN_CENTER)
	else
		currentAuxY = currentY
		currentY = currentY + SEPARATION + 4	
		while not rsMovimientos.eof
				i = i + 1
				call dibujarRemito(currentY, rsMovimientos("FECHA"), rsMovimientos("IDREMITO"), rsMovimientos("NROREMITO"), rsMovimientos("IDPROVEEDOR"))
				currentY = currentY + (SEPARATION)
				if (RPT_verDetalle = "SI") then currentY = dibujarDetalleRemito(currentY, rsMovimientos("IDREMITO"))
				if currentY > PAGE_HEIGHT_SIZE and i <= totalReg then	
					call dibujarTitulosDetalle(currentAuxY)
					nuevaPagina()	
				end if
			rsMovimientos.moveNext
		wend
		Call GF_horizontalLine(oPDF,20,currentY + SEPARATION/2 ,558)
		Call GF_writeTextAlign(oPDF, 20, currentY + SEPARATION/2 , GF_TRADUCIR("FIN DEL REPORTE"), 550, PDF_ALIGN_CENTER)		
		call dibujarTitulosDetalle(currentAuxY)
	end if
end Function
'------------------------------------------------------------------------------------------
function nuevaPagina()
	Call GF_newPage(oPDF)
	currentY = PAGE_TOP_INIT + (SEPARATION) + 4		
	nroPagina = nroPagina + 1
	call dibujarTitulo(GF_TRADUCIR("REMITOS CARGADOS"))
	currentAuxY = PAGE_TOP_INIT 
end function
'------------------------------------------------------------------------------------------
function dibujarTitulosDetalle(pY)
	Call GF_squareBox(oPDF, 20, pY, 100, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 120, pY, 80, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 200, pY, 100, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)	
	Call GF_squareBox(oPDF, 300, pY, 280, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)	
	call GF_setFont(oPDF,"ARIAL",8,8)
	call GF_setFontColor("FFFFFF")
	Call GF_writeTextAlign(oPDF, 20, pY+2	, GF_TRADUCIR("FECHA")		, 100	, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, 125, pY+2	, GF_TRADUCIR("ID REMITO")	, 80	, PDF_ALIGN_LEFT)	
	Call GF_writeTextAlign(oPDF, 205, pY+2	, GF_TRADUCIR("Nº REMITO")	, 100	, PDF_ALIGN_LEFT)	
	Call GF_writeTextAlign(oPDF, 305, pY+2	, GF_TRADUCIR("PROVEEDOR")	, 280	, PDF_ALIGN_LEFT)		
	call GF_setFontColor("000000")
end function
'------------------------------------------------------------------------------------------
function dibujarRemito(pY, fecha, idRemito, nroRemito, idProveedor)
	Dim strSQL, rsEmpresa, conn, auxDsProveedor
	call GF_setFont(oPDF,"COURIER",8,0)
	contador = contador + 1
	if ((contador mod 2) and (RPT_verDetalle <> "SI")) then 
		color = "#FFFFFF"			
	else
		color = "#dcf7dc"
	end if	
	Call GF_squareBox(oPDF, 20, pY, 560, 10, 0, color, color, 1, PDF_SQUARE_NORMAL)
	Call GF_writeTextAlign(oPDF,20, pY, GF_FN2DTE(fecha), 100, PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(oPDF,127, pY, idRemito, 70, PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF,207, pY, nroRemito, 90, PDF_ALIGN_LEFT)
	if (idProveedor <> 0 ) then auxDsProveedor = get_proveedor(idProveedor)
	if (auxDsProveedor <> "") then Call GF_writeTextAlign(oPDF,307, pY, left(ucase(auxDsProveedor),45), 280, PDF_ALIGN_LEFT)	
end function
'------------------------------------------------------------------------------------------
function dibujarTitulo(pTitulo)
	Call GF_squareBox(oPDF, 2, 2, 590, 833, 0, "", "#0B3B0B", 2, PDF_SQUARE_ROUND)
	Call GF_writeImage(oPDF, Server.MapPath("images\kogge64.gif"), 20, 10, 48, 48, 0)
	call GF_setFont(oPDF,"ARIAL",16,8)
	Call GF_writeTextAlign(oPDF,2, 25, GF_TRADUCIR(pTitulo), 590, PDF_ALIGN_CENTER)
	Call GF_horizontalLine(oPDF,2,65,590)
	call GF_setFont(oPDF,"ARIAL",8,0)
	Call GF_writeTextAlign(oPDF, 10 , 840, "Pagina  " & nroPagina		 , 580 , PDF_ALIGN_RIGHT)		

	Call GF_setFont(oPDF,"COURIER",8,0)
	GP_CONFIGURARMOMENTOS
	Call GF_writeTextAlign(oPDF,5,5,GF_FN2DTE(session("MmtoSistema")), 580 , PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,5,5+SEPARATION,session("Usuario"), 580 , PDF_ALIGN_RIGHT)
end function
'------------------------------------------------------------------------------------------
Function armadoCabecera(p_Titulo, idAlmacen, fechaDesde, fechaHasta, idArticulo, idProveedor, idCargo, idCategoria, verDetalle)
	dim yInicio, ySeparation, auxCargo, auxCatDS
	yInicio = PAGE_TOP_INIT
	ySeparation = 0
	call dibujarTitulo(p_Titulo)
	call GF_setFont(oPDF,"COURIER",8,0)
	Call GF_writeTextAlign(oPDF, 20, yInicio, GF_TRADUCIR("Almacen.........:")		, 50, PDF_ALIGN_LEFT)
	auxAlmacen = GF_Traducir("Todos")
	if (idAlmacen <> 0) then 
		set rsAlmacen = obtenerListaAlmacenes(idAlmacen)
		if not rsAlmacen.eof then auxAlmacen = rsAlmacen("CDALMACEN") & " - " & rsAlmacen("DSALMACEN")
	end if
	Call GF_writeTextAlign(oPDF, 110, yInicio, ucase(auxAlmacen), 200, PDF_ALIGN_LEFT)
	ySeparation = ySeparation + SEPARATION

	auxFechaDesde = GF_Traducir("Todas") 
	Call GF_writeTextAlign(oPDF, 20, yInicio + ySeparation, GF_TRADUCIR("Fecha Desde.....:")	, 50, PDF_ALIGN_LEFT)
	if (fechaDesde <> "") then auxFechaDesde = fechaDesde 
	Call GF_writeTextAlign(oPDF, 110, yInicio + ySeparation, ucase(auxFechaDesde), 200, PDF_ALIGN_LEFT)
	ySeparation = ySeparation + SEPARATION

	auxFechaHasta = GF_Traducir("Todas") 
	Call GF_writeTextAlign(oPDF, 20, yInicio + ySeparation, GF_TRADUCIR("Fecha Hasta.....:")	, 50, PDF_ALIGN_LEFT)
	if (fechaHasta <> "") then auxFechaHasta = fechaHasta  
	Call GF_writeTextAlign(oPDF, 110, yInicio + ySeparation, ucase(auxFechaHasta), 200, PDF_ALIGN_LEFT)
	ySeparation = ySeparation + SEPARATION
	
	Call GF_writeTextAlign(oPDF, 110, yInicio + ySeparation, ucase(auxObra), 200, PDF_ALIGN_LEFT)		
	ySeparation = ySeparation + SEPARATION
		
	auxArt = GF_Traducir("Todos")  	 
	Call GF_writeTextAlign(oPDF, 20, yInicio + ySeparation, GF_TRADUCIR("Articulo........:")		, 50, PDF_ALIGN_LEFT)
	if (idArticulo <> 0) then 	
		call getArticuloFull(idArticulo, auxArtDS, auxArtAbrev)
		auxArt = auxArtDS & " - " & auxArtAbrev
	end if
	Call GF_writeTextAlign(oPDF, 110, yInicio + ySeparation, ucase(auxArt), 200, PDF_ALIGN_LEFT)
	ySeparation = ySeparation + SEPARATION

	auxCatDS = GF_Traducir("Todas")  	 
	Call GF_writeTextAlign(oPDF, 20, yInicio + ySeparation, GF_TRADUCIR("Categoria.......:")		, 50, PDF_ALIGN_LEFT)
	if (idCategoria <> 0) then 	auxCatDS = getCategoriaFull(idCategoria)			
	Call GF_writeTextAlign(oPDF, 110, yInicio + ySeparation, ucase(auxCatDS), 200, PDF_ALIGN_LEFT)
	ySeparation = ySeparation + SEPARATION
	
	auxProveedor = GF_Traducir("Todos")  	  
	Call GF_writeTextAlign(oPDF, 20, yInicio + ySeparation, GF_TRADUCIR("Proveedor.......:")	, 50, PDF_ALIGN_LEFT)	
	if (idProveedor <> 0)	then
		auxProveedor = get_proveedor(idProveedor)
	end if
	Call GF_writeTextAlign(oPDF, 110, yInicio + ySeparation, ucase(auxProveedor), 200, PDF_ALIGN_LEFT)
	ySeparation = ySeparation + SEPARATION
	
	auxCargo = GF_Traducir("Todos")  	  
	Call GF_writeTextAlign(oPDF, 20, yInicio + ySeparation, GF_TRADUCIR("Cargado por.....:")	, 50, PDF_ALIGN_LEFT)
	if (idCargo <> "")	then 	
		auxCargo = RPT_cdCargo & " - " & RPT_dsCargo
	end if
	Call GF_writeTextAlign(oPDF, 110, yInicio + ySeparation, ucase(auxCargo), 200, PDF_ALIGN_LEFT)
	ySeparation = ySeparation + SEPARATION
	
	Call GF_writeTextAlign(oPDF, 20, yInicio + ySeparation, GF_TRADUCIR("Ver Detalle.....:")	, 50, PDF_ALIGN_LEFT)	
	if (verDetalle = "") then 	
		Call GF_writeTextAlign(oPDF, 110, yInicio + ySeparation, GF_TRADUCIR("NO"), 200, PDF_ALIGN_LEFT)
	else
		Call GF_writeTextAlign(oPDF, 110, yInicio + ySeparation, GF_TRADUCIR("SI"), 200, PDF_ALIGN_LEFT)
	end if
	ySeparation = ySeparation + SEPARATION

	yFinal = yInicio + (ySeparation + SEPARATION)

	armadoCabecera = yFinal 
	
end Function
'------------------------------------------------------------------------------------------	
Function dibujarDetalleRemito(currentY, nroRemito)
	Dim strSQL, rsREMDetalle, conn, contador, totalRegDetalle, txt
	
	strSQL = "Select * from TBLREMDETALLE where IDREMITO = " & nroRemito & " order by IDARTICULO asc"
	Call executeQueryDB(DBSITE_SQL_INTRA, rsREMDetalle, "OPEN", strSQL)
	totalRegDetalle = rsREMDetalle.RecordCount
	if (currentY > (PAGE_HEIGHT_SIZE - 20)) then
		call dibujarTitulosDetalle(currentAuxY)
		nuevaPagina()
	end if
	if (not rsREMDetalle.eof) then
		Call dibujarTitularBoxDetalle(currentY)
		currentY = currentY + SEPARATION + 1
		while (not rsREMDetalle.eof)
			contador = contador + 1
			Call getArticuloFull(rsREMDetalle("IDARTICULO"), pDS, pAbrev)
			call GF_setFont(oPDF,"COURIER",6,0)
			Call GF_squareBox(oPDF, 100, currentY, 380, 10, 0, "#F5F6CE", "#F5F6CE", 1, PDF_SQUARE_NORMAL)
			Call GF_writeTextAlign(oPDF,104, currentY+2, rsREMDetalle("IDARTICULO"), 70, PDF_ALIGN_LEFT)	
			Call GF_writeTextAlign(oPDF,174, currentY+2, pDS, 260, PDF_ALIGN_LEFT)	
			Call GF_writeTextAlign(oPDF,430, currentY+2, rsREMDetalle("CANTIDAD") & " " & pAbrev, 45, PDF_ALIGN_RIGHT)
			
			currentY = currentY + SEPARATION
			
			if ((currentY > PAGE_HEIGHT_SIZE) and (contador <= totalRegDetalle)) then
				call dibujarTitulosDetalle(currentAuxY)
				nuevaPagina()
				Call dibujarTitularBoxDetalle((currentAuxY + SEPARATION + 5))
				currentY = currentY + SEPARATION + 2
			end if
			
			rsREMDetalle.moveNext
		wend
		currentY = currentY + 5
	end if
	dibujarDetalleRemito = currentY
end Function
'------------------------------------------------------------------------------------------	
Function dibujarTitularBoxDetalle(pY)
	Call GF_squareBox(oPDF, 100, pY, 70, 10, 0, "#F2F5A9", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 170, pY, 260, 10, 0, "#F2F5A9", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 430, pY, 50, 10, 0, "#F2F5A9", "#000000", 1, PDF_SQUARE_NORMAL)
	call GF_setFont(oPDF,"ARIAL",6,8)
	call GF_setFontColor("000000")
	Call GF_writeTextAlign(oPDF, 102, pY+2	, GF_TRADUCIR("CODIGO"), 70	, PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF, 172, pY+2	, GF_TRADUCIR("DESCRIPCIÓN"), 260	, PDF_ALIGN_LEFT)	
	Call GF_writeTextAlign(oPDF, 430, pY+2	, GF_TRADUCIR("CANTIDAD")	, 50	, PDF_ALIGN_CENTER)		
	call GF_setFontColor("000000")
end Function
'------------------------------------------------------------------------------------------	
Function get_proveedor(id)
	dim aux
	strSQL = "Select nomemp from met001a where nroemp = " & id
	Call executeQueryDB(DBSITE_SQL_MAGIC, rsEmpresa, "OPEN", strSQL)
	if (not rsEmpresa.eof) then aux = rsEmpresa("nomemp")
	get_proveedor = aux
end Function
'------------------------------------------------------------------------------------------	
%>