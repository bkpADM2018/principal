<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosPDF.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<%
dim rs, conn, strSQL, oPDF, rsMovimientos, contador, color, nroPagina, currentY, currentAuxY
dim RPT_Almacen, RPT_FechaDesde, RPT_FechaHasta, RPT_idObra, RPT_idArticulo, RPT_dsArticulo, RPT_cdSolicitante, RPT_dsSolicitante, RPT_idBudgetArea, RPT_idBudgetDetalle, RPT_idCategoria
dim SEPARATION, MARGIN, PAGE_HEIGHT_SIZE, PAGE_TOP_INIT, RPT_cdCargo, RPT_idCargo, RPT_dsCargo, RPT_verDetalle, RPT_idVale, RPT_auxIdVale, RPT_idSector

SEPARATION = 10
MARGIN = 0
PAGE_HEIGHT_SIZE = 800
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

Set oPDF = GF_createPDF("PDFTemp")
Call GF_setPDFMODE(PDF_STREAM_MODE)
call armadoPDF()
Call GF_closePDF(oPDF)
'-----------------------------------------------------------------------------------------------
Function filtrarMovimientos(ByRef myWhere, idAlmacen, fechaDesde, fechaHasta, idObra, idArticulo, cdSolicitante, cdCargo, idCategoria, idBudgetArea, idBudgetDetalle, idSector)
		
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
Function obtenerListaMovimientos(idAlmacen, fechaDesde, fechaHasta, idObra, idArticulo, cdSolicitante, cdCargo, idCategoria, idBudgetArea, idBudgetDetalle, idSector)
	Dim strSQL, rs, myWhere, conn
	Call filtrarMovimientos(myWhere, idAlmacen, fechaDesde, fechaHasta, idObra, idArticulo, cdSolicitante, cdCargo, idCategoria, idBudgetArea, idBudgetDetalle, idSector)	
	strSQL = "Select A.FECHA, A.IDVALE, A.CDVALE, A.NRVALE, A.CDSOLICITANTE, A.IDOBRA, A.ESTADO, A.IDBUDGETAREA, A.IDBUDGETDETALLE, A.IDSECTOR, B.IDARTICULO, B.CANTIDAD, D.IDCATEGORIA from TBLVALESCABECERA A inner join TBLVALESDETALLE B on A.IDVALE=B.IDVALE" 
	strSQL = strSQL & " inner join TBLARTICULOS C on B.IDARTICULO = C. IDARTICULO inner join TBLARTCATEGORIAS D on C.IDCATEGORIA=D.IDCATEGORIA " & myWhere
	strSQL = strSQL & " group by A.FECHA, A.IDVALE, A.CDVALE, A.NRVALE, A.CDSOLICITANTE, A.IDOBRA, A.ESTADO, A.IDBUDGETAREA, A.IDBUDGETDETALLE, A.IDSECTOR, B.IDARTICULO, B.CANTIDAD, D.IDCATEGORIA order by A.FECHA, A.IDVALE asc"
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set obtenerListaMovimientos = rs
End Function
'------------------------------------------------------------------------------------------
Function armadoPDF()
dim totalReg, i, cambioPagina
	cambioPagina = false
	currentY = armadoCabecera(GF_TRADUCIR("VALES EMITIDOS"), RPT_Almacen, RPT_FechaDesde, RPT_FechaHasta, RPT_idObra, RPT_idBudgetArea, RPT_idBudgetDetalle, RPT_idArticulo, RPT_cdSolicitante, RPT_idCargo, RPT_idCategoria, RPT_verDetalle, RPT_idSector)
	set rsMovimientos = obtenerListaMovimientos(RPT_Almacen, RPT_FechaDesde, RPT_FechaHasta, RPT_idObra, RPT_idArticulo, RPT_cdSolicitante, RPT_cdCargo, RPT_idCategoria, RPT_idBudgetArea, RPT_idBudgetDetalle, RPT_idSector)
	totalReg = rsMovimientos.RecordCount
	if rsMovimientos.eof then
		Call GF_horizontalLine(oPDF,20,currentY ,558)
		Call GF_writeTextAlign(oPDF,20, currentY + SEPARATION/2, ucase("No se han encontrado resultados"), 590, PDF_ALIGN_CENTER)
	else
		currentAuxY = currentY
		currentY = currentY + SEPARATION + 4	
		while not rsMovimientos.eof
				i = i + 1
				RPT_idVale = rsMovimientos("IdVale")
				if (RPT_idVale <> 0) then
					if ((RPT_idVale <> RPT_auxIdVale) or (cambioPagina)) then
						call dibujarVale(currentY, rsMovimientos("FECHA"), RPT_idVale, rsMovimientos("CdVale"),rsMovimientos("nrVale"), rsMovimientos("CDSolicitante"), rsMovimientos("IDOBRA"), CInt(rsMovimientos("IDBUDGETAREA")), rsMovimientos("IDBUDGETDETALLE"),rsMovimientos("ESTADO"), rsMovimientos("IDSECTOR"))
						currentY = currentY + (SEPARATION)
						if (RPT_verDetalle) then
							Call dibujarTitularBoxDetalle(currentY)
							currentY = currentY + SEPARATION + 1
						end if
						cambioPagina = false
					end if
					if (RPT_verDetalle) then currentY = dibujarDetalleVale(currentY, rsMovimientos("IDARTICULO"), rsMovimientos("CANTIDAD"), rsMovimientos("IDCATEGORIA"))
					if currentY > PAGE_HEIGHT_SIZE and i <= totalReg then	
						call dibujarTitulosDetalle(currentAuxY)
						Call nuevaPagina()
						cambioPagina = true
					end if
					RPT_auxIdVale = RPT_idVale
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
	call dibujarTitulo(GF_TRADUCIR("VALES EMITIDOS"))
	currentAuxY = PAGE_TOP_INIT 
end function
'------------------------------------------------------------------------------------------
function dibujarTitulosDetalle(pY)
	Call GF_squareBox(oPDF, 20	, pY, 53	, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 73	, pY, 65	, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)	
	Call GF_squareBox(oPDF, 138	, pY, 45	, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 183	, pY, 180	, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 363	, pY, 110	, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 473	, pY, 105	, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)	
	call GF_setFont(oPDF,"ARIAL",8,8)
	call GF_setFontColor("FFFFFF")
	Call GF_writeTextAlign(oPDF, 20, pY+2	, GF_TRADUCIR("FECHA")			, 53	, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, 73, pY+2	, GF_TRADUCIR("VALE")			, 65	, PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(oPDF, 138, pY+2	, GF_TRADUCIR("TIPO")			, 45	, PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(oPDF, 183, pY+2	, GF_TRADUCIR("SOLICITANTE")	, 180	, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, 363, pY+2	, GF_TRADUCIR("PTDA. PRES. / SECTOR")	, 110	, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, 473, pY+2	, GF_TRADUCIR("ESTADO")			, 105	, PDF_ALIGN_CENTER)	
	call GF_setFontColor("000000")
end function
'------------------------------------------------------------------------------------------
function dibujarVale(pY, fechaVale, idVale, cdVale, nrVale, cdSolicitante, idObra, idBudgetArea, idBudgetDetalle, estado, idsector)
	Dim auxValeRelacionado, dsPartSector
	
	call GF_setFont(oPDF,"COURIER",8,0)
	contador = contador + 1
	if ((contador mod 2) and (not RPT_verDetalle)) then 
		color = "#FFFFFF"			
	else
		color = "#CECEF6"
	end if	
	Call GF_squareBox(oPDF, 20, pY, 558, 10, 0, color, color, 1, PDF_SQUARE_NORMAL)
	Call GF_writeTextAlign(oPDF,20, pY, GF_FN2DTE(fechaVale), 53, PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(oPDF,75, pY, nrVale, 65, PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF,138, pY, cdVale, 45, PDF_ALIGN_CENTER)
	auxDsSolicitante = getUserDescription(cdSolicitante)
	Call GF_writeTextAlign(oPDF,183, pY, left(ucase(auxDsSolicitante),31), 180, PDF_ALIGN_LEFT)	
	dsPartSector = " "
	if (idObra <> 0) then
		call loadDatosObra(idObra, RPT_cdObra, RPT_dsObra, 0, "", 0, "", 0, "", "", "", "", "")
		if (RPT_cdObra <> "") then 	
			dsPartSector = ucase(RPT_cdObra)
			if (idBudgetArea <> 0) then
				set rsBudget = obtenerListaBudgetObra(idObra, idBudgetArea, idBudgetDetalle)
				if not rsBudget.eof then dsPartSector = dsPartSector & " (" & idBudgetArea & "-" & idBudgetDetalle & ")"
			end if
		end if
	elseif (idsector <> 0) then
		dsPartSector = left(uCase(getSectorDS(idsector)), 23)
	end if
	Call GF_writeTextAlign(oPDF,363, pY, dsPartSector, 110, PDF_ALIGN_LEFT)
	if (estado <> ESTADO_ACTIVO) then
		auxValeRelacionado = getValeRelacionado(idVale)
		Select case estado 
			Case ESTADO_BAJA
				auxEstado = "ANULADO POR "	& auxValeRelacionado
			Case ESTADO_ANULACION
				auxEstado = "ANULA AL "	& auxValeRelacionado
		End Select
		Call GF_writeTextAlign(oPDF,473, pY, auxEstado, 105, PDF_ALIGN_LEFT)	
	end if
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
Function armadoCabecera(p_Titulo, idAlmacen, fechaDesde, fechaHasta, idObra, idBudgetArea, idBudgetDetalle, idArticulo, cdSolicitante, idCargo, idCategoria, verDetalle, idSector)
	dim yInicio, ySeparation, auxCargo, auxCatDS, auxDetalle, auxSector
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
	
	auxObra = GF_Traducir("Todas")  	
	Call GF_writeTextAlign(oPDF, 20, yInicio + ySeparation, GF_TRADUCIR("Ptda. Presup....:")			, 50, PDF_ALIGN_LEFT)
	if (idObra <> 0) then
		Call loadDatosObra(idObra, auxObraCD, auxObraDS, 0, "", 0, "", 0, "", "", "", 0, "")
		auxObra = auxObraCD & " - " & auxObraDS
		if (idBudgetArea <> 0) then
			set rsBudget = obtenerListaBudgetObra(idObra, idBudgetArea, idBudgetDetalle)
			if not rsBudget.eof then auxObra = auxObra & " (Detalle Partida: " & idBudgetArea & " - " & idBudgetDetalle & ")"
		end if
	end if
	Call GF_writeTextAlign(oPDF, 110, yInicio + ySeparation, ucase(auxObra), 200, PDF_ALIGN_LEFT)		
	ySeparation = ySeparation + SEPARATION

	auxSector = GF_Traducir("Todos")
	Call GF_writeTextAlign(oPDF, 20, yInicio + ySeparation, GF_TRADUCIR("Sector..........:")	, 50, PDF_ALIGN_LEFT)	
	if (idSector <> 0) then auxSector = getSectorDS(idSector)
	Call GF_writeTextAlign(oPDF, 110, yInicio + ySeparation, ucase(auxSector), 200, PDF_ALIGN_LEFT)
	ySeparation = ySeparation + SEPARATION

	auxArt = GF_Traducir("Todos")  	 
	Call GF_writeTextAlign(oPDF, 20, yInicio + ySeparation, GF_TRADUCIR("Articulo........:")		, 50, PDF_ALIGN_LEFT)
	if (idArticulo <> 0) then 	
		call getArticuloFull(idArticulo, auxArtDS, auxArtAbrev)
		auxArt = idArticulo & " - " & auxArtDS & "[" & auxArtAbrev & "]"
	end if
	Call GF_writeTextAlign(oPDF, 110, yInicio + ySeparation, ucase(auxArt), 200, PDF_ALIGN_LEFT)
	ySeparation = ySeparation + SEPARATION

	auxCatDS = GF_Traducir("Todas")  	 
	Call GF_writeTextAlign(oPDF, 20, yInicio + ySeparation, GF_TRADUCIR("Categoria.......:")		, 50, PDF_ALIGN_LEFT)
	if (idCategoria <> 0) then 	auxCatDS = getCategoriaFull(idCategoria)			
	Call GF_writeTextAlign(oPDF, 110, yInicio + ySeparation, ucase(auxCatDS), 200, PDF_ALIGN_LEFT)
	ySeparation = ySeparation + SEPARATION
	
	auxSolicitante = GF_Traducir("Todos")  	  
	Call GF_writeTextAlign(oPDF, 20, yInicio + ySeparation, GF_TRADUCIR("Solicitante.....:")	, 50, PDF_ALIGN_LEFT)	
	if (cdSolicitante <> "")	then auxSolicitante = cdSolicitante & " - " & getUserDescription(cdSolicitante)
	Call GF_writeTextAlign(oPDF, 110, yInicio + ySeparation, ucase(auxSolicitante), 200, PDF_ALIGN_LEFT)
	ySeparation = ySeparation + SEPARATION
	
	auxCargo = GF_Traducir("Todos")  	  
	Call GF_writeTextAlign(oPDF, 20, yInicio + ySeparation, GF_TRADUCIR("Cargado por.....:")	, 50, PDF_ALIGN_LEFT)
	if (idCargo <> "")	then auxCargo = RPT_cdCargo & " - " & RPT_dsCargo
	Call GF_writeTextAlign(oPDF, 110, yInicio + ySeparation, ucase(auxCargo), 200, PDF_ALIGN_LEFT)
	ySeparation = ySeparation + SEPARATION

	auxDetalle = GF_Traducir("NO")  	  
	Call GF_writeTextAlign(oPDF, 20, yInicio + ySeparation, GF_TRADUCIR("Ver Detalle.....:")	, 50, PDF_ALIGN_LEFT)
	if (verDetalle)	then auxDetalle = GF_Traducir("SI")
	Call GF_writeTextAlign(oPDF, 110, yInicio + ySeparation, auxDetalle, 200, PDF_ALIGN_LEFT)
	ySeparation = ySeparation + SEPARATION

	yFinal = yInicio + (ySeparation + SEPARATION)

	armadoCabecera = yFinal 

end Function
'------------------------------------------------------------------------------------------	
Function dibujarDetalleVale(currentY, idArticulo, cantidad, idCategoria)
	Dim pDS, pAbrev
	
	if (idArticulo <> 0) then
		Call getArticuloFull(idArticulo, pDS, pAbrev)
		call GF_setFont(oPDF,"COURIER",6,0)
		Call GF_squareBox(oPDF, 100, currentY, 380, 10, 0, "#F5F6CE", "#F5F6CE", 1, PDF_SQUARE_NORMAL)
		Call GF_writeTextAlign(oPDF,104, currentY+2, idArticulo, 70, PDF_ALIGN_LEFT)	
		Call GF_writeTextAlign(oPDF,174, currentY+2, pDS, 260, PDF_ALIGN_LEFT)	
		Call GF_writeTextAlign(oPDF,430, currentY+2, cantidad & " " & pAbrev, 45, PDF_ALIGN_RIGHT)
		currentY = currentY + SEPARATION
	end if
	
	dibujarDetalleVale = currentY
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
%>