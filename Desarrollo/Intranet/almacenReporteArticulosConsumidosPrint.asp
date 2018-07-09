<!-- #include file="Includes/procedimientosAlmacenes.asp"-->
<!-- #include file="Includes/procedimientosUser.asp"-->
<!-- #include file="Includes/procedimientosFormato.asp"-->
<!-- #include file="Includes/procedimientosPDF.asp"-->
<!-- #include file="Includes/procedimientosVales.asp"-->
<!-- #include file="Includes/procedimientosParametros.asp"-->
<!-- #include file="Includes/procedimientosTraducir.asp"-->
<!-- #include file="Includes/procedimientosFechas.asp"-->


<%
Const LINEA_POST_ENCABEZADO = 75
Const PAGE_HEIGHT_SIZE = 800
Const SEPARACION_Y = 10

Dim nroPagina,oPDF,lineaActual, g_dsPeriodoActual, g_anio, g_dsDivision1, g_dsDivision2, g_idDivision1, g_idDivision2

nroPagina = 1

'-----------------------------------------------------------------------------------------------
Function dibujarFiltros(p_parametros)
'Funcion que dibuja los parametros de busqueda 
'recibe como parametro un vector con la siguiente estructura en cada posicion:
'	"nombreBusqueda,valorBusqueda"
'Devuelve la posicion donde se puede seguir escribiendo
	Dim aux,x_inicial,y_inicial,font_size
	Dim max_len
	
	my_font_size = 8
	x_inicial = 20
	y_inicial = 75
	
	max_len = 0
	for i = 0 to ubound(p_parametros)
		aux = split(p_parametros(i),",")
		if (len(aux(0)) > max_len) then
			max_len = len(aux(0))
		end if
	next
	
	for i = 0 to ubound(p_parametros)
		aux = split(p_parametros(i),",")
		Call GF_setFont(oPDF,"COURIER",my_font_size,0)
		Call GF_writeTextAlign(oPDF,x_inicial,y_inicial+(i*my_font_size), completarEspacios(aux(0),max_len) & ": " & aux(1) ,580, PDF_ALIGN_LEFT)
	next 
	
	dibujarFiltros = y_inicial+(i*my_font_size)+my_font_size
	
End Function
'--------------------------------------------------------------------------------------------------
'JAS - OK!
Function dibujarEncabezado(oPDF)
	Dim titulo

	Call GF_squareBox(oPDF, 2, 2, 590, 833, 0, "", "#0B3B0B", 2, PDF_SQUARE_ROUND)
	Call GF_writeImage(oPDF, Server.MapPath("images\kogge64.gif"), 20, 10, 48, 48, 0)
	
	titulo = "ARTICULOS CONSUMIDOS"
	
	Call GF_setFont(oPDF,"ARIAL",16,FONT_STYLE_BOLD)
	Call GF_writeTextAlign(oPDF,20,25,GF_TRADUCIR(titulo), 550 , PDF_ALIGN_CENTER)
	
	Call GP_CONFIGURARMOMENTOS()
	
	Call GF_setFont(oPDF,"COURIER",8,FONT_STYLE_NORMAL)
	Call GF_writeTextAlign(oPDF,5,5,GF_FN2DTE(session("MmtoSistema")), 580 , PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,5,5+pdf_currentFontSize,session("Usuario"), 580 , PDF_ALIGN_RIGHT)
	
	Call GF_writeTextAlign(oPDF, 10 , 840, "Pagina  " & nroPagina		 , 580 , PDF_ALIGN_RIGHT)
	Call GF_horizontalLine(oPDF,2,65,590)

	dibujarEncabezado = LINEA_POST_ENCABEZADO
	
end Function
'-----------------------------------------------------------------------------------------------
Function completarEspacios(p_palabra,p_len)
	Dim rtrn
	rtrn = p_palabra
	for i = len(p_palabra) to p_len
		rtrn = rtrn & "."
	next
	completarEspacios = rtrn
End Function
'--------------------------------------------------------------------------------------------------
'JAS - OK!
'/*
' * Parametros
' *		pPeriodoActual	: Periodo en análisis
' *		pAnioActual		: Anio de análisis
' *		dsDivision1		: Descripción de la división contra la que se compara
' *		dsDivision2		: Descripción de la división contra la que se compara
' */
Function dibujarTitulos()

	Call GF_squareBox(oPDF,  10, lineaActual, 35 , 20, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF,  45, lineaActual, 255, 20, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 300, lineaActual, 70 , 20, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 370, lineaActual, 70 , 20, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 440, lineaActual, 140, 10, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 440, lineaActual+10, 70 , 10, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 510, lineaActual+10, 70 , 10, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	'JAS - Call GF_squareBox(oPDF, 530, lineaActual, 50 , 20, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	
	Call GF_setFontColor("#FFFFFF")
	Call GF_setFont(oPDF,"ARIAL",8,FONT_STYLE_BOLD)
	Call GF_writeTextAlign(oPDF, 10 , lineaActual+5, GF_TRADUCIR("Id Art.")			,  35, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, 45 , lineaActual+5, GF_TRADUCIR("Descripcion")		, 255, PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(oPDF, 300, lineaActual+1, GF_TRADUCIR(g_dsPeriodoActual) ,  70, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, 300, lineaActual+9, g_anio							,  70, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, 370, lineaActual+1, GF_TRADUCIR(g_dsPeriodoActual)	,  70, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, 370, lineaActual+9, g_anio-1						,  70, PDF_ALIGN_CENTER)
	Call GF_setFont(oPDF,"ARIAL",6,FONT_STYLE_BOLD)
	Call GF_writeTextAlign(oPDF, 440, lineaActual+2, GF_TRADUCIR(g_dsPeriodoActual) & " " & g_anio ,  140, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, 440, lineaActual+12, GF_TRADUCIR(g_dsDivision1)		,  70, PDF_ALIGN_CENTER)	     	
	Call GF_writeTextAlign(oPDF, 510, lineaActual+12, GF_TRADUCIR(g_dsDivision2)		,  70, PDF_ALIGN_CENTER)		
	Call GF_setFont(oPDF,"ARIAL",8,FONT_STYLE_NORMAL)
	Call GF_setFontColor("#000000")
	
	dibujarTitulos = lineaActual + 21
	
End Function
'--------------------------------------------------------------------------------------------------
'Funcion que convierte un periodo elegido por pantalla a un par de fecha inicio-fin para armar el reporte.
'Esta funcion es totalmente dependiente del vector de periodos (ver getPeriodosDS)
Function obtenerFechas(pPeriodo, pAnio, byRef pFechaDesde, ByRef pFechaHasta)
	Dim mesInicio, mesFin, diaFin, diaInicio
	
	'Se determina en mes.	
	if (pPeriodo <= 11) then
		'Es un mes determinado. (Arranca Enero=0, esto es muy dependiente de las posiciones en el vector de periodos)
		mesInicio = pPeriodo+1
		mesFin = mesInicio
	else if (pPeriodo=16) then
			'Es el año entero.
			mesInicio = 1
			mesFin = 12
		else
			'Es un trimestre.
			'12: Primer Trimestre			
			'13: Segundo Trimestre
			'14: Tercer Trimestre
			'15: Cuarto Trimestre
			mesInicio = 1 + 3*(pPeriodo-12)
			mesFin	  = mesInicio + 2
		end if
	end if
	'Se determina el día
	diaInicio = 1
	diaFin = LastDayOfMonth(pAnio, mesFin)
	
	Call GF_STANDARIZAR_FECHA(diaInicio, mesInicio, pAnio)
	pFechaDesde = pAnio & mesInicio & diaInicio
	Call GF_STANDARIZAR_FECHA(diaFin, mesFin, pAnio)
	pFechaHasta = pAnio & mesFin & diaFin
End Function
'--------------------------------------------------------------------------------------------------
'Funcion responsable de dar formato de impresión a los valores para el consumo del periodo, del periodo del año anterior 
'y, si corresponde, acumular el valor para totalizar la informacion.
Function calcularValores(rsConsumo, idArticuloRef, abrevRef, tipoReporte, ByRef actual, ByRef anterior, ByRef totalActual, ByRef totalAnterior, ByRef actual1, ByRef totalActual1, ByRef actual2, ByRef totalActual2)

	Dim abrev
	
	actual = rsConsumo("consumoActual")
	if (isnull(actual)) then actual = 0
	anterior = rsConsumo("consumoAnterior")
	if (isnull(anterior)) then anterior = 0
	actual1 = rsConsumo("consumoActual1")
	if (isnull(actual1)) then actual1 = 0
	actual2 = rsConsumo("consumoActual2")
	if (isnull(actual2)) then actual2 = 0
								
	if (LCase(tipoReporte) <> REPORTE_CANTIDAD) then
		abrev = getSimboloMoneda(tipoReporte)
							
		totalActual = totalActual + CDbl(actual)
		totalAnterior = totalAnterior + CDbl(anterior)											
		totalActual1 = totalActual1 + CDbl(actual1)
		totalActual2 = totalActual2 + CDbl(actual2)
		
		actual = abrev & " " & GF_EDIT_DECIMALS(actual,2) 
		anterior = abrev & " " & GF_EDIT_DECIMALS(anterior,2)
		actual1 = abrev & " " & GF_EDIT_DECIMALS(actual1,2)
		actual2 = abrev & " " & GF_EDIT_DECIMALS(actual2,2)
	else
	    abrev = GF_nChars(abrevRef,3," ",CHR_AFT)
		actual = actual & " " & abrev
		anterior = anterior & " " & abrev
		actual1 = actual1 & " " & abrev
		actual2 = actual2 & " " & abrev
	end if
End Function
'--------------------------------------------------------------------------------------------------
Function dibujarContenido(pPeriodo, pTipoReporte, cdSolicitante, idDivision, idAlmacen, Detalle, articulo, categoria )
	Dim rsConsumo,auxDs,actual,anterior,i,myColor,abrev,totalActual,totalAnterior,totalAnual
	Dim fechaDesde, fechaHasta
	Dim actual1, anterior1, totalActual1, totalAnterior1, actual2, anterior2, totalActual2, totalAnterior2
	'Se obtienen las fechas del periodo en estudio.


	Call obtenerFechas(pPeriodo, g_anio, fechaDesde, fechaHasta)	
	set rsConsumo = getSQLConsumo(fechaDesde, fechaHasta, pTipoReporte, cdSolicitante, idDivision, g_idDivision1, g_idDivision2, idAlmacen, articulo, categoria)
	if (rsConsumo.EoF) then 
		'No hay datos asociados a los filtros introducidos.
		Call GF_squareBox(oPDF, 10 , lineaActual, 570 , 10, 0, "#FFFFFF", "#000000", 1, PDF_SQUARE_NORMAL)
		Call GF_writeTextAlign(oPDF, 10 , lineaActual+1, GF_TRADUCIR("No se encontraron resultados.") , 570, PDF_ALIGN_CENTER)
	else
		'Se procesan los resultados y se arma el reporte.
		i = 0		
		while not rsConsumo.EoF
			Call GF_setFont(oPDF,"COURIER",6,FONT_STYLE_NORMAL)
			myColor = "#BBFFBB"
			if ((i mod 2) = 0) then myColor = "#FFFFFF"
			Call GF_squareBox(oPDF, 10 , lineaActual, 570 , 10, 0, myColor, myColor, 1, PDF_SQUARE_NORMAL)			
			Call calcularValores(rsConsumo, 0, rsConsumo("abreviatura"), pTipoReporte, actual, anterior, totalActual, totalAnterior, actual1, totalActual1, actual2, totalActual2)
			if (len(auxDs) > 50) then auxDs = left(auxDs,50) & "..."
			Call GF_writeTextAlign(oPDF,   5,  lineaActual+1, rsConsumo("idArticulo")	, 35, PDF_ALIGN_RIGHT)
			Call GF_writeTextAlign(oPDF,  50,  lineaActual+1, rsConsumo("dsArticulo")	, 250, PDF_ALIGN_LEFT)
			Call GF_writeTextAlign(oPDF, 300,  lineaActual+1, actual					, 68, PDF_ALIGN_RIGHT)			
			Call GF_writeTextAlign(oPDF, 370,  lineaActual+1, anterior					, 68, PDF_ALIGN_RIGHT)
			Call GF_writeTextAlign(oPDF, 440,  lineaActual+1, actual1					, 68, PDF_ALIGN_RIGHT)			
			Call GF_writeTextAlign(oPDF, 510,  lineaActual+1, actual2					, 68, PDF_ALIGN_RIGHT)			
			if(lineaActual >= PAGE_HEIGHT_SIZE) then lineaActual = nuevaHoja()
			if(Detalle = "SI")then
			    lineaActual = lineaActual + SEPARACION_Y + 1
				call dibujarTitularBoxDetalle(lineaActual)		
				Call executeQueryDB(DBSITE_SQL_INTRA, rsConsumoDet, "OPEN", getSQLConsumoPeriodo(fechaDesde, fechaHasta, pTipoReporte, cdSolicitante, idDivision, idAlmacen, articulo, categoria, Detalle))				
				while (not rsConsumoDet.eof)
				    lineaActual = lineaActual + SEPARACION_Y
				    lineaActual = dibujarDetalleArticulo(lineaActual, pTipoReporte, rsConsumoDet("cdsolicitante"), rsConsumoDet("consumo"),rsConsumoDet("abreviatura"))
                    rsConsumoDet.MoveNext()
                wend                    				    
			end if

			lineaActual = lineaActual + SEPARACION_Y		
			'Control de paginacion.
			if(lineaActual >= PAGE_HEIGHT_SIZE) then lineaActual=nuevaHoja()
			i = i +1
			rsConsumo.MoveNext		
		wend
		if (LCase(pTipoReporte) <> REPORTE_CANTIDAD) then
			abrev = getSimboloMoneda(ptipoReporte)
			Call GF_setFont(oPDF,"COURIER",7,FONT_STYLE_BOLD)
			Call GF_squareBox(oPDF, 10 , lineaActual, 305 , 15, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
			Call GF_squareBox(oPDF, 300, lineaActual, 70 , 15,  0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
			Call GF_squareBox(oPDF, 370, lineaActual, 70 , 15,  0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
			Call GF_squareBox(oPDF, 440, lineaActual, 70 , 15,  0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
			Call GF_squareBox(oPDF, 510, lineaActual, 70 , 15,  0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)			
			Call GF_setFontColor("#FFFFFF")
			lineaActual = lineaActual +4
			Call GF_writeTextAlign(oPDF, 10 , lineaActual+1, "TOTAL" , 270, PDF_ALIGN_RIGHT)
			Call GF_writeTextAlign(oPDF, 300 , lineaActual+1, abrev & " " & GF_EDIT_DECIMALS(totalActual, 2), 68, PDF_ALIGN_RIGHT)
			Call GF_writeTextAlign(oPDF, 370 , lineaActual+1, abrev & " " & GF_EDIT_DECIMALS(totalAnterior, 2) , 68, PDF_ALIGN_RIGHT)			
			Call GF_writeTextAlign(oPDF, 440 , lineaActual+1, abrev & " " & GF_EDIT_DECIMALS(totalActual1, 2), 68, PDF_ALIGN_RIGHT)
			Call GF_writeTextAlign(oPDF, 510 , lineaActual+1, abrev & " " & GF_EDIT_DECIMALS(totalActual2, 2), 68, PDF_ALIGN_RIGHT)
		end if
	end if

End Function
'--------------------------------------------------------------------------------------------------
Function dibujarDetalleArticulo(lineaActual, tipoReporte, cdusuario, consumoActual, abreviatura)
    Dim abrev, consumo

    if (LCase(tipoReporte) <> REPORTE_CANTIDAD) then
		abrev = getSimboloMoneda(tipoReporte)
		consumo = abrev & " " & GF_EDIT_DECIMALS(consumoActual,2) 
	else
	    abrev = GF_nChars(abrevRef,3," ",CHR_AFT)
		consumo = consumoActual & " " & abrev		
	end if	
	Call GF_squareBox(oPDF, 100, lineaActual+2, 190, 10, 0, "#F5F6CE", "#F5F6CE", 1, PDF_SQUARE_NORMAL)
	Call GF_writeTextAlign(oPDF,102, lineaActual+3, getUserDescription(cdusuario), 120, PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(oPDF,222, lineaActual+3, consumo, 60, PDF_ALIGN_RIGHT)		
	dibujarDetalleArticulo = lineaActual	
end function 
'--------------------------------------------------------------------------------------------------
Function nuevaHoja()
	Call GF_newPage(oPDF)
	lineaActual = LINEA_POST_ENCABEZADO	
	nroPagina = nroPagina + 1
	Call dibujarEncabezado(oPDF)
	nuevaHoja = dibujarTitulos()
End Function
'--------------------------------------------------------------------------------------------------
'Funcion responsable de generar la SQL para obtener los consumos correspondientes a un periodo.
'Parametros
'	fechaInicioActual: Fecha de incio del periodo en estudio
'	fechaFinActual	 : Fecha de fin del periodo en estudio
'	unidad			 : Cantidad, P=Pesos o D=Dolares
'	cdSolicitante	 : Nombre de usuario para el que se quieren los consumos (Opcional)
'	idDivision		 : Id de la division foco del reporte
'	idDivision1		 : Id de una de las divisiones con la cual se compara 
'	idDivision2		 : Id de una de las divisiones con la cual se compara 
'	idAlmacen		 : Id del almacen (Opcional)
'   idArticulo       : Id del Articulo por el cual se filtra
'   idcategoria      : Id de la Categoria de articulos por la cual se filtra
'   Detalle          : Filtro de detalle del solicitante de los movimientos.
Function getSQLConsumo(fechaInicioActual, fechaFinActual, pTipoReporte, cdSolicitante, idDivision, idDivision1, idDivision2, idAlmacen, idarticulo, idcategoria)
	
	Dim strSQL,rs,conn, fechaFinAnterior, fechaInicioAnterior
	
	'Se determinan los limites de los periodos a analizar.
	
	fechaInicioAnterior = GF_DTEADD(fechaInicioActual, -1, "A")
	fechaFinAnterior = GF_DTEADD(fechaFinActual, -1, "A")
	
	strSQL = 		  "SELECT   actual.idarticulo    , "
	strSQL = strSQL & "         actual.dsarticulo    , "
	strSQL = strSQL & "         actual.consumo   consumoActual, "
	strSQL = strSQL & "         actual.abreviatura   abreviatura, "
	strSQL = strSQL & "         anterior.consumo consumoAnterior, "
	strSQL = strSQL & "         actual1.consumo   consumoActual1, "
	strSQL = strSQL & "         actual2.consumo   consumoActual2, "	
	strSQL = strSQL & "         actual.cantidad CANT "		
	strSQL = strSQL & "FROM		(" & getSQLConsumoPeriodo(fechaInicioActual,fechaFinActual,pTipoReporte, cdSolicitante, idDivision, idAlmacen,idarticulo, idcategoria, "") & ") actual"
	strSQL = strSQL & "			LEFT JOIN (" & getSQLConsumoPeriodo(fechaInicioAnterior, fechaFinAnterior, pTipoReporte, cdSolicitante, idDivision, idAlmacen,idarticulo, idcategoria, "") & ") anterior"
	strSQL = strSQL & "         	ON actual.idarticulo = anterior.idarticulo "
	strSQL = strSQL & "			LEFT JOIN (" & getSQLConsumoPeriodo(fechaInicioActual,fechaFinActual, pTipoReporte, cdSolicitante, idDivision1, 0, idarticulo, idcategoria, "") & ") actual1"
	strSQL = strSQL & "         	ON actual.idarticulo = actual1.idarticulo "
	strSQL = strSQL & "			LEFT JOIN (" & getSQLConsumoPeriodo(fechaInicioActual,fechaFinActual, pTipoReporte, cdSolicitante, idDivision2, 0, idarticulo, idcategoria, "") & ") actual2"
	strSQL = strSQL & "         	ON actual.idarticulo = actual2.idarticulo "
	strSQL = strSQL & "ORDER BY actual.consumo desc"	
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	set getSQLConsumo = rs
End Function
'--------------------------------------------------------------------------------------------------
'Funcion responsable de generar la SQL para obtener los consumos correspondientes a un periodo.
'Parametros
'	pFecIni			: Fecha de incio del periodo en estudio
'	pFecFin			: Primer dia posterior a la fecha de fin del periodo en estudio
'	pTipoReporte	: Cantidad, P=Pesos o D=Dolares
'	cdSolicitante	: Nombre de usuario para el que se quieren los consumos (Opcional)
'	idDivision		: Id de la division
'	idAlmacen		: Id del almacen (Opcional)
'				--------------MODIFICACION---------- 
'- CNA - Ajaya Nahuel - 20/03/2012
'-Se sumo la cantidad (existencia + sobrante) para saber cuantos articulos hay  (se guarda en la variable "cantidad")
Function getSQLConsumoPeriodo(pFecIni, pFecFin, pTipoReporte, cdSolicitante, idDivision, idAlmacen, idarticulo, idcategoria, verDetalle)
	dim strSQL 
	
	strSQL = "			SELECT  art.idarticulo, "
	strSQL = strSQL & "         art.dsarticulo, "
	strSQL = strSQL & "         uni.abreviatura abreviatura, "
	if(verDetalle = "SI") then strSQL = strSQL & "         cab.cdsolicitante, "	
	strSQL = strSQL & "         SUM(ABS(deta.cantidad)) cantidad, "	
	strSQL = strSQL & "SUM("
	select case (pTipoReporte)
	    case REPORTE_CANTIDAD:
	        strSQL = strSQL & "ABS(deta.cantidad)"
		case REPORTE_PESO:
			strSQL = strSQL & "ABS(deta.EXISTENCIA)*deta.VLUPESOS"
		case REPORTE_DOLAR:
			strSQL = strSQL & "ABS(deta.EXISTENCIA)*deta.VLUDOLARES"
	end select
	strSQL = strSQL & " ) consumo"
	strSQL = strSQL & " FROM "
	'Se toman solo las cabeceras de vales que pertenecen al periodo en analisis, que corresponden a consumos y que no estan dadas de baja.
	strSQL = strSQL & "			(	Select C.* from tblvalescabecera C inner join TBLALMACENES AL on C.IDALMACEN=AL.IDALMACEN"
	strSQL = strSQL & "				WHERE   c.fecha  >= " & pFecIni
	strSQL = strSQL & "				AND		c.fecha  <= " & pFecFin	
	strSQL = strSQL & "				AND		c.estado = "  & ESTADO_ACTIVO
	if (cdSolicitante <> "") then strSQL = strSQL & " AND c.cdsolicitante = '" & cdSolicitante & "' "	
	if (idAlmacen <> 0)		 then strSQL = strSQL & " AND c.idalmacen = " & idAlmacen
	if (idDivision <> 0)	 then strSQL = strSQL & " AND al.idDivision = " & idDivision
	strSQL = strSQL & "			) cab "
	'Se unen los datos del detalle de los vales.
	strSQL = strSQL & "         INNER JOIN tblvalesdetalle deta "
	strSQL = strSQL & "         ON       deta.idvale = cab.idvale "
	'Se completan los datos del articulo y su unidad.
	strSQL = strSQL & "         INNER JOIN tblarticulos art "
	strSQL = strSQL & "         ON       art.idarticulo = deta.idarticulo "
	strSQL = strSQL & "         INNER JOIN tblunidades uni "
	strSQL = strSQL & "         ON       uni.idunidad = art.idunidad "
	strSQL = strSQL & "			WHERE (cab.cdvale IN ('" & CODIGO_VS_SALIDA & "', '" & CODIGO_VS_AJUSTE_TRANSFERENCIA & "', '" & CODIGO_VS_AJUSTE_VALE & "') or (CAB.CDVALE = '" & CODIGO_VS_AJUSTE_STOCK & "' and DETA.EXISTENCIA < 0)) " & filtrarCategoriaArticulo(idcategoria, idarticulo)
	strSQL = strSQL & " GROUP BY art.idarticulo, "
	strSQL = strSQL & "          art.dsarticulo, "
	strSQL = strSQL & "          uni.abreviatura "
	if(verDetalle = "SI")then strSQL = strSQL & ",         cab.cdsolicitante "		
	getSQLConsumoPeriodo = Trim(strSQL)		
End Function 
'--------------------------------------------------------------------------------------------------
Function filtrarCategoriaArticulo(idcategoria, idarticulo)
dim myWhere
myWhere = ""
if(idarticulo <> 0) then  	myWhere = " AND deta.idarticulo = "& idarticulo
if(idcategoria > 0) then  	myWhere = myWhere &" AND art.idcategoria = "& idcategoria
filtrarCategoriaArticulo = myWhere
End function 
'--------------------------------------------------------------------------------------------------
'Funcion responsable por la creacion del PDF.
Function crearPDF(pDivision, pAlmacen, pPeriodo, pSolicitante, pTipoReporte, pVerDetalle,pidArticulo,pCategoria)
	
	Dim vParam(7),usr, auxCatDS
	
	Set oPDF = GF_createPDF("")
	Call GF_setPDFMode(PDF_STREAM_MODE)
	Call dibujarEncabezado(oPDF)
	
	usr = getUserDescription(pSolicitante)
	if (usr = "") then usr = "Todos"
	'Se dibujan los filtros que se aplicaron al reporte.
	vParam(0) = "Division," & getDivisionDS(pDivision)
	vParam(1) = "Almacen," & getAlmacenDs(pAlmacen)
	vParam(2) = "Periodo," & g_dsPeriodoActual & " " & pAnio
	vParam(3) = "Solicitante," & usr
	vParam(4) = "Consumo,Cantidades"
	if (pTipoReporte <> REPORTE_CANTIDAD) then
		if (pTipoReporte = REPORTE_PESO) then
			vParam(4) = "Consumo,Importe (Pesos)"
		else
			vParam(4) = "Consumo,Importe (Dolares)"
		end if
	end if
	vParam(5) = "Detalle," & pVerDetalle  
	if (pCategoria <= 0) then 
		vParam(6) = "Categoria,TODAS"
	else
		auxCatDS = getCategoriaFull(pCategoria)	
		vParam(6) = "Categoria," &auxCatDS 
	end if
	if(pidArticulo = 0) then 
		vParam(7) = "Articulo,TODOS"
	else
		call getArticuloFull(pidArticulo, auxArtDS, auxArtAbrev)
		vParam(7) = "Articulo," & pidArticulo &"-"& auxArtDS		
	end if	
	lineaActual = dibujarFiltros(vParam)
	lineaActual = dibujarTitulos()
	'Se dibuja el cuerpo del reporte con los datos resultado.
	Call dibujarContenido(pPeriodo, pTipoReporte, pSolicitante, pDivision, pAlmacen,pVerDetalle,pidArticulo,pCategoria)
	
	Call GF_closePDF(oPDF)	
	
End Function
'--------------------------------------------------------------------------------------------------
'Funcion responsable por obtener los id de las divisiones contra las cuales se van a comparar los consumos.
Function setDivisionesComparacion(pIdDivisionActual, ByRef pIdDivision1, ByRef pIdDivision2)

	Dim strSql, rs
	
	'Busco las divisiones que corresponden a los puertos que no son el que se esta analizando.
	strSQL="Select * from TBLDIVISIONES where IDDIVISION<>" & pIdDivisionActual & " and CDDIVISIONABR<>'" & CODIGO_EXPORTACION & "'"
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	
	pIdDivision1 = ""
	pIdDivision2 = ""
	if (not rs.eof) then
		pIdDivision1 = rs("IDDIVISION")
		rs.MoveNext()
	end if
	if (not rs.eof) then
		pIdDivision2 = rs("IDDIVISION")
		rs.MoveNext()
	end if

End Function
'--------------------------------------------------------------------------------------------------
Function getAlmacenDs(pId)
	Dim strSQL,rs,conn,rtrn
	
	rtrn = ""
	strSQL = "Select * from TBLALMACENES where idalmacen = " & pId
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	
	if (not rs.EoF) then rtrn = rs("dsAlmacen")
	
	getAlmacenDs = rtrn
	
End Function 
'--------------------------------------------------------------------------------------------------
Function emitirReporte(pAlmacen, pDivision, pPeriodo, pPeriodoDS, pAnio, pSolicitante, pTipoReporte, pVerDetalle, pidArticulo, pCategoria)
	'--- DATOSGLOBALES ---
	'Se determinan las 2 divisiones contra las que se comparan los datos.
	Call setDivisionesComparacion(pDivision, g_idDivision1, g_idDivision2)
	'Se toman sus descripciones	
	g_dsDivision1 = getDivisionDS(g_idDivision1)
	g_dsDivision1 = left(g_dsDivision1,16)
	g_dsDivision2 = getDivisionDS(g_idDivision2)
	g_dsDivision2 = left(g_dsDivision2,16)

	g_anio = pAnio

	g_dsPeriodoActual = pPeriodoDS

	'-- Se crea el Reporte.	
	Call crearPDF(pDivision, pAlmacen, pPeriodo, pSolicitante, pTipoReporte, pVerDetalle, pidArticulo, pCategoria)
End Function
'--------------------------------------------------------------------------------------------------
Function dibujarTitularBoxDetalle(pY)
	Call GF_squareBox(oPDF, 100, pY, 120, 10, 0, "#F2F5A9", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 220, pY, 70, 10, 0, "#F2F5A9", "#000000", 1, PDF_SQUARE_NORMAL)
	call GF_setFont(oPDF,"ARIAL",6,8)
	call GF_setFontColor("000000")
	Call GF_writeTextAlign(oPDF, 102, pY+2	, GF_TRADUCIR("USUARIO"), 120	, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, 222, pY+2	, GF_TRADUCIR("CANTIADAD"), 70	, PDF_ALIGN_CENTER)	
	call GF_setFontColor("000000")
end Function

'**************************************************************************************************
'*                                                                                                *
'*                                   INICIO DE PAGINA                                             *
'*                                                                                                *
'**************************************************************************************************

Dim pAlmacen,pPeriodo,pAnio,pSolicitante,pUnidad, pDivision, pVerDetalle, pidArticulo, pCategoria
Dim pPeriodoDS

pAlmacen	= GF_PARAMETROS7("almacen", 0, 6)
pDivision	= GF_PARAMETROS7("division", 0, 6)
pPeriodo	= GF_PARAMETROS7("periodo", 0, 6)
pPeriodoDS	= GF_PARAMETROS7("periodoDS", "", 6)
pAnio		= GF_PARAMETROS7("anio", 0,6)
pSolicitante= GF_PARAMETROS7("cdSolicitante","",6)
pTipoReporte= GF_PARAMETROS7("tipoReporte","",6)
pVerDetalle = GF_PARAMETROS7("verDetalle","",6)
if(pVerDetalle = "") then pVerDetalle = "NO"
pidArticulo = GF_PARAMETROS7("idArticulo",0,6)
pCategoria	= GF_PARAMETROS7("categoria", 0, 6)

Call emitirReporte(pAlmacen, pDivision, pPeriodo, pPeriodoDS, pAnio, pSolicitante, pTipoReporte, pVerDetalle, pidArticulo, pCategoria)


%>