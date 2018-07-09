<!-- #include file="Includes/procedimientosMG.asp"-->
<!-- #include file="Includes/procedimientosAlmacenes.asp"-->
<!-- #include file="Includes/procedimientosUser.asp"-->
<!-- #include file="Includes/procedimientosFormato.asp"-->
<!-- #include file="Includes/procedimientosPDF.asp"-->
<!-- #include file="Includes/procedimientosVales.asp"-->
<!-- #include file="Includes/procedimientosParametros.asp"-->
<!-- #include file="Includes/procedimientosTraducir.asp"-->
<!-- #include file="Includes/procedimientosFechas.asp"-->
<!-- #include file="Includes/procedimientosCTZ.asp"-->

<%
Const LINEA_POST_ENCABEZADO = 75
Const ARTICULOS_CON_STOCK = 1
Const ARTICULOS_SIN_STOCK = 2
Const TITULO_SIN_STOCK = "sin"
Const TITULO_CON_STOCK = "con"

'-----------------------------------------------------------------------------------------------
Function dibujarEncabezado(oPDF)
	Dim titulo

	Call GF_squareBox(oPDF, 2, 2, 590, 833, 0, "", "#0B3B0B", 2, PDF_SQUARE_ROUND)
	Call GF_writeImage(oPDF, Server.MapPath("images\kogge64.gif"), 20, 10, 48, 48, 0)
	
	titulo = "ARTICULOS NO CONSUMIDOS"
	
	Call GF_setFont(oPDF,"ARIAL",16,FONT_STYLE_BOLD)
	Call GF_writeTextAlign(oPDF,20,25,GF_TRADUCIR(titulo), 550 , PDF_ALIGN_CENTER)
	
	Call GP_CONFIGURARMOMENTOS()
	
	Call GF_setFont(oPDF,"COURIER",8,FONT_STYLE_NORMAL)
	Call GF_writeTextAlign(oPDF,5,5,GF_FN2DTE(session("MmtoSistema")), 580 , PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,5,5+pdf_currentFontSize,session("Usuario"), 580 , PDF_ALIGN_RIGHT)
	
	Call GF_writeTextAlign(oPDF, 10 , 840, "Pagina  " & nroPagina		 , 580 , PDF_ALIGN_RIGHT)
	Call GF_horizontalLine(oPDF,2,65,590)

	dibujarEncabezado = LINEA_POST_ENCABEZADO
End Function
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
	y_inicial = LINEA_POST_ENCABEZADO
	
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
'-----------------------------------------------------------------------------------------------
Function completarEspacios(p_palabra,p_len)
	Dim rtrn
	rtrn = p_palabra
	for i = len(p_palabra) to p_len
		rtrn = rtrn & "."
	next
	completarEspacios = rtrn
End Function
'-------------------------------------------------------------------------------------------------'
'Articulos que figuran en los vales dentro del periodo'
Function sqlArticulosEnVales(fechaInicio,fechaFin)
	Dim strSQL

	strSQL =          "SELECT DISTINCT(v.idarticulo) "
	strSQL = strSQL & "FROM             tblvalesdetalle v "
	strSQL = strSQL & "INNER JOIN tblvalescabecera c "
	strSQL = strSQL & "ON               v.idvale = c.idvale "
	strSQL = strSQL & " AND              c.fecha  < " & fechaFin
	strSQL = strSQL & " AND              c.fecha  > " & fechaInicio

	sqlArticulosEnVales = strSQL

End Function
'-------------------------------------------------------------------------------------------------'
'Articulos que figuran en los PICs dentro del periodo'
Function sqlArticulosEnPics(fechaInicio,fechaFin)
	Dim strSQL

	strSQL =          "SELECT DISTINCT(v.idarticulo) "
	strSQL = strSQL & "FROM tblctzdetalle v "
	strSQL = strSQL & "INNER JOIN tblctzcabecera c "
	strSQL = strSQL & "		ON    v.idcotizacion = c.idcotizacion "
	strSQL = strSQL & "		AND   c.momento      < "&fechaFin&"000000 "
	strSQL = strSQL & "		AND   c.momento      > "&fechaInicio&"000000 "
	strSQL = strSQL & "		AND   c.estado      <> '"&CTZ_ANULADA&"' "
	
	sqlArticulosEnPics = strSQL

End Function
'-------------------------------------------------------------------------------------------------'
'Devuelve una sql con los articulos que tienen o no stock (dependiendo del prametro pTipo)'
'que no tuvieron movimientos en el periodo dado'
Function sqlArticulos(pTipo,fechaInicio,fechaFin)
	Dim strSQL

	strSQL =          "SELECT   art.idarticulo "
	strSQL = strSQL & "FROM     tblarticulos art "
	strSQL = strSQL & "         LEFT JOIN tblarticulosdatos artd "
	strSQL = strSQL & "         ON       art.idarticulo = artd.idarticulo "
	strSQL = strSQL & "WHERE     art.idarticulo NOT IN ( " & sqlArticulosEnVales(fechaInicio,fechaFin) & ") "
	strSQL = strSQL & " AND      art.idarticulo NOT IN ( " & sqlArticulosEnPics(fechaInicio,fechaFin) & ") "
	strSQL = strSQL & " AND      estado              = " & ESTADO_ACTIVO
	select case pTipo
		case ARTICULOS_CON_STOCK
			strSQL = strSQL & " AND      existencia+sobrante > 0 "
		case ARTICULOS_SIN_STOCK
			strSQL = strSQL & " AND      existencia+sobrante = 0 "
	end select
	strSQL = strSQL & "GROUP BY art.idarticulo"

	sqlArticulos = strSQL
End Function
'-------------------------------------------------------------------------------------------------'
Function dibujarSeccionConStock(oPDF,fechaInicio,fechaFin)
	Dim strSQL,rs,lastArticulo,lastArticuloDs,lastUnidad,vCantidad,i,campoVacio

	'servira para guardar cantidades, importe en pesos o importe en dolares dependiendo del tipo
	'de reporte que se haya solicitado
	redim vCantidad(4) 

	strSQL =          "SELECT art.idarticulo , "
	strSQL = strSQL & "       art.dsarticulo , "
	strSQL = strSQL & "       isnull(artd.existencia,0) existencia, "
	strSQL = strSQL & "       isnull(artd.sobrante,0) sobrante, "
	strSQL = strSQL & "       uni.abreviatura, "
	strSQL = strSQL & "       isnull(artp.vlupesos,0) "&MONEDA_PESO&", "
	strSQL = strSQL & "       isnull(artp.vludolares,0) "&MONEDA_DOLAR&", "
	strSQL = strSQL & "       artd.idalmacen, "
	strSQL = strSQL & "       alm.iddivision "
	strSQL = strSQL & "FROM   (	SELECT * "
	strSQL = strSQL & "       	FROM    tblarticulos art "
	strSQL = strSQL & "       	WHERE   estado = " & ESTADO_ACTIVO
	strSQL = strSQL & "       	AND     idarticulo IN ("&sqlArticulos(ARTICULOS_CON_STOCK,fechaInicio,fechaFin)&") "
	strSQL = strSQL & "       ) art "
	strSQL = strSQL & "       LEFT JOIN tblarticulosdatos artd "
	strSQL = strSQL & "       	ON     art.idarticulo = artd.idarticulo "
	strSQL = strSQL & "       INNER JOIN tblunidades uni "
	strSQL = strSQL & "       	ON     art.idunidad = uni.idunidad "
	strSQL = strSQL & "       LEFT JOIN "
	strSQL = strSQL & "              ( SELECT b.idarticulo, "
	strSQL = strSQL & "                      vlupesos     , "
	strSQL = strSQL & "                      vludolares   , "
	strSQL = strSQL & "                      a.ultimo "
	strSQL = strSQL & "              FROM    tblarticulosprecios b "
	strSQL = strSQL & "                      INNER JOIN "
	strSQL = strSQL & "                              (	SELECT  MAX(mmtoprecio) ultimo,idarticulo "
	strSQL = strSQL & "                              	FROM     tblarticulosprecios "
	strSQL = strSQL & "                              	GROUP BY idarticulo "
	strSQL = strSQL & "                              ) a "
	strSQL = strSQL & "                      ON      a.ultimo     = b.mmtoprecio "
	strSQL = strSQL & "                      AND     a.idarticulo = b.idarticulo "
	strSQL = strSQL & "              ) artp "
	strSQL = strSQL & "       	ON     artp.idarticulo = art.idarticulo "
	strSQL = strSQL & "         LEFT JOIN tblalmacenes alm on artd.idalmacen = alm.idalmacen"
	strSQL = strSQL & " GROUP BY art.idarticulo , "
	strSQL = strSQL & "         art.dsarticulo , "
	strSQL = strSQL & "         artd.existencia, "
	strSQL = strSQL & "         artd.sobrante  , "
	strSQL = strSQL & "         uni.abreviatura, "
	strSQL = strSQL & "         artp.vlupesos  , "
	strSQL = strSQL & "         artp.vludolares, "
	strSQL = strSQL & "         artd.idalmacen, "
	strSQL = strSQL & "         alm.iddivision "


	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)

	
	Call GF_setFontColor("#FFFFFF")
	Call GF_setFont(oPDF,"COURIER",13,FONT_STYLE_BOLD)

	Call GF_squareBox(oPDF, 10 , lineaActual, 575 , 15, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_writeTextAlign(oPDF, 10 , lineaActual+1, "ARTICULOS CON STOCK" , 575, PDF_ALIGN_CENTER)

	lineaActual = lineaActual + 15
	Call GF_setFont(oPDF,"COURIER",8,FONT_STYLE_BOLD)
	Call dibujarTitulo(TITULO_CON_STOCK)

	if (not rs.EoF) then 
		lastArticuloId = rs("idarticulo")
		lastArticuloDs = rs("dsarticulo")
		lastUnidad     = rs("abreviatura")
	end if

	Call GF_setFontColor("#000000")		
	Call GF_setFont(oPDF,"COURIER",8,FONT_STYLE_NORMAL)
	
	while not rs.EoF
		'los que posean idalmacen nulo es porque no figuran en la tabla articulosdatos 
		' y por lo tanto no tienen asignado ningun almacen
		if (not isnull(rs("idalmacen"))) then
			if (lastArticuloId <> rs("idarticulo")) then
				contRegistros = contRegistros + 1
				myColor = "#ffffff"
				if (contRegistros mod 2 = 0) then myColor = "#bbffbb"

				'imprimo el array'
				Call GF_squareBox(oPDF, 10 , lineaActual, 275 , 10, 0, myColor, "#ffffff", 1, PDF_SQUARE_NORMAL)
				Call GF_writeTextAlign(oPDF, 15 , lineaActual+1, lastArticuloId & " - " & lastArticuloDs , 275, PDF_ALIGN_LEFT)

				for i = 0 to 2
					Call GF_squareBox(oPDF, 285+(i*100) , lineaActual, 100 , 10, 0, myColor, "#ffffff", 1, PDF_SQUARE_NORMAL)

					if (len(vCantidad(i+2)) <> 0) then
						Call GF_writeTextAlign(oPDF, 285+(i*100) , lineaActual+1, vCantidad(i+2) , 95, PDF_ALIGN_RIGHT)
					else
						Call GF_writeTextAlign(oPDF, 285+(i*100) , lineaActual+1, campoVacio , 95, PDF_ALIGN_RIGHT)	
					end if
				next	

				lineaActual = lineaActual + 10
				lastArticuloId = rs("idarticulo")
				lastArticuloDs = rs("dsarticulo")
				lastUnidad     = rs("abreviatura")
				'limpio el array'
				redim vCantidad(4)

			end if

			'Si la fecha no es null la guardo en el array para su posterior insercion'
			select case pTipoReporte
				case REPORTE_CANTIDAD
					campoVacio = "0 " & lastUnidad
					vCantidad(rs("iddivision")) = cdbl(rs("existencia"))+cdbl(rs("sobrante")) & " " & lastUnidad
				case REPORTE_PESO
					campoVacio = getSimboloMoneda(MONEDA_PESO) & " 0,00"
					vCantidad(rs("iddivision")) = getSimboloMoneda(MONEDA_PESO)  & " " & GF_EDIT_DECIMALS(cdbl(rs(MONEDA_PESO)) * cdbl(rs("existencia")),2)
				case REPORTE_DOLAR
					campoVacio = getSimboloMoneda(MONEDA_DOLAR) & " 0,00"
					vCantidad(rs("iddivision")) = getSimboloMoneda(MONEDA_DOLAR)  & " " & GF_EDIT_DECIMALS(cdbl(rs(MONEDA_DOLAR)) * cdbl(rs("existencia")),2)
			end select

			'control de paginacion'		
			if (lineaActual >= 800) then
				Call nuevaHoja()
				Call dibujarTitulo(TITULO_CON_STOCK)
			end if
		end if
		rs.MoveNext
	wend
End Function
'-------------------------------------------------------------------------------------------------'
Function dibujarSeccionSinStock(oPDF,pFechaInicio,pFechaFin)
	Dim strSQL,rs,lastArticulo,lastArticuloDs,vFechas,i

	redim vFechas(4)
	lastArticuloId = ""
	lastArticuloDs = ""

	strSQL =          "SELECT   artSinStock.idarticulo, "
	strSQL = strSQL & "         artSinStock.dsarticulo, "
	strSQL = strSQL & "         artSinStock.idalmacen , "
	strSQL = strSQL & "         MAX(c.fecha) ultimo "
	strSQL = strSQL & "FROM     ( SELECT art.idarticulo, "
	strSQL = strSQL & "                 dsarticulo     , "
	strSQL = strSQL & "                 idalmacen      , "
	strSQL = strSQL & "                 existencia+sobrante stockTotal "
	strSQL = strSQL & "         FROM    tblarticulos art "
	strSQL = strSQL & "                 INNER JOIN tblarticulosdatos artd "
	strSQL = strSQL & "                 ON      art.idarticulo = artd.idarticulo "
	strSQL = strSQL & "         WHERE   art.idarticulo NOT IN ( " & sqlArticulosEnVales(fechaInicio,fechaFin) &") "
	strSQL = strSQL & "         AND     art.idarticulo NOT IN ( " & sqlArticulosEnPics(fechaInicio,fechaFin) & ") "
	strSQL = strSQL & "         AND     estado              = " & ESTADO_ACTIVO
	strSQL = strSQL & "         AND     existencia+sobrante = 0 "
	strSQL = strSQL & "         ) "
	strSQL = strSQL & "         artSinStock "
	strSQL = strSQL & "         LEFT JOIN tblvalesdetalle d "
	strSQL = strSQL & "         ON       d.idarticulo = artSinStock.idarticulo "
	strSQL = strSQL & "         LEFT JOIN tblvalescabecera c "
	strSQL = strSQL & "         ON       c.idvale = d.idvale "
	strSQL = strSQL & "         AND      c.fecha  < " & fechaInicio
	strSQL = strSQL & " GROUP BY artSinStock.idarticulo, "
	strSQL = strSQL & "         artSinStock.dsarticulo, "
	strSQL = strSQL & "         artSinStock.idalmacen "
	strSQL = strSQL & "ORDER BY artSinStock.idarticulo, "
	strSQL = strSQL & "         artSinStock.idalmacen"
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)

	Call GF_setFontColor("#FFFFFF")
	Call GF_setFont(oPDF,"COURIER",13,FONT_STYLE_BOLD)

	Call GF_squareBox(oPDF, 10 , lineaActual, 575 , 15, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_writeTextAlign(oPDF, 10 , lineaActual+1, "ARTICULOS SIN STOCK" , 575, PDF_ALIGN_CENTER)

	lineaActual = lineaActual + 15
	Call GF_setFont(oPDF,"COURIER",8,FONT_STYLE_BOLD)
	Call dibujarTitulo(TITULO_SIN_STOCK)


	if (not rs.EoF) then 
		lastArticuloId = rs("idarticulo")
		lastArticuloDs = rs("dsarticulo")
	end if

	Call GF_setFontColor("#000000")		
	Call GF_setFont(oPDF,"COURIER",8,FONT_STYLE_NORMAL)

	while not rs.EoF
		if (lastArticuloId <> rs("idarticulo")) then
			'imprimo el array'
			contRegistros = contRegistros + 1
			myColor = "#ffffff"
			if (contRegistros mod 2 = 0) then myColor = "#bbffbb"

			Call GF_squareBox(oPDF, 10 , lineaActual, 275 , 10, 0, myColor, "#ffffff", 1, PDF_SQUARE_NORMAL)
			Call GF_writeTextAlign(oPDF, 15 , lineaActual+1, lastArticuloId & " - " & lastArticuloDs , 275, PDF_ALIGN_LEFT)

			for i = 0 to 2
				Call GF_squareBox(oPDF, 285+(i*100) , lineaActual, 100 , 10, 0, myColor, "#ffffff", 1, PDF_SQUARE_NORMAL)
				if (len(vFechas(i+2)) <> 0) then
					
					Call GF_writeTextAlign(oPDF, 285+(i*100) , lineaActual+1, GF_FN2DTE(vFechas(i+2)) , 100, PDF_ALIGN_CENTER)
				end if
			next	

			lineaActual = lineaActual + 10
			lastArticuloId = rs("idarticulo")
			lastArticuloDs = rs("dsarticulo")

			'limpio el array'
			redim vFechas(4)

		end if

		'Si la fecha no es null la guardo en el array para su posterior insercion'
		if (not isnull(rs("ultimo"))) then
			vFechas(cdbl(getDivisionAlmacen(rs("idalmacen")))) = rs("ultimo")
		end if

		'control de paginacion'
		if (lineaActual >= 800) then
			Call nuevaHoja()
			Call dibujarTitulo(TITULO_SIN_STOCK)
		end if
		
		rs.MoveNext
	wend

End Function
'--------------------------------------------------------------------------------------------------
Function dibujarTitulo(pTipo)
	Call GF_setFontColor("#FFFFFF")
	Call GF_setFont(oPDF,"COURIER",8,FONT_STYLE_BOLD)
	Call GF_squareBox(oPDF, 10 , lineaActual, 275 , 30, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_writeTextAlign(oPDF, 10 , lineaActual+10, "Articulo" , 275, PDF_ALIGN_CENTER)
	Call GF_squareBox(oPDF, 285 , lineaActual, 300 , 15, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	
	
	if (pTipo = TITULO_SIN_STOCK) then
		Call GF_writeTextAlign(oPDF, 285 , lineaActual+3, "Fecha Ultimo Movimiento Registrado" ,300 , PDF_ALIGN_CENTER)
	elseif (pTipo = TITULO_CON_STOCK) then
		select case pTipoReporte 
			case  REPORTE_CANTIDAD
				Call GF_writeTextAlign(oPDF, 285 , lineaActual+3, "Stock Total" ,300 , PDF_ALIGN_CENTER)
			case  REPORTE_PESO
				Call GF_writeTextAlign(oPDF, 285 , lineaActual+3, "Valoracion en Pesos" ,300 , PDF_ALIGN_CENTER)
			case REPORTE_DOLAR
				Call GF_writeTextAlign(oPDF, 285 , lineaActual+3, "Valoracion en Dolares" ,300 , PDF_ALIGN_CENTER)
		end select
	end if

	lineaActual = lineaActual + 15

	for i = 0 to 2
		Call GF_squareBox(oPDF, 285+(i*100) , lineaActual, 100 , 15, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
		Call GF_writeTextAlign(oPDF, 285+(i*100) , lineaActual+3, vDivisionesDs(i+1) , 100, PDF_ALIGN_CENTER)
	next
	lineaActual = lineaActual + 15
	Call GF_setFont(oPDF,"COURIER",8,FONT_STYLE_NORMAL)
	Call GF_setFontColor("#000000")

End Function
'--------------------------------------------------------------------------------------------------
Function nuevaHoja()
	nroPagina = nroPagina + 1
	lineaActual = LINEA_POST_ENCABEZADO
	Call GF_newPage(oPDF)
	Call dibujarEncabezado(oPDF)
End Function
'#################################################################################################'
'#################################################################################################'
'############################        INICIO DE PAGINA          ###################################'
'#################################################################################################'
'#################################################################################################'

Dim pTipoReporte,oPDF,fechaInicio,fechaFin,vFiltros,lineaActual,vDivisionesDs(3)
Dim nroPagina,contRegistros,myColor,pTipoReporte2

contRegistros = 0
fechaInicio	= GF_PARAMETROS7("fechaInicio", "", 6)
fechaFin	= GF_PARAMETROS7("fechaFin", "", 6)
pTipoReporte= GF_PARAMETROS7("tipoReporte","",6)
if (pTipoReporte = "") then pTipoReporte = REPORTE_CANTIDAD
pTipoReporte2= GF_PARAMETROS7("tipoReporte2","",6)
if (pTipoReporte2 = "") then pTipoReporte2 = REPORTE2_AMBOS

	nroPagina = 1	

	'creo el PDF'
	Set oPDF = GF_createPDF("")
	Call GF_setPDFMode(PDF_STREAM_MODE)

	Call dibujarEncabezado(oPDF)

	redim vFiltros (2)

	vFiltros(0) = "Fecha Inicio,"&GF_FN2DTE(fechaInicio)
	vFiltros(1) = "Fecha Fin,"&GF_FN2DTE(fechaFin)
	if (pTipoReporte = REPORTE_CANTIDAD) then
		vFiltros(2) = "Tipo Reporte,"&pTipoReporte
	else
		if (pTipoReporte = REPORTE_DOLARES) then
			vFiltros(2) = "Tipo Reporte,Dolares ("&getSimboloMonedaLetras(pTipoReporte)&")"
		else
			vFiltros(2) = "Tipo Reporte,Pesos ("&getSimboloMonedaLetras(pTipoReporte)&")"
		end if
	end if
	lineaActual = dibujarFiltros(vFiltros)

	vDivisionesDs(0) = "Exportacion"
	vDivisionesDs(1) = "Arroyo"
	vDivisionesDs(2) = "Piedra Buena"
	vDivisionesDs(3) = "El Transito"

	if (pTipoReporte2 = REPORTE2_SIN_STOCK or pTipoReporte2 = REPORTE2_AMBOS) then
		Call dibujarSeccionSinStock(oPDF,fechaInicio,fechaFin)
		'si me pidieron los 2 reportes, creo una nueva hoja para el siguiente'
		if (pTipoReporte2 = REPORTE2_AMBOS) then Call nuevaHoja()
	end if
	if (pTipoReporte2 = REPORTE2_CON_STOCK or pTipoReporte2 = REPORTE2_AMBOS) then
		Call dibujarSeccionConStock(oPDF,fechaInicio,fechaFin)
	end if
	Call GF_closePDF(oPDF)
%>