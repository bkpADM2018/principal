<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosPDF.asp"-->
<!--#include file="Includes/procedimientossql.asp"-->
<%

Dim strSQL, conn,rs,nroPagina,oPDF,rsAlmacenes, tablaDestino, conFecha
Dim accion,categoria,metodo,almacen,COMIENZO_COL_CODIGO, valorizar, fechaBusqueda

Const FILTRO_TODOS		= 0
Const FILTRO_EXISTENTES = 1

Const TEXT_SIZE = 8

Const COMIENZO_COL_ID			= 15
Const COMIENZO_COL_DS			= 40
Const COMIENZO_COL_EXISTENCIA	= 265
Const COMIENZO_COL_SOBRANTE		= 310
Const COMIENZO_COL_TOTAL		= 360
Const COMIENZO_COL_VALORACION	= 405
Const COMIENZO_COL_PRECIOUNIT   = 475

'******************************************************************************************
Function dibujarEncabezado(p_titulo,p_subtitulo)
	'Devuelve la posicion donde se puede seguir escribiendo
	Dim parametros(),rs1,rs2,dscategoria,dsalmacen,busqueda, fecha
	if (valorizar) then
		redim parametros(4)
	else
		redim parametros(3)
	end if
	
	Call GF_setFont(oPDF,"ARIAL",16,8)
	Call GF_writeTextAlign(oPDF,5,25,p_titulo, 580 , PDF_ALIGN_CENTER)
	Call GF_setFont(oPDF,"ARIAL",8,0)
	Call GF_writeTextAlign(oPDF,5,42,p_subtitulo, 580 , PDF_ALIGN_CENTER)
	Call GF_horizontalLine(oPDF,2,65,590)	
	
	Call GF_setFont(oPDF,"COURIER",8,0)
	GP_CONFIGURARMOMENTOS
	Call GF_writeTextAlign(oPDF,5,5,GF_FN2DTE(session("MmtoSistema")), 580 , PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,5,5+pdf_currentFontSize,session("Usuario"), 580 , PDF_ALIGN_RIGHT)
	'obtengo las descripciones de los parametros de busqueda
	if (nroPagina = 1) then
		if (categoria <> -1) then
			strSQl = "select idcategoria id,dscategoria ds from tblartcategorias where idcategoria = " & categoria
			Call executeQueryDB(DBSITE_SQL_INTRA, rs1, "OPEN", strSQL)
			dscategoria= rs1("ds")
		else
			dscategoria="Todas"
		end if
		strSQL = "select idalmacen id,dsalmacen ds from tblalmacenes where idalmacen = " & almacen		
		Call executeQueryDB(DBSITE_SQL_INTRA, rs2, "OPEN", strSQL)
		if (not rs2.eof) then dsalmacen=rs2("ds")
		if (metodo = FILTRO_TODOS) then
			busqueda="Todos"
		else
			busqueda="Con Stock"
		end if
		fecha = GF_FN2DTE(fechaBusqueda)
		
		parametros(0)= "Categoria" & "," & dscategoria
		parametros(1)="Almacen" & "," & dsalmacen
		parametros(2)="Buscar" & "," & busqueda
		if incluir then
			parametros(3)="Stocks al" & "," & fecha & " (Datos calculados al final de la fecha seleccionada.)"
		else
			parametros(3)="Stocks al" & "," & fecha & " (Datos calculados al inicio de la fecha seleccionada.)"		
		end if	
		
		if (valorizar) then
			parametros(4)="Valorizar" & "," & "Si"
		end if
		
		dibujarEncabezado = dibujarFiltros(parametros)
	else
		dibujarEncabezado = 75
	end if
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
'-----------------------------------------------------------------------------------------------
Function completarEspacios(p_palabra,p_len)
	Dim rtrn
	rtrn = p_palabra
	for i = len(p_palabra) to p_len
		rtrn = rtrn & "."
	next
	completarEspacios = rtrn
End Function
'-----------------------------------------------------------------------------------------------
Function nuevaPagina()
	nroPagina = nroPagina +1
	if (nroPagina > 1) then Call GF_newPage(oPDF)
	nuevaPagina = dibujarPagina()
End Function
'-----------------------------------------------------------------------------------------------
Function dibujarPagina()
	Dim y_comienzo
	
	Call GF_squareBox(oPDF, 2, 2, 590, 833, 0, "", "#0B3B0B", 2, PDF_SQUARE_ROUND)
	
	Call GF_writeImage(oPDF, Server.MapPath("images\kogge64.gif"), 20, 10, 48, 48, 0)
	
	y_comienzo =  dibujarEncabezado("STOCK DE ARTICULOS","")
	
	'dibuja los recuadros del titulo
	Call GF_squareBox(oPDF, COMIENZO_COL_ID, y_comienzo, 25, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, COMIENZO_COL_DS, y_comienzo, 230, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, COMIENZO_COL_EXISTENCIA, y_comienzo,45, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, COMIENZO_COL_SOBRANTE  , y_comienzo,50, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, COMIENZO_COL_TOTAL     , y_comienzo,45, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	if (valorizar) then
		Call GF_squareBox(oPDF, COMIENZO_COL_VALORACION, y_comienzo,70, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	end if
	Call GF_squareBox(oPDF, COMIENZO_COL_CODIGO    , y_comienzo,580-COMIENZO_COL_CODIGO, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	
	Call GF_setFont(oPDF,"ARIAL",8,8)
	Call GF_setFontColor("#FFFFFF")
	Call GF_writeTextAlign(oPDF, COMIENZO_COL_ID , y_comienzo+2, "ID" , 25 , PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, COMIENZO_COL_DS , y_comienzo+2, "DESCRIPCION" , 230 , PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, COMIENZO_COL_EXISTENCIA , y_comienzo+2, "EXIST." , 45 , PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, COMIENZO_COL_SOBRANTE , y_comienzo+2, "SOBRANTE" , 50 , PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, COMIENZO_COL_TOTAL , y_comienzo+2, "TOTAL" , 45 , PDF_ALIGN_CENTER)
	if (valorizar) then
		Call GF_writeTextAlign(oPDF, COMIENZO_COL_VALORACION , y_comienzo+2, "VALORACION" , 70 , PDF_ALIGN_CENTER)
	end if
	Call GF_writeTextAlign(oPDF, COMIENZO_COL_CODIGO , y_comienzo+2, "UBICACION" , 580-COMIENZO_COL_CODIGO , PDF_ALIGN_CENTER)
	Call GF_setFontColor("#000000")
	Call GF_setFont(oPDF,"ARIAL",8,0)
	Call GF_writeTextAlign(oPDF, 10 , 840, "Pagina  " & nroPagina		 , 580 , PDF_ALIGN_RIGHT)
	
	dibujarPagina = y_comienzo +16
End Function
'-----------------------------------------------------------------------------------------------
Function cargarDatos(p_ultimaLinea)
	Dim y_inicial,x_inicial,i,existencia,sobrante,total,y,separacion,cdarticulo,color
	Dim descripcion,valoracion, strSQL, conn, rs

	strSQL = getSQLStock()
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)

	if (not rs.eof) then
		y_inicial = p_ultimaLinea
		x_inicial = 10
		i=0
		while not rs.eof
			if (isnull(rs("existencia")))then existencia = 0 else existencia = cdbl(rs("existencia"))  end if
			if (isnull(rs("sobrante")))  then sobrante   = "0" else sobrante = rs("sobrante")  end if
			if (isnull(rs("total")))     then total      = "0" else total = rs("total")  end if
			if (isnull(rs("cdinterno"))) then cdarticulo = ""  else cdarticulo = rs("cdinterno") end if
			if (len(rs("dsarticulo"))>40) then descripcion = left(rs("dsarticulo"),40) & "..." else descripcion = rs("dsarticulo") end if
			
			if ((valorizar) and (not isNull(rs("vlupesos")))) then
				valoracion = GF_EDIT_DECIMALS(existencia * CDbl(rs("vlupesos")) ,2)				
			else
				valoracion = GF_EDIT_DECIMALS("000",2)
			end if
			Call GF_setFont(oPDF,"COURIER",TEXT_SIZE,0)
	
			separacion = 3
			y = y_inicial+(i*(pdf_currentFontSize+separacion))
			
			if i mod 2 then 
				color = "#dcf7dc"			
			else
				color = "#FFFFFF"
			end if
				
			Call GF_squareBox(oPDF, COMIENZO_COL_ID , y-1 ,565, 13, 0, color, color, 1, PDF_SQUARE_NORMAL)
			Call GF_writeTextAlign(oPDF,COMIENZO_COL_ID , y, rs("idArticulo")     ,23 , PDF_ALIGN_RIGHT)
			Call GF_writeTextAlign(oPDF,COMIENZO_COL_DS+5 , y, descripcion , 400, PDF_ALIGN_LEFT)
			Call GF_writeTextAlign(oPDF,COMIENZO_COL_EXISTENCIA, y, existencia 		 , 30 , PDF_ALIGN_RIGHT)
			Call GF_writeTextAlign(oPDF,COMIENZO_COL_SOBRANTE, y, sobrante   		 , 35 , PDF_ALIGN_RIGHT)
			Call GF_writeTextAlign(oPDF,COMIENZO_COL_TOTAL, y, total      		 , 30 , PDF_ALIGN_RIGHT)
			Call GF_writeTextAlign(oPDF,COMIENZO_COL_EXISTENCIA+35, y, rs("unidad")  , 15 , PDF_ALIGN_LEFT)
			Call GF_writeTextAlign(oPDF,COMIENZO_COL_SOBRANTE+40, y, rs("unidad")	 , 15 , PDF_ALIGN_LEFT)
			Call GF_writeTextAlign(oPDF,COMIENZO_COL_TOTAL+35, y, rs("unidad")  , 15 , PDF_ALIGN_LEFT)
			if (valorizar) then
				Call GF_writeTextAlign(oPDF,COMIENZO_COL_VALORACION+10, y,"$" , 5 , PDF_ALIGN_RIGHT)
				Call GF_writeTextAlign(oPDF,COMIENZO_COL_VALORACION, y, valoracion , 65 , PDF_ALIGN_RIGHT)
			end if
			Call GF_setFont(oPDF,"COURIER",6,0)
			Call GF_writeTextAlign(oPDF,COMIENZO_COL_CODIGO+5, y, cdarticulo 		 , 570-COMIENZO_COL_CODIGO , PDF_ALIGN_LEFT)
	
			Call GF_setFont(oPDF,"COURIER",8,0)
			if (y_inicial+(i*(pdf_currentFontSize+separacion)) >= 800) then
				y_inicial = nuevaPagina()
				i = -1
			end if
			i = i + 1
			rs.movenext
		wend
	else
		Call GF_squareBox(oPDF, COMIENZO_COL_ID, p_ultimaLinea-16, 560, 15, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
		Call GF_setFont(oPDF,"ARIAL",8,0)
		Call GF_setFontColor("#FFFFFF")
		
		Call GF_writeTextAlign(oPDF, COMIENZO_COL_ID , p_ultimaLinea-13, "NO SE ENCONTRARON DATOS EN LA BUSQUEDA" , 560 , PDF_ALIGN_CENTER)
	end if
End Function
'-----------------------------------------------------------------------------------------------
Function crearPDF()
	Dim ultimaLinea
	Set oPDF = GF_createPDF("PDFTemp")
	Call GF_setPDFMODE(PDF_STREAM_MODE)
	nroPagina = 0
	ultimaLinea = nuevaPagina()	 
	Call cargarDatos(ultimaLinea)
	Call GF_closePDF(oPDF)
End Function
'-----------------------------------------------------------------------------------------------
Function getSQLStock()
	dim mywhere, strSQL
	strSQL = ""

	if (metodo = FILTRO_EXISTENTES) then mywhere = " WHERE (cie.existencia <> 0 OR cie.sobrante <> 0)"
	
	Call mkWhere(mywhere, "cie.idalmacen", almacen,"=", 1)

	strSQL = "SELECT			cie.idarticulo                          , "
	strSQL = strSQL & "         art.dsarticulo				            , "
	strSQL = strSQL & "         cat.idcategoria                         , "
	strSQL = strSQL & "         cat.dscategoria                         , "
	strSQL = strSQL & "         cie.existencia					        , "
	strSQL = strSQL & "         uni.abreviatura unidad                  , "
	strSQL = strSQL & "         cie.sobrante							, "
	strSQL = strSQL & "         cie.existencia + cie.sobrante total		, "
	strSQL = strSQL & "         cie.vlupesos						    , "
	strSQL = strSQL & "         cie.vludolares							, "
	strSQL = strSQL & "         ard.cdinterno "
	strSQL = strSQL & "FROM     TBLREPORTESTOCKWF cie "
	strSQL = strSQL & "         INNER JOIN TBLARTICULOS art "
	strSQL = strSQL & "				ON	cie.idarticulo  = art.idarticulo "
	strSQL = strSQL & "				AND cie.cdusuario = '" & session("Usuario") & "'"
	strSQL = strSQL & "			INNER JOIN TBLARTCATEGORIAS cat "
	strSQL = strSQL & "				ON  art.idcategoria = cat.idcategoria "
	strSQL = strSQL & "         INNER JOIN TBLUNIDADES uni "
	strSQL = strSQL & "				ON  art.idunidad = uni.idunidad "
	strSQL = strSQL & "         LEFT JOIN TBLARTICULOSDATOS ard "
	strSQL = strSQL & "				ON  ard.idArticulo = cie.idArticulo "
	strSQL = strSQL & "				AND ard.idAlmacen = " & almacen
	strSQL = strSQL & mywhere
	strSQL = strSQL & "         ORDER BY art.idarticulo"	
	getSQLStock = strSQL
	'Response.Write strSQL
	'Response.End 
End Function
'-----------------------------------------------------------------------------------------------
'******************************************************************************************
'******************************************************
'					INICIO DE LA PAGINA
'******************************************************
accion    = GF_PARAMETROS7("accion"   , "", 6)
metodo    = GF_PARAMETROS7("metodo"   , 0 , 6)
categoria = GF_PARAMETROS7("categoria", 0 , 6)
almacen   = GF_PARAMETROS7("almacen"  , 0 , 6)
valorizar = GF_PARAMETROS7("valorizar", "", 6)
fechaBusqueda = GF_PARAMETROS7("fechaBusqueda", "", 6)
incluir = GF_PARAMETROS7("incluir", "", 6)
fechaBusqueda = GF_DTE2FN(fechaBusqueda)

if (valorizar = "on") then valorizar = true

if (valorizar) then
	COMIENZO_COL_CODIGO	= 475
else
	COMIENZO_COL_CODIGO	= 405
end if
if (accion = ACCION_PROCESAR) then
	Call crearPDF()	
end if
%>