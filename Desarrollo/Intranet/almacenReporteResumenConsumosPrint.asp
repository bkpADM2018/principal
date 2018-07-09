<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosPDF.asp"-->
<!--#include file="Includes/procedimientosLog.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<%
dim oPDF, rs, conn, strSQL, nroPagina, RPT_accion, ultimaLinea
dim RPT_Division, RPT_Month, RPT_Year, RPT_Filtro, RPT_Generando

Const TIPO_CATEGORIA = 1
Const TIPO_PART_PRES = 2
Const TIPO_TODOS = 3

Const PAGINA_CABECERA  = 0
Const PAGINA_SIGUIENTE = 1

Const LINEA_POST_ENCABEZADO = 75

Const X_COL_CUENTA  = 20
Const X_COL_CCOSTO  = 80
Const X_COL_CAT_OBR = 110
Const X_COL_TOTAL_P = 390
Const X_COL_TOTAL_UD= 480

Const ANCHO_COL_CUENTA  = 60
Const ANCHO_COL_CCOSTO  = 30
Const ANCHO_COL_CAT_OBR = 280
Const ANCHO_COL_TOTAL_P = 90
Const ANCHO_COL_TOTAL_UD= 90
Const ANCHO_TOTAL       = 550

Const MAX_Y_PAGINA      = 760

Const LARGO_TITULOS = 13

Const SEPARACION = 10
Const SEPARACION_OBRAS = 5

'-----------------------------------------------------------------------------------------
Function armadoPDF(oPDF)
	Call dibujarPagina(oPDF, PAGINA_CABECERA)
	if (RPT_Generando = TIPO_CATEGORIA) then
		Call writeDatosCategorias(oPDF)
	else
		Call writeDatosObras(oPDF)
	end if
End Function
'-----------------------------------------------------------------------------------------
Function dibujarPagina(oPDF, pTipoPagina)
	ultimaLinea = dibujarEncabezado(oPDF)
	if (pTipoPagina = PAGINA_CABECERA) then ultimaLinea = dibujarIndice(oPDF)
	ultimaLinea = dibujarTitulosCol(oPDF)
End Function
'-----------------------------------------------------------------------------------------
'Devuelve la posicion donde se puede seguir escribiendo
Function dibujarEncabezado(oPDF)
	Dim titulo

	Call GF_squareBox(oPDF, 2, 2, 590, 833, 0, "", "#0B3B0B", 2, PDF_SQUARE_ROUND)
	Call GF_writeImage(oPDF, Server.MapPath("images\kogge64.gif"), 20, 10, 48, 48, 0)

	if (RPT_Generando = TIPO_CATEGORIA) then
		titulo = "RESUMEN DE CONSUMOS POR CATEGORIAS"
	else
		titulo = "RESUMEN DE CONSUMOS POR PART. PRESUPUESTARIA"
	end if

	Call GF_setFont(oPDF,"ARIAL",16,FONT_STYLE_BOLD)
	Call GF_writeTextAlign(oPDF,20,25,GF_TRADUCIR(titulo), 550 , PDF_ALIGN_CENTER)
	
	GP_CONFIGURARMOMENTOS
	Call GF_setFont(oPDF,"COURIER",8,FONT_STYLE_NORMAL)
	Call GF_writeTextAlign(oPDF,5,5,GF_FN2DTE(session("MmtoSistema")), 580 , PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,5,5+pdf_currentFontSize,session("Usuario"), 580 , PDF_ALIGN_RIGHT)
	
	Call GF_writeTextAlign(oPDF, 10 , 840, "Pagina  " & nroPagina		 , 580 , PDF_ALIGN_RIGHT)
	Call GF_horizontalLine(oPDF,2,65,590)

	dibujarEncabezado = LINEA_POST_ENCABEZADO
End Function
'-----------------------------------------------------------------------------------------
Function dibujarIndice(oPDF)
	dim parametros(2), division, conn, rsDivision, strSQL

	if (RPT_Division <> 0) then 
		strSQL = "Select * from TBLDIVISIONES where IDDIVISION=" & RPT_Division
		Call executeQueryDB(DBSITE_SQL_INTRA, rsDivision, "OPEN", strSQL)
		if (not rsDivision.eof) then division = rsDivision("DSDIVISION")
	end if
	parametros(0)= "Division" & "," & division
	parametros(1)= "Mes de" & "," & GF_INT2MES(RPT_Month)
	parametros(2)= "Año" & "," & RPT_Year
	dibujarIndice = dibujarFiltros(oPDF, parametros)
End Function
'-----------------------------------------------------------------------------------------
'Funcion que dibuja los parametros de busqueda 
'recibe como parametro un vector con la siguiente estructura en cada posicion:
'	"nombreBusqueda,valorBusqueda"
'Devuelve la posicion donde se puede seguir escribiendo
Function dibujarFiltros(oPDF, p_parametros)
	Dim aux,x_inicial,y_inicial,font_size
	Dim max_len, my_font_size
	
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
		Call GF_writeTextAlign(oPDF,x_inicial,y_inicial+(i*my_font_size), completarEspacios(aux(0),max_len, ".") & ": " & aux(1) ,580, PDF_ALIGN_LEFT)
	next 
	
	dibujarFiltros = y_inicial+(i*my_font_size)+my_font_size
End Function
'-----------------------------------------------------------------------------------------
Function completarEspacios(p_palabra,p_len, p_agregado)
	Dim rtrn
	rtrn = p_palabra
	for i = len(p_palabra) to p_len
		if (p_agregado = ".") then
			rtrn = rtrn & p_agregado
		else
			rtrn = p_agregado & rtrn
		end if
	next
	completarEspacios = rtrn
End Function
'-----------------------------------------------------------------------------------------
Function dibujarTitulosCol(oPDF)

	Call GF_squareBox(oPDF, X_COL_CUENTA, ultimaLinea, ANCHO_COL_CUENTA, LARGO_TITULOS, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, X_COL_CCOSTO, ultimaLinea, ANCHO_COL_CAT_OBR + ANCHO_COL_CCOSTO, LARGO_TITULOS, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	if (RPT_Generando <> TIPO_CATEGORIA) then	Call GF_squareBox(oPDF, X_COL_CCOSTO, ultimaLinea, ANCHO_COL_CCOSTO, LARGO_TITULOS, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, X_COL_TOTAL_P, ultimaLinea, ANCHO_COL_TOTAL_P, LARGO_TITULOS, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, X_COL_TOTAL_UD, ultimaLinea, ANCHO_COL_TOTAL_UD, LARGO_TITULOS, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	
	Call GF_setFont(oPDF,"ARIAL",9,8)
	Call GF_setFontColor("#FFFFFF")
	if (RPT_Generando = TIPO_CATEGORIA) then
		Call GF_writeTextAlign(oPDF, X_COL_CUENTA , ultimaLinea+2, GF_TRADUCIR("CUENTA") , ANCHO_COL_CUENTA, PDF_ALIGN_CENTER)
		Call GF_writeTextAlign(oPDF, X_COL_CCOSTO + 5 , ultimaLinea+2, GF_TRADUCIR("DESCRIPCIÓN") , ANCHO_COL_CAT_OBR + ANCHO_COL_CCOSTO, PDF_ALIGN_LEFT)
	else
		Call GF_writeTextAlign(oPDF, X_COL_CUENTA , ultimaLinea+2, GF_TRADUCIR("CUENTA") , ANCHO_COL_CUENTA, PDF_ALIGN_CENTER)
		Call GF_writeTextAlign(oPDF, X_COL_CCOSTO , ultimaLinea+2, GF_TRADUCIR("C. C.") , ANCHO_COL_CCOSTO , PDF_ALIGN_CENTER)
		Call GF_writeTextAlign(oPDF, X_COL_CAT_OBR + 5 , ultimaLinea+2, GF_TRADUCIR("DESCRIPCIÓN") , ANCHO_COL_CAT_OBR , PDF_ALIGN_LEFT)
	end if
	Call GF_writeTextAlign(oPDF, X_COL_TOTAL_P , ultimaLinea+2, GF_TRADUCIR("TOTAL $") , ANCHO_COL_TOTAL_P , PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, X_COL_TOTAL_UD , ultimaLinea+2, GF_TRADUCIR("TOTAL U$S") , ANCHO_COL_TOTAL_UD , PDF_ALIGN_CENTER)

	Call GF_setFontColor("#000000")
	dibujarTitulosCol = ultimaLinea + LARGO_TITULOS + SEPARACION_OBRAS

End Function
'-----------------------------------------------------------------------------------------
Function writeDatosCategorias(oPDF)
	dim cdcategoria, descripcion, ccostos, fcolor, cdCuenta
	dim strSQL, conn, rsCategorias, contador, regTotal
	dim pesos, dolares, TotalPesos, TotalDolares
	
	pesos = 0
	dolares = 0
	Call getSQLCategorias(strSQL)
	call executeQueryDb(DBSITE_SQL_INTRA, rsCategorias, "OPEN", strSQL)
	if (not rsCategorias.eof) then
		totalRs = rsCategorias.RecordCount
		While (not rsCategorias.eof)
			contRs = contRs + 1
			if (contRs mod 2) then	fcolor = "#CECEF6"	else	fcolor = "#FFFFFF"	end if
			descripcion = rsCategorias("CDCATEGORIA") & "-" & rsCategorias("DSCATEGORIA") 
			cdCuenta = rsCategorias("CDCUENTA") 
			ccostos = " "
			pesos = cDbl(rsCategorias("TOTALPESOS"))
			dolares = cDbl(rsCategorias("TOTALDOLARES"))
			Call writeRegistro(oPDF, ultimaLinea,  cdCuenta, descripcion, ccostos, pesos, dolares, fcolor)
			ultimaLinea = ultimaLinea + SEPARACION
			if ((ultimaLinea >= MAX_Y_PAGINA) and (contRs < totalRs)) then
				Call nuevaPagina(oPDF, PAGINA_SIGUIENTE)
			end if
			TotalPesos = TotalPesos + pesos
			TotalDolares = TotalDolares + dolares
			rsCategorias.MoveNext
		Wend
		ultimaLinea = writeTotalesCategorias(oPDF, ultimaLinea, TotalPesos, TotalDolares)
	else
		Call GF_setFontColor("#FFFFFF")
		Call GF_squareBox(oPDF, X_COL_CUENTA, ultimaLinea, ANCHO_TOTAL, LARGO_TITULOS, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
		Call GF_writeTextAlign(oPDF, X_COL_CUENTA , ultimaLinea+2, GF_TRADUCIR("NO SE ENCONTRARON DATOS EN LA BUSQUEDA") , ANCHO_TOTAL, PDF_ALIGN_CENTER)
		Call GF_setFontColor("#000000")
	end if
	'pasa al resumen por obras si se requirio
	if (RPT_Filtro = TIPO_TODOS) then	Call armarPDFObras(oPDF)
End Function
'-----------------------------------------------------------------------------------------
Function writeDatosObras(oPDF)
	dim cuenta, descripcion, ccostos, pesos, dolares, fcolor
	dim strSQL, conn, rsObras, contRs, totalRs, TotalPesos, TotalDolares
	dim idObra, actualObra, idArea, actualArea, dsdetalle
	dim TotalObraPesos, TotalObraDolar, cambioPagina
	
	cambioPagina = false
	pesos = 0
	dolares = 0
	Call getSQLObras(strSQL)
	Call executeQueryDB(DBSITE_SQL_INTRA, rsObras, "OPEN", strSQL)
	if (not rsObras.eof) then
		totalRs = rsObras.RecordCount
		While (not rsObras.eof)
			contRs = contRs + 1
			idObra = rsObras("IDOBRA")
			idArea = rsObras("IDAREA")
			idSector = rsObras("IDSECTOR")
			dsSector = rsObras("DSSECTOR")
			if (idObra <> actualObra) then
				if (contRs > 1) then 
					if actualObra = 0 then
						ultimaLinea = writeTotalesObra(oPDF, "TOTAL SIN PARTIDA: ", ultimaLinea, TotalObraPesos, TotalObraDolar)
					else
						ultimaLinea = writeTotalesObra(oPDF, "TOTAL DE LA PART. PRES.: ", ultimaLinea, TotalObraPesos, TotalObraDolar)
					end if	
				end if	
				ultimaLinea = WriteCabeceraObra(oPDF, ultimaLinea, idObra)
				TotalObraPesos = 0
				TotalObraDolar = 0
			else
				if ((cambioPagina) or ((idObra = 0) and (contRs = 1))) then ultimaLinea = WriteCabeceraObra(oPDF, ultimaLinea, idObra)
			end if
			if ((idArea <> actualArea) or (cambioPagina) or ((idArea = 0) and (TotalObraPesos = 0))) then
				ultimaLinea = WriteAreaObra(oPDF, ultimaLinea, idObra, idArea)
			end if
			
			if idObra <> 0 or idArea<>0 then
				cuenta = rsObras("CDCUENTA")
				descripcion = completarEspacios(rsObras("IDDETALLE"),3," ") & " - "
				dsdetalle = rsObras("DSBUDGET")
				if ((dsdetalle = "") or (isnull(dsdetalle))) then dsdetalle = "SIN DETALLE"
				descripcion = descripcion & dsdetalle
				ccostos = " "
				if not (isInversion(idObra)) then ccostos = rsObras("CCOSTOS")
			else
				if idSector<>0 then 
					ccostos = rsObras("CCOSTOSSECTOR")
					cuenta = rsObras("CDCUENTASECTOR")
					
					descripcion = completarEspacios(idSector,3," ") & " - "
					dsdetalle = dsSector
					if ((dsdetalle = "") or (isnull(dsdetalle))) then dsdetalle = "SIN DETALLE"
					descripcion = descripcion & dsdetalle					
				else 'Es un AJS
					cuenta = CUENTA_AJUSTE_STOCK
					descripcion = "   0 - AJUSTES DE STOCK"
				end if	
			end if			
			
			

			pesos = cDbl(rsObras("TOTALPESOS"))
			dolares = cDbl(rsObras("TOTALDOLARES"))
			Call writeRegistro(oPDF, ultimaLinea, trim(cuenta), descripcion, trim(ccostos), pesos, dolares, "")
			ultimaLinea = ultimaLinea + SEPARACION
			if ((ultimaLinea >= MAX_Y_PAGINA) and (contRs < totalRs)) then
				Call nuevaPagina(oPDF, PAGINA_SIGUIENTE)
				cambioPagina = true
			else
				cambioPagina = false
			end if
			TotalObraPesos = TotalObraPesos + pesos
			TotalObraDolar = TotalObraDolar + dolares
			TotalPesos = TotalPesos + pesos
			TotalDolares = TotalDolares + dolares
			actualObra = idObra
			actualArea = idArea
			rsObras.MoveNext
		Wend
		ultimaLinea = writeTotalesObra(oPDF, "TOTAL DE LA PART. PRES.: ", ultimaLinea, TotalObraPesos, TotalObraDolar)
	else
		Call GF_setFontColor("#FFFFFF")
		Call GF_squareBox(oPDF, X_COL_CUENTA, ultimaLinea, ANCHO_TOTAL, LARGO_TITULOS, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
		Call GF_writeTextAlign(oPDF, X_COL_CUENTA , ultimaLinea+2, GF_TRADUCIR("NO SE ENCONTRARON DATOS EN LA BUSQUEDA") , ANCHO_TOTAL, PDF_ALIGN_CENTER)
		Call GF_setFontColor("#000000")
	end if
	if (contRs > 0) then
		ultimaLinea = ultimaLinea + (SEPARACION * 2)
		ultimaLinea = writeTotalesObra(oPDF, "TOTAL DEL RESUMEN DE CONSUMOS POR PART. PRES.: ", ultimaLinea, TotalPesos, TotalDolares)
	end if
End Function
'-----------------------------------------------------------------------------------------
Function writeRegistro(oPDF, pY, cuenta, descripcion, ccostos, Pesos, Dolares, fcolor)
	if (fcolor <> "") then Call GF_squareBox(oPDF, X_COL_CUENTA, pY, ANCHO_TOTAL, SEPARACION, 0, fcolor, fcolor, 0.5, PDF_SQUARE_NORMAL)
	Call GF_setFont(oPDF,"COURIER",8,0)
	Call GF_setFontColor("#000000")
	if (cuenta <> "") then Call GF_writeTextAlign(oPDF, X_COL_CUENTA , pY+1, cuenta , ANCHO_COL_CUENTA, PDF_ALIGN_CENTER)
	if (RPT_Generando = TIPO_CATEGORIA) then
		if (len(descripcion) > 65) then descripcion = left(descripcion, 65) & "..."
		if (descripcion <> "") then Call GF_writeTextAlign(oPDF, X_COL_CCOSTO + 5 , pY+1, descripcion , ANCHO_COL_CAT_OBR + ANCHO_COL_CCOSTO , PDF_ALIGN_LEFT)
	else
		if (len(descripcion) > 48) then descripcion = left(descripcion, 48) & "..."
		if (descripcion <> "") then Call GF_writeTextAlign(oPDF, X_COL_CAT_OBR + 5 , pY+1, uCase(descripcion) , ANCHO_COL_CAT_OBR , PDF_ALIGN_LEFT)
		if (ccostos <> "") then Call GF_writeTextAlign(oPDF, X_COL_CCOSTO , pY+1, ccostos , ANCHO_COL_CCOSTO , PDF_ALIGN_CENTER)
	end if
	if (Pesos <= 0) then	Call GF_setFontColor("#DF0101")
	Call GF_writeTextAlign(oPDF, X_COL_TOTAL_P , pY+1, "$ " & GF_EDIT_DECIMALS(round(Pesos,0),2) , ANCHO_COL_TOTAL_P - 5 , PDF_ALIGN_RIGHT)
	if (Dolares <= 0) then	Call GF_setFontColor("#DF0101")
	Call GF_writeTextAlign(oPDF, X_COL_TOTAL_UD , pY+1, "U$S " & GF_EDIT_DECIMALS(round(Dolares,0),2) , ANCHO_COL_TOTAL_UD - 5 , PDF_ALIGN_RIGHT)
	Call GF_setFontColor("#000000")
End Function
'-----------------------------------------------------------------------------------------
Function armarPDFObras(oPDF)
	RPT_Generando = TIPO_PART_PRES
	Call nuevaPagina(oPDF, PAGINA_CABECERA)
	Call writeDatosObras(oPDF)
End Function
'-----------------------------------------------------------------------------------------
Function nuevaPagina(oPDF, tipoPag)
	Call GF_newPage(oPDF)
	nroPagina = nroPagina + 1
	Call dibujarPagina(oPDF, tipoPag)
End Function
'-----------------------------------------------------------------------------------------
Function getSQLCategorias(ByRef strSQL)
	dim almacenes, fecha, fechaDesde, fechaHasta

	if (RPT_Month < 10) then RPT_Month = "0" & right(RPT_Month,1)
	fecha = RPT_Year & RPT_Month
	fechaDesde = cDbl(fecha & "01" & "000000")
	fechaHasta = cDbl(fecha & "31" & "235959")
	almacenes = getAlmacenesPorDivision(RPT_Division)

	strSQL = ""
	strSQL = "SELECT " &_
			 "	       cat.cdcategoria, " &_
			 "	       cat.dscategoria, " &_
			 "	       cat.cdcuenta, " &_
			 "	       cat.ccostos, " &_
			 "	       SUM(vd.existencia * vd.vlupesos)   AS totalpesos, " &_
			 "	       SUM(vd.existencia * vd.vludolares) AS totaldolares " &_
			 "	FROM   tblvalescabecera vc " &_
			 "	       INNER JOIN tblvalesdetalle vd ON vc.idvale = vd.idvale " &_			 
			 "	       INNER JOIN tblarticulos art " &_
			 "	         ON vd.idarticulo = art.idarticulo " &_
			 "	       INNER JOIN tblartcategorias cat " &_
			 "	         ON art.idcategoria = cat.idcategoria " &_
			 "	WHERE      vc.idalmacen IN ( " & almacenes & " ) " &_
			 "	       AND vc.fecha LIKE '" & fecha & "%' " &_
			 "	       AND vc.estado = 1 " &_
			 "	       AND vc.cdvale IN ('" & CODIGO_VS_SALIDA & "','" & CODIGO_VS_AJUSTE_VALE & "','" & CODIGO_VS_AJUSTE_STOCK & "') " &_
			 "	       AND vd.existencia <> 0 " &_
			 "	GROUP  BY cat.cdcategoria, cat.dscategoria, cat.ccostos, cat.cdcuenta " &_
			 "	ORDER  BY cat.cdcategoria"
End Function
'-----------------------------------------------------------------------------------------
Function getSQLObras(ByRef strSQL)
	dim almacenes, fecha, fechaDesde, fechaHasta

	if (RPT_Month < 10) then RPT_Month = "0" & right(RPT_Month,1)
	fecha = RPT_Year & RPT_Month
	fechaDesde = cDbl(fecha & "01" & "000000")
	fechaHasta = cDbl(fecha & "31" & "235959")
	almacenes = getAlmacenesPorDivision(RPT_Division)

	strSQL = ""
	strSQL = "SELECT " &_
			 "	       tg.idobra, " &_
			 "	       tg.dsbudget, " &_
			 "	       tg.idbudgetarea as idarea, " &_
			 "	       tg.idbudgetdetalle as iddetalle, " &_
			 "	       tg.idsector, " &_
			 "	       tg.dssector, " &_
			 "	       tg.cdcuentaSector, " &_
			 "	       tg.ccostosSector, " &_
			 "	       tg.cdcuenta, " &_
			 "	       tg.ccostos, " &_
			 "	       tg.totalpesos, " &_
			 "	       tg.totaldolares " &_
			 "	FROM   (SELECT t1.idobra, " &_
			 "	               t1.idbudgetarea, " &_
			 "	               t1.idbudgetdetalle, " &_
 			 "	               t1.idsector, " &_
 			 "	               t1.dssector, " &_
 			 "	               t1.cdcuentaSector, " &_
 			 "	               t1.ccostosSector, " &_
			 "	               t1.dsbudget, " &_
			 "	               t1.cdcuenta, " &_
			 "	               t1.ccostos, " &_
			 "	               SUM(totalpesos)   AS totalpesos, " &_
			 "	               SUM(totaldolares) AS totaldolares " &_
			 "	        FROM   (SELECT vc.idobra, " &_
			 "	                       vc.idbudgetarea, " &_
			 "	                       vc.idbudgetdetalle, " &_
			 "	                       vc.idsector, " &_	
	 		 "	                       sec.dssector, " &_	
	 		 "	                       sec.cdcuenta as cdCuentaSector, " &_	
	 		 "	                       sec.ccostos as ccostosSector, " &_	
			 "	                       bo.dsbudget, " &_
			 "	                       bo.cdcuenta, " &_
			 "	                       bo.ccostos, " &_
			 "	                       SUM(vd.existencia * vd.vlupesos)   AS totalpesos, " &_
			 "	                       SUM(vd.existencia * vd.vludolares) AS totaldolares " &_
			 "	                FROM   tblvalescabecera vc " &_
			 "	                       INNER JOIN tblvalesdetalle vd " &_
			 "	                         ON vc.idvale = vd.idvale " &_
			 "	                       LEFT JOIN TBLSECTORES SEC " &_
			 "	                         ON vc.idsector = sec.idsector " &_
			 "	                       LEFT JOIN tblbudgetobras bo " &_
			 "	                         ON bo.idobra = vc.idobra " &_
			 "	                            AND bo.idarea = vc.idbudgetarea " &_
			 "	                            AND bo.iddetalle = vc.idbudgetdetalle " &_
			 "	                WHERE      vc.idalmacen IN ( " & almacenes & " ) " &_
			 "	                       AND vc.fecha LIKE '" & fecha & "%' " &_
			 "	                       AND vc.estado = 1 " &_
			 "	                       AND vc.cdvale IN ('" & CODIGO_VS_SALIDA & "','" & CODIGO_VS_AJUSTE_VALE & "','" & CODIGO_VS_AJUSTE_STOCK & "') " &_
			 "	                       AND vd.existencia <> 0 " &_
			 "	                GROUP  BY vc.idobra, " &_
			 "	                          vc.idbudgetarea, " &_
			 "	                          vc.idbudgetdetalle, " &_
			 "	                          vc.idsector, " &_
			 "	                          sec.dssector, " &_
			 "	                          sec.cdcuenta, " &_
			 "	                          sec.ccostos, " &_
			 "	                          bo.dsbudget, " &_
			 "	                          bo.cdcuenta, " &_
			 "	                          bo.ccostos)t1 " &_
			 "	        GROUP  BY t1.idobra, " &_
			 "	                  t1.idbudgetarea, " &_
			 "	                  t1.idbudgetdetalle, " &_
			 "	                  t1.idsector, " &_
			 "	                  t1.dssector, " &_
			 "	                  t1.cdCuentaSector, " &_
			 "	                  t1.ccostosSector, " &_
			 "	                  t1.dsbudget, " &_
			 "	                  t1.cdcuenta, " &_
			 "	                  t1.ccostos " &_
			 "	        )tg " &_
			 "	ORDER  BY tg.idobra, idarea, iddetalle, idsector"
 
End Function
'-----------------------------------------------------------------------------------------
Function WriteCabeceraObra(oPDF, p_y, id)
	dim tituloObra, ObraCD, ObraDS
	p_y = p_y + SEPARACION_OBRAS
	Call loadDatosObra(id, ObraCD, ObraDS, "", "", "", "", "", "", "", "", "", "")
	tituloObra = ObraCD & " - " & ObraDS
	if (id = 0) then
		tituloObra = "SIN PARTIDA"
	else
		if (isInversion(id)) then
			tituloObra = "OBRA DE INVERSION: " & tituloObra
		else
			tituloObra = "OBRA DE MANTENIMIENTO: " & tituloObra
		end if
	end if
	Call GF_setFont(oPDF,"ARIAL",9,8)
	Call GF_squareBox(oPDF, X_COL_CUENTA, p_y, ANCHO_TOTAL, LARGO_TITULOS, 0, "#D8D8D8", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_writeTextAlign(oPDF, X_COL_CUENTA + 5 , p_y+1, tituloObra, ANCHO_TOTAL , PDF_ALIGN_LEFT)
	WriteCabeceraObra = p_y + LARGO_TITULOS + SEPARACION_OBRAS
End Function
'-----------------------------------------------------------------------------------------
Function writeTotalesObra(oPDF, titulo, p_y, TotalPesos, TotalDolares)
	p_y = p_y + SEPARACION_OBRAS
	Call GF_squareBox(oPDF, X_COL_CUENTA, p_y, ANCHO_TOTAL, LARGO_TITULOS, 0, "#D8D8D8", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_verticalLine(oPDF, X_COL_TOTAL_P, p_y, SEPARACION)
	Call GF_verticalLine(oPDF, X_COL_TOTAL_UD, p_y, SEPARACION)
	Call GF_setFont(oPDF,"ARIAL",9,8)
	Call GF_writeTextAlign(oPDF, X_COL_CAT_OBR , p_y+1, titulo, ANCHO_COL_CAT_OBR - 5, PDF_ALIGN_RIGHT)
	Call GF_setFont(oPDF,"COURIER",9,8)
	if ((TotalPesos <= 0) or (TotalDolares <= 0)) then	Call GF_setFontColor("#DF0101")
	Call GF_writeTextAlign(oPDF, X_COL_TOTAL_P , p_y+1, "$ " & GF_EDIT_DECIMALS(round(TotalPesos,0),2), ANCHO_COL_TOTAL_P - 5 , PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF, X_COL_TOTAL_UD , p_y+1, "U$S " & GF_EDIT_DECIMALS(round(TotalDolares,0),2), ANCHO_COL_TOTAL_UD - 5 , PDF_ALIGN_RIGHT)
	Call GF_setFontColor("#000000")
	writeTotalesObra = p_y + LARGO_TITULOS + SEPARACION_OBRAS
End Function
'-----------------------------------------------------------------------------------------
Function WriteAreaObra(oPDF, p_y, idobra, idArea)
	dim tituloArea, rs, strSQL, conn
	p_y = p_y + SEPARACION_OBRAS
	tituloArea = "SIN AREA"
	if idObra <> 0 or idArea<>0 then
		strSQL = "select DSBUDGET from TBLBUDGETOBRAS where IDOBRA=" & idobra & " and IDAREA="&idArea & " and IDDETALLE=0"
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if (not rs.eof) then tituloArea = idArea & " " & rs("DSBUDGET")
	else
		tituloArea = "VARIOS"
	end if
	Call GF_setFont(oPDF,"ARIAL",9,8)
	Call GF_squareBox(oPDF, X_COL_CUENTA, p_y, ANCHO_TOTAL, SEPARACION, 0, "#CECEF6", "", 0, PDF_SQUARE_NORMAL)
	Call GF_writeTextAlign(oPDF, X_COL_CUENTA + 5 , p_y+1, tituloArea, ANCHO_TOTAL , PDF_ALIGN_LEFT)
	WriteAreaObra = p_y + SEPARACION
End Function
'-----------------------------------------------------------------------------------------
Function writeTotalesCategorias(oPDF, p_y, TotalPesos, TotalDolares)
	p_y = p_y + SEPARACION_OBRAS
	Call GF_squareBox(oPDF, X_COL_CUENTA, p_y, ANCHO_TOTAL, LARGO_TITULOS, 0, "#D8D8D8", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_verticalLine(oPDF, X_COL_TOTAL_P, p_y, SEPARACION)
	Call GF_verticalLine(oPDF, X_COL_TOTAL_UD, p_y, SEPARACION)
	Call GF_setFont(oPDF,"ARIAL",9,8)
	Call GF_writeTextAlign(oPDF, X_COL_CAT_OBR , p_y+1, "TOTAL DEL RESUMEN DE CONSUMOS POR CATEGORIAS: ", ANCHO_COL_CAT_OBR - 5, PDF_ALIGN_RIGHT)
	Call GF_setFont(oPDF,"COURIER",9,8)
	if ((TotalPesos <= 0) or (TotalDolares <= 0)) then	Call GF_setFontColor("#DF0101")
	Call GF_writeTextAlign(oPDF, X_COL_TOTAL_P , p_y+1, "$ " & GF_EDIT_DECIMALS(round(TotalPesos,0),2), ANCHO_COL_TOTAL_P - 5 , PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF, X_COL_TOTAL_UD , p_y+1, "U$S " & GF_EDIT_DECIMALS(round(TotalDolares,0),2), ANCHO_COL_TOTAL_UD - 5 , PDF_ALIGN_RIGHT)
	Call GF_setFontColor("#000000")
	writeTotalesCategorias = p_y + LARGO_TITULOS + SEPARACION_OBRAS
End Function
'-----------------------------------------------------------------------------------------
'**************************************************************************
'**************************** INICIO PAGINA *******************************
'**************************************************************************

RPT_Division = GF_Parametros7("idDivision", 0, 6)
RPT_Month    = GF_Parametros7("month", "", 6)
RPT_Year     = GF_Parametros7("year", "", 6)
RPT_Filtro     = GF_Parametros7("filtro", 0, 6)
RPT_accion   = GF_Parametros7("accion", "", 6)

if (RPT_accion = ACCION_PROCESAR) then
	if ((RPT_Filtro = TIPO_CATEGORIA) or (RPT_Filtro = TIPO_TODOS)) then
		RPT_Generando = TIPO_CATEGORIA
	elseif (RPT_Filtro = TIPO_PART_PRES) then
		RPT_Generando = TIPO_PART_PRES
	end if
	nroPagina = 1
	Set oPDF = GF_createPDF("PDFTemp")
	Call GF_setPDFMODE(PDF_STREAM_MODE)
	Call armadoPDF(oPDF)
	Call GF_closePDF(oPDF)
else
	Response.Redirect "comprasAccesoDenegado.asp"
end if
%>