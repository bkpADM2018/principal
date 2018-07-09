<!-- #include file="Includes/procedimientosAlmacenes.asp"-->
<!-- #include file="Includes/procedimientosUser.asp"-->
<!-- #include file="Includes/procedimientosFormato.asp"-->
<!-- #include file="Includes/procedimientosExcel.asp"-->
<!-- #include file="Includes/procedimientosVales.asp"-->
<!-- #include file="Includes/procedimientosParametros.asp"-->
<!-- #include file="Includes/procedimientosTraducir.asp"-->
<!-- #include file="Includes/procedimientosFechas.asp"-->
<%
Const LINEA_POST_ENCABEZADO = 75
Const PAGE_HEIGHT_SIZE = 800
Const SEPARACION_Y = 10

Dim nroPagina,lineaActual, g_dsPeriodoActual, g_anio, g_dsDivision1, g_dsDivision2, g_idDivision1, g_idDivision2

nroPagina = 1

'-----------------------------------------------------------------------------------------------
Function dibujarFiltros(p_parametros)
'Funcion que dibuja los parametros de busqueda 
'recibe como parametro un vector con la siguiente estructura en cada posicion:
'	"nombreBusqueda,valorBusqueda"
'Devuelve la posicion donde se puede seguir escribiendo
	Dim aux
    
  	Dim max_len
	
	max_len = 0
	for i = 0 to ubound(p_parametros)
		aux = split(p_parametros(i),",")
		if (len(aux(0)) > max_len) then
			max_len = len(aux(0))
		end if
	next
	
	for i = 0 to ubound(p_parametros)
		aux = split(p_parametros(i),",")
        writeXLS("<tr>")
            writeXLS("<td align='left'>"& aux(0) & ": </td>")
            writeXLS("<td align='left' colspan='5'>"& aux(1) &"</td>")
        writeXLS("</tr>")
	next
	
End Function
'--------------------------------------------------------------------------------------------------
'JAS - OK!
Function dibujarEncabezado()
Dim rs

        writeXLS("<tr>")  
            writeXLS("<TD align='right' colspan='6' style='font-size:10px;' >"&GF_FN2DTE(session("MmtoSistema"))&"</TD>")
        writeXLS("</tr>")
        writeXLS("<tr>")  
            writeXLS("<TD align='right' colspan='6' style='font-size:10px;' >"&session("Usuario")&"</TD>")
        writeXLS("</tr>")
        writeXLS("<tr>")  
            writeXLS("<TH align='center' colspan='6' style='font-size:20px;font-family:arial;' >ARTICULOS CONSUMIDOS</TH>")
        writeXLS("</tr>")
end Function

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
        writeXLS("<tr>")  
            writeXLS("<TH align='center' rowspan='2' style='font-family:ARIAL;background-color:#517b4a;color:#FFFFFF;'>"&GF_TRADUCIR("ID Art.")& "</TH>")
            writeXLS("<TH align='center' rowspan='2' style='font-family:ARIAL;width:300px;background-color:#517b4a;color:#FFFFFF;' >"&GF_TRADUCIR("Descripcion")&"</TH>")
            writeXLS("<TH align='center' rowspan='2' style='font-family:ARIAL;background-color:#517b4a;color:#FFFFFF;' >"&GF_TRADUCIR(g_dsPeriodoActual)&" "&g_anio&"</TH>")
            writeXLS("<TH align='center' rowspan='2' style='font-family:ARIAL;background-color:#517b4a;color:#FFFFFF;'>"&GF_TRADUCIR(g_dsPeriodoActual)&" "&g_anio-1&"</TH>")
            writeXLS("<TH align='center' colspan='2' style='font-family:ARIAL;background-color:#517b4a;color:#FFFFFF;'>"&GF_TRADUCIR(g_dsPeriodoActual) & " " & g_anio&"</TH>")
        writeXLS("</tr>")
        writeXLS("<tr>")
            writeXLS("<TH align='center' style='font-family:ARIAL;background-color:#517b4a;color:#FFFFFF;' >"&GF_TRADUCIR(g_dsDivision1)&"</TH>")
            writeXLS("<TH align='center' style='font-family:ARIAL;background-color:#517b4a;color:#FFFFFF;' >"&GF_TRADUCIR(g_dsDivision2)&"</TH>")
        writeXLS("</tr>") 
	
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
							
		totalActual = totalActual + CDbl(actual)
		totalAnterior = totalAnterior + CDbl(anterior)											
		totalActual1 = totalActual1 + CDbl(actual1)
		totalActual2 = totalActual2 + CDbl(actual2)
		
		actual = GF_EDIT_DECIMALS(actual,2) 
		anterior = GF_EDIT_DECIMALS(anterior,2)
		actual1 = GF_EDIT_DECIMALS(actual1,2)
		actual2 = GF_EDIT_DECIMALS(actual2,2)
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
	Dim rsConsumo,auxDs,actual,anterior,i,abrev,totalActual,totalAnterior,totalAnual
	Dim fechaDesde, fechaHasta
	Dim actual1, anterior1, totalActual1, totalAnterior1, actual2, anterior2, totalActual2, totalAnterior2
	'Se obtienen las fechas del periodo en estudio.
	Call obtenerFechas(pPeriodo, g_anio, fechaDesde, fechaHasta)	
	set rsConsumo = getSQLConsumo(fechaDesde, fechaHasta, pTipoReporte, cdSolicitante, idDivision, g_idDivision1, g_idDivision2, idAlmacen, articulo, categoria, Detalle)
	if (rsConsumo.EoF) then 
		'No hay datos asociados a los filtros introducidos.
        writeXLS("<tr>")  
            writeXLS("<TD align='center' colspan='6' >"&GF_TRADUCIR("No se encontraron resultados.")&"</TD>")
        writeXLS("</tr>")
	else
		'Se procesan los resultados y se arma el reporte.
		i = 0		
		while not rsConsumo.EoF
			writeXLS("<tr>")  
		
			Call calcularValores(rsConsumo, 0, rsConsumo("abreviatura"), pTipoReporte, actual, anterior, totalActual, totalAnterior, actual1, totalActual1, actual2, totalActual2)
			if (len(auxDs) > 50) then auxDs = left(auxDs,50) & "..."
            writeXLS("<TD align='right' >"&rsConsumo("idArticulo")&"</TD>")
            writeXLS("<TD align='left' >"&rsConsumo("dsArticulo")&"</TD>")
            writeXLS("<TD align='right' >"&actual&"</TD>")
            writeXLS("<TD align='right' >"&anterior&"</TD>")
            writeXLS("<TD align='right' >"&actual1&"</TD>")
            writeXLS("<TD align='right' >"&actual2&"</TD>")
            writeXLS("</TR>")		
			if(Detalle = "SI")then
				call dibujarTitularBoxDetalle()		
				call dibujarDetalleArticulo(rsConsumo("cdusuario"), rsConsumo("CANT"),rsConsumo("abreviatura"))
			end if		
            i = i + 1
			rsConsumo.MoveNext	
		wend
        
		if (LCase(pTipoReporte) <> REPORTE_CANTIDAD) then			
            writeXLS("<TR>")
                writeXLS("<TD align='right' colspan='2' style='background-color:#517b4a;font-family:COURIER;color:#FFFFFF;'>TOTAL</TD>")
                writeXLS("<TD align='right' style='background-color:#517b4a;font-family:COURIER;color:#FFFFFF;'>" & GF_EDIT_DECIMALS(totalActual, 2)&"</TD>")
                writeXLS("<TD align='right' style='background-color:#517b4a;font-family:COURIER;color:#FFFFFF;'>" & GF_EDIT_DECIMALS(totalAnterior, 2)&"</TD>")
                writeXLS("<TD align='right' style='background-color:#517b4a;font-family:COURIER;color:#FFFFFF;'>" & GF_EDIT_DECIMALS(totalActual1, 2)&"</TD>")
                writeXLS("<TD align='right' style='background-color:#517b4a;font-family:COURIER;color:#FFFFFF;'>" & GF_EDIT_DECIMALS(totalActual2, 2)&"</TD>")
            WriteXLS("</TR>")
		end if
	end if
    
End Function
'--------------------------------------------------------------------------------------------------
Function dibujarDetalleArticulo(cdusuario, consumoActual, abreviatura)  
        writeXLS("<TD></TD>")
        writeXLS("<TD align='center'  style='background-color:#F5F6CE;font-family:COURIER;' >"&getUserDescription(cdusuario)&"</TD>")
        writeXLS("<TD align='center' colspan='2' style='background-color:#F5F6CE;font-family:COURIER;' >"&consumoActual&" "&abreviatura&"</TD>")
end function 
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
Function getSQLConsumo(fechaInicioActual, fechaFinActual, pTipoReporte, cdSolicitante, idDivision, idDivision1, idDivision2, idAlmacen, idarticulo, idcategoria, Detalle)
	
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
	if(Detalle = "SI")then strSQL = strSQL & "         actual.cdusuario, "
	strSQL = strSQL & "         actual.cantidad CANT "		
	strSQL = strSQL & "FROM		(" & getSQLConsumoPeriodo(fechaInicioActual,fechaFinActual,pTipoReporte, cdSolicitante, idDivision, idAlmacen,idarticulo, idcategoria, Detalle) & ") actual"
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
	if(verDetalle = "SI") then strSQL = strSQL & "         cab.cdusuario, "	
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
	if(verDetalle = "SI")then strSQL = strSQL & ",         cab.cdusuario "	
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
'Funcion responsable por la creacion del XLS.
Function crearXLS(pDivision, pAlmacen, pPeriodo, pSolicitante, pTipoReporte, pVerDetalle,pidArticulo,pCategoria)
	
	Dim vParam(7),usr, auxCatDS
	
	Call GF_createXLS("Reporte")
	writeXLS("<TABLE style='border:1px solid #ddd;'>")
    writeXLS("<THEAD>")
	Call dibujarEncabezado()
	
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
	writeXLS("</THEAD>")
    writeXLS("</TABLE>")
	Call closeXLS()	
	
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
Function emitirReporteXLS(pAlmacen, pDivision, pPeriodo, pPeriodoDS, pAnio, pSolicitante, pTipoReporte, pVerDetalle, pidArticulo, pCategoria)
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
	Call crearXLS(pDivision, pAlmacen, pPeriodo, pSolicitante, pTipoReporte, pVerDetalle, pidArticulo, pCategoria)
End Function
'--------------------------------------------------------------------------------------------------
Function dibujarTitularBoxDetalle()
        writeXLS("<TR>")
        writeXLS("<TD>")
        writeXLS("</TD>")
        writeXLS("<TD align='center' style='background-color:#F2F5A9;font-family:ARIAL;color:#000000;' >"&GF_TRADUCIR("USUARIO")&"</TD>")
        writeXLS("<TD align='center' style='background-color:#F2F5A9;font-family:ARIAL;color:#000000;' colspan='2' >"&GF_TRADUCIR("CANTIADAD")&"</TD>")
        writeXLS("</TR>")
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

Call emitirReporteXLS(pAlmacen, pDivision, pPeriodo, pPeriodoDS, pAnio, pSolicitante, pTipoReporte, pVerDetalle, pidArticulo, pCategoria)

%>