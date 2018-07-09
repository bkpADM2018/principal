<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosPDF.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<%
'Constantes de colores
Const VERDE  = "#396E8F"
Const VERDE2 = "#000000"
Const ROJO   = "#FFAA99"
Const NEGRO  = "#000000"
Const GRIS   = "#ADAFA7"
Const BLANCO = "#FFFFFF"

Const TRIMESTRE_1 = 0
Const TRIMESTRE_2 = 1
Const TRIMESTRE_3 = 2
Const TRIMESTRE_4 = 3
Const RESUMEN     = 4
Const NOTAS_REASIG = 5

Const MAX_LEN_DESC = 34
Const MAX_LEN_DESC_RESUMEN = 45

Const MAX_PAGE_LINES = 70

Const MES_1 = 1
Const MES_2 = 2
Const MES_3 = 3

DIM V_MESES()

redim V_MESES(4, 3, 2)
V_MESES(0,0,0) = "Enero"
V_MESES(0,0,1) = 1
V_MESES(0,1,0) = "Febrero"
V_MESES(0,1,1) = 2
V_MESES(0,2,0) = "Marzo"
V_MESES(0,2,1) = 3
V_MESES(1,0,0) = "Abril"
V_MESES(1,0,1) = 4
V_MESES(1,1,0) = "Mayo"
V_MESES(1,1,1) = 5
V_MESES(1,2,0) = "Junio"
V_MESES(1,2,1) = 6
V_MESES(2,0,0) = "Julio"
V_MESES(2,0,1) = 7
V_MESES(2,1,0) = "Agosto"
V_MESES(2,1,1) = 8
V_MESES(2,2,0) = "Septiembre"
V_MESES(2,2,1) = 9
V_MESES(3,0,0) = "Octubre"
V_MESES(3,0,1) = 10
V_MESES(3,1,0) = "Noviembre"
V_MESES(3,1,1) = 11
V_MESES(3,2,0) = "Diciembre"
V_MESES(3,2,1) = 12

Dim V_TRIM(4)
V_TRIM(0) = "Trimestre 1"
V_TRIM(1) = "Trimestre 2"
V_TRIM(2) = "Trimestre 3"
V_TRIM(3) = "Trimestre 4"

DIM V_INICIO_TRIM(4)
V_INICIO_TRIM(0) = "0101"
V_INICIO_TRIM(1) = "0401"
V_INICIO_TRIM(2) = "0701"
V_INICIO_TRIM(3) = "1001"




'--------------------------------------------------------------------------------------------------
' Autor: 	GFG - Guido Fonticelli
' Fecha: 	01/01/11
' Objetivo:	
'			Genera las SQL de los trimestres/obras
' Devuelve:
'			[str] - La SQL
' Modificaciones:
'			06/01/11 - GFG
'--------------------------------------------------------------------------------------------------
Function generarSQLGastos(pTrim)
	Dim strSQL,campoMoneda1,campoMoneda2,campoMoneda3, filtroTrimestre
	
	'Se prepara el filtro de trimestre.
	filtroTrimestre = ""
	if (pTrim <>RESUMEN) then filtroTrimestre = "where Trimestre = " & pTrim

	'Esta SQL devolvera un query que contendra la siguiente estructura
	' TRIMESTRE ( de 0 a 3 ) | IDAREA | IDDETALLE | MES (mes al que pertenece el importe) | GASTO
	'
	' Para obtener el trimestre se realizo el calculo siguiente:
	' parte entera ( (mesInicioObra-1) / 3 )	
		
	strSQL = ""
	if ((gChkVales) or (gChkPIC) or (gChkFacturacion))then	
		campoImporte = "importedolares"	
		campoImporte2= "det.vluDOLARES"
		campoImporte31= "((det.IMPORTEPESOS-det.IMPORTEPESOSFACTURADO)/" & gTipoCambio & ")"
		campoImporte32= "(det.IMPORTEDOLARES-det.IMPORTEDOLARESFACTURADO)"
		campoImporte41 = "(IMPORTEASIGNADO/" & gTipoCambio & ")"
		campoImporte42= "IMPORTEASIGNADO"
		if (gMoneda=MONEDA_PESO) then 
			campoImporte = "importepesos"
			campoImporte2= "det.vlupesos"
			campoImporte31= "(det.IMPORTEPESOS-det.IMPORTEPESOSFACTURADO)"
			campoImporte32= "((det.IMPORTEDOLARES-det.IMPORTEDOLARESFACTURADO)*" & gTipoCambio & ")"
			campoImporte41= "IMPORTEASIGNADO"
		    campoImporte42= "(IMPORTEASIGNADO*" & gTipoCambio & ")"
		end if
	
		myFechaHasta = "1" & right(fechaHasta,6)
		myFechaAMD = 20 & right(myFechaHasta,len(myFechaHasta)-1)
	
		strSQL = ""
		strSQL = strSQL & "SELECT   trimestre   , "
		strSQL = strSQL & "         area   , "
		strSQL = strSQL & "         detalle, "
		if (pTrim = RESUMEN) then
			strSQL = strSQL & "         mes      , "
		else
			strSQL = strSQL & "         (mes-(3*" & pTrim & ")) mes      , "
		end if
		strSQL = strSQL & "			case when sum(gasto) is null then 0 "
		strSQL = strSQL & "			else sum(gasto) end as gasto "		
		strSQL = strSQL & "FROM     ( "
		' sql para las facturas que corresponden a la inversion.		
		strSQL = strSQL & "			SELECT  1 as id, CAST((CAST(acd7.mes AS INTEGER)-1)/3 AS INTEGER) trimestre, "
		strSQL = strSQL & "                  acd7.IDAREA         Area   , "
		strSQL = strSQL & "                  acd7.IDDETALLE         detalle, "
		strSQL = strSQL & "                  acd7.mes , "			
		strSQL = strSQL & "                  SUM( "&campoImporte&")*100 gasto "			
		strSQL = strSQL & "         FROM     VWMEP001C acd7 "
		strSQL = strSQL & "                  INNER JOIN VWCOMPROBANTES acds ON acd7.NROINT = acds.NROINT"
		strSQL = strSQL & "					 INNER JOIN tblarticulos art on art.idarticulo = acd7.IDARTICULO "
		strSQL = strSQL & "					 INNER JOIN tblartcategorias cat on art.idcategoria=cat.idcategoria "	
		strSQL = strSQL & "         WHERE    acd7.IDOBRA          = " & gIdObra
		strSQL = strSQL & "			AND		 acd7.IDARTICULO NOT IN (" & ITEM_FONDO_REPARO_ARS & "," & ITEM_FONDO_REPARO_USD & ", " & ITEM_FONDO_REPARO_ARS_IVA & "," & ITEM_FONDO_REPARO_USD_IVA & ")" 
		strSQL = strSQL & "			and		 cat.tipocategoria	  <> '" & TIPO_CAT_IMPUESTOS & "'"
		strSQL = strSQL & "         AND      convert(varchar(10), acds.feccbt, 112) <= '" & myFechaAMD & "'"
	    strSQL = strSQL & "         GROUP BY CAST((CAST(acd7.mes AS INTEGER)-1)/3 AS INTEGER), "
		strSQL = strSQL & "                  acd7.mes, "
		strSQL = strSQL & "                  acd7.IDAREA       , "
		strSQL = strSQL & "                  acd7.IDDETALLE "
		strSQL = strSQL & "          "		
		' sql para obtener los vales que correspondan a la inversion.
		if (gChkVales) then			
		    strSQL = strSQL & " UNION "					
			strSQL = strSQL & "         SELECT   2 as id, CAST((CAST(SUBSTRING(CAST(cab.momento AS VARCHAR(16)),5,2) AS INTEGER)-1)/3 AS INTEGER) trimestre   , "
			strSQL = strSQL & "                  cab.idbudgetarea                 area   , "
			strSQL = strSQL & "                  cab.idbudgetdetalle              detalle, "
			strSQL = strSQL & "                  SUBSTRING(CAST(cab.fecha AS VARCHAR(16)),5,2)            mes    , "
			strSQL = strSQL & "                  SUM(det.existencia*"&campoImporte2&") gasto "
			strSQL = strSQL & "         FROM     tblvalescabecera cab "
			strSQL = strSQL & "                  INNER JOIN tblvalesdetalle det "
			strSQL = strSQL & "                  ON       cab.idvale = det.idvale "
			strSQL = strSQL & "         WHERE    cab.idobra          = " & gIdObra
			strSQL = strSQL & "         AND      cab.estado = "& ESTADO_ACTIVO
			strSQL = strSQL & "         AND      cab.fecha <= " & myFechaAMD
			strSQL = strSQL & "         GROUP BY CAST((CAST(SUBSTRING(CAST(cab.momento AS VARCHAR(16)),5,2) AS INTEGER)-1)/3 AS INTEGER), "
			strSQL = strSQL & "                  SUBSTRING(CAST(cab.fecha AS VARCHAR(16)),5,2), "
			strSQL = strSQL & "                  cab.idbudgetarea     , "
			strSQL = strSQL & "                  cab.idbudgetdetalle "	
		end if
		'sub sql para los pic que correspondan a la inversion.
		if (gChkPIC) then
			strSQL = strSQL & " UNION "
			strSQL = strSQL & "         SELECT   31 as id, CAST((CAST(SUBSTRING(CAST(cab.momento AS VARCHAR),5,2) AS INTEGER)-1)/3 AS INTEGER) trimestre	, "
			strSQL = strSQL & "                  det.idarea                 area	, "
			strSQL = strSQL & "                  det.iddetalle              detalle , "
			strSQL = strSQL & "                  SUBSTRING(CAST(cab.momento AS VARCHAR(16)),5,2)	mes     , "
			strSQL = strSQL & "                  SUM("&campoImporte31&")		gasto "
			strSQL = strSQL & "         FROM     tblctzcabecera cab "
			strSQL = strSQL & "                  INNER JOIN tblctzdetalle det "
			strSQL = strSQL & "                  ON       cab.idcotizacion = det.idcotizacion "
			strSQL = strSQL & "         WHERE    cab.idobra                = " & gIdObra
			strSQL = strSQL & "					 AND cab.idcontrato=0 "
			strSQL = strSQL & "					 AND cab.estado <> '" & CTZ_ANULADA & "' "
			strSQL = strSQL & "                  AND cab.momento <= " & myFechaAMD & "595959"									
			strSQL = strSQL & "                  AND  cab.CDMONEDA = '" & MONEDA_PESO & "'"
			strSQL = strSQL & "         GROUP BY CAST((CAST(SUBSTRING(CAST(cab.momento AS VARCHAR),5,2) AS INTEGER)-1)/3 AS INTEGER), "
			strSQL = strSQL & "                  SUBSTRING(CAST(cab.momento AS VARCHAR(16)),5,2), "	
			strSQL = strSQL & "					 det.idarea             , "
			strSQL = strSQL & "                  det.iddetalle            "
			strSQL = strSQL & " UNION "
			strSQL = strSQL & "         SELECT   32 as id, CAST((CAST(SUBSTRING(CAST(cab.momento AS VARCHAR(16)),5,2) AS INTEGER)-1)/3 AS INTEGER) trimestre	, "
			strSQL = strSQL & "                  det.idarea                 area	, "
			strSQL = strSQL & "                  det.iddetalle              detalle , "
			strSQL = strSQL & "                  SUBSTRING(CAST(cab.momento AS VARCHAR(16)),5,2)	mes     , "
			strSQL = strSQL & "                  SUM("&campoImporte32&")	gasto "
			strSQL = strSQL & "         FROM     tblctzcabecera cab "
			strSQL = strSQL & "                  INNER JOIN tblctzdetalle det "
			strSQL = strSQL & "                  ON       cab.idcotizacion = det.idcotizacion "
			strSQL = strSQL & "         WHERE    cab.idobra                = " & gIdObra
			strSQL = strSQL & "					 AND cab.idcontrato=0 "
			strSQL = strSQL & "					 AND cab.estado <> '" & CTZ_ANULADA & "' "
			strSQL = strSQL & "                  AND cab.momento <= " & myFechaAMD & "595959"							
			strSQL = strSQL & "                  AND  cab.CDMONEDA = '" & MONEDA_DOLAR & "'"
			strSQL = strSQL & "         GROUP BY CAST((CAST(SUBSTRING(CAST(cab.momento AS VARCHAR(16)),5,2) AS INTEGER)-1)/3 AS INTEGER), "
			strSQL = strSQL & "                  SUBSTRING(CAST(cab.momento AS VARCHAR(16)),5,2), "	
			strSQL = strSQL & "					 det.idarea             , "
			strSQL = strSQL & "                  det.iddetalle            "				
			strSQL = strSQL & " UNION "
			'Se adicionan a lo comprometido los contratos con AREA-DETALLE.
			strSQL = strSQL & "         SELECT 41 as id, TRIMESTRE, AREA, DETALLE, MES, Sum(SALDOCTC) GASTO    "
	        strSQL = strSQL & "         FROM    "
	        strSQL = strSQL & "             (   SELECT			"
            strSQL = strSQL & "                  CASE WHEN P.FECHAINICIO >= " & gObraFechaInicio & " THEN CAST((CAST(SUBSTRING(CAST(P.FECHAINICIO AS VARCHAR(16)),5,2) AS INTEGER)-1)/3 AS INTEGER) ELSE CAST((CAST(SUBSTRING(CAST(" & gObraFechaInicio & " AS VARCHAR(16)),5,2) AS INTEGER)-1)/3 AS INTEGER) END	TRIMESTRE	, "
			strSQL = strSQL & "                  P.IDAREA                 AREA	, "
			strSQL = strSQL & "                  P.IDDETALLE              DETALLE , "
			strSQL = strSQL & "                  CASE WHEN SUBSTRING(CAST(P.FECHAINICIO AS VARCHAR(16)),1,8) >= " & gObraFechaInicio & " THEN SUBSTRING(CAST(P.FECHAINICIO AS VARCHAR(16)),5,2) ELSE " & Mid(gObraFechaInicio, 5, 2)  & " END 	MES     , "			
			strSQL = strSQL & "                  " & campoImporte41 & "	SALDOCTC "						
			strSQL = strSQL & "                 FROM TBLOBRACONTRATOS CTC "
			strSQL = strSQL & "                 INNER JOIN TBLCTCPARTIDAS P on P.IDCONTRATO=CTC.IDCONTRATO "
			strSQL = strSQL & "                 WHERE P.CDMONEDA='" & MONEDA_PESO & "' and CTC.ESTADO in (" & ESTADO_CTC_AUTORIZADO & ", " & ESTADO_CTC_FINALIZADO & ", " & ESTADO_CTC_EN_AJUSTE & ") AND P.IDOBRA=" & gIdObra
			strSQL = strSQL & "                 AND P.IDAREA<>0 and P.IDDETALLE<>0 "
			strSQL = strSQL & "             ) T"
			strSQL = strSQL & "         GROUP BY T.TRIMESTRE, "
			strSQL = strSQL & "                  T.MES, "	
			strSQL = strSQL & "					 T.AREA             , "
			strSQL = strSQL & "                  T.DETALLE            "
			strSQL = strSQL & " UNION "
			'Se adicionan a lo comprometido los contratos con AREA-DETALLE.
			strSQL = strSQL & "         SELECT 42 as id,TRIMESTRE, AREA, DETALLE, MES, Sum(SALDOCTC) GASTO    "
	        strSQL = strSQL & "         FROM    "
	        strSQL = strSQL & "             (   SELECT			"
            strSQL = strSQL & "                  CASE WHEN P.FECHAINICIO >= " & gObraFechaInicio & " THEN CAST((CAST(SUBSTRING(CAST(P.FECHAINICIO AS VARCHAR(16)),5,2) AS INTEGER)-1)/3 AS INTEGER) ELSE CAST((CAST(SUBSTRING(CAST(" & gObraFechaInicio & " AS VARCHAR(16)),5,2) AS INTEGER)-1)/3 AS INTEGER) END	TRIMESTRE	, "
			strSQL = strSQL & "                  P.IDAREA                 AREA	, "
			strSQL = strSQL & "                  P.IDDETALLE              DETALLE , "
			strSQL = strSQL & "                  CASE WHEN SUBSTRING(CAST(P.FECHAINICIO AS VARCHAR(16)),1,8) >= " & gObraFechaInicio & " THEN SUBSTRING(CAST(P.FECHAINICIO AS VARCHAR(16)),5,2) ELSE " & Mid(gObraFechaInicio, 5, 2)  & " END 	MES     , "			
			strSQL = strSQL & "                  " & campoImporte42 & "	SALDOCTC "						
			strSQL = strSQL & "                 FROM TBLOBRACONTRATOS CTC "
			strSQL = strSQL & "                 INNER JOIN TBLCTCPARTIDAS P on P.IDCONTRATO=CTC.IDCONTRATO "
			strSQL = strSQL & "                 WHERE P.CDMONEDA='" & MONEDA_DOLAR & "' and CTC.ESTADO in (" & ESTADO_CTC_AUTORIZADO & ", " & ESTADO_CTC_FINALIZADO & ", " & ESTADO_CTC_EN_AJUSTE & ") AND P.IDOBRA=" & gIdObra
			strSQL = strSQL & "                 AND P.IDAREA<>0 and P.IDDETALLE<>0 "
			strSQL = strSQL & "             ) T"
			strSQL = strSQL & "         GROUP BY T.TRIMESTRE, "
			strSQL = strSQL & "                  T.MES, "	
			strSQL = strSQL & "					 T.AREA             , "
			strSQL = strSQL & "                  T.DETALLE            "			
		end if
		 
		strSQL = strSQL & "         ) AS "
		strSQL = strSQL & "         aux "
		strSQL = strSQL & "GROUP BY TRIMESTRE   , "
		strSQL = strSQL & "         area   , "
		strSQL = strSQL & "         detalle, "
		strSQL = strSQL & "         mes "
		strSQL = strSQL & "ORDER BY area   , "
		strSQL = strSQL & "         detalle   , "
		strSQL = strSQL & "         TRIMESTRE, "
		strSQL = strSQL & "         mes"	
	else
		'No se selecciono ningun tipo de consumo, se crea una SQL solo para evitar 
		strSQL = "Select * from TBLBUDGETOBRAS where 1=2"
	end if
	'Response.Write strSQL
	'Response.End 
	generarSQLGastos = strSQL
End Function
'--------------------------------------------------------------------------------------------------
Function cargaDatos()

	resumenGeneral	= GF_Parametros7("resumen"   , "", 6)
	trimestre1  	= GF_Parametros7("trimestre1", "", 6)
	trimestre2  	= GF_Parametros7("trimestre2", "", 6)
	trimestre3  	= GF_Parametros7("trimestre3", "", 6)
	trimestre4  	= GF_Parametros7("trimestre4", "", 6)

	Call loadDatosObra(gIdObra, obraCD, obraDS, obraDivID, obraDivDS, obraImorte, obraFechaBudget, obraMonedaID, gObraFechaInicio, gObraFechaFin, obraFechaAjustada, obraRespCD, obraRespDS)
	
	gMostrarResumen = false
	if (resumenGeneral = "true" or resumenGeneral = "") then gMostrarResumen = true

	gMostrar1erTrim = false
	if (trimestre1 = "true") then gMostrar1erTrim = true

	gMostrar2doTrim = false
	if (trimestre2 = "true") then gMostrar2doTrim = true

	gMostrar3erTrim = false
	if (trimestre3 = "true") then gMostrar3erTrim = true

	gMostrar4toTrim = false
	if (trimestre4 = "true") then gMostrar4toTrim = true

	if (gMoneda = "") then gMoneda = MONEDA_DOLAR
	'Se pre-determina la fecha de hoy como el dia actual.	
	gHoy = CLng(left(session("MmtoDato"),8))
	'Se pre-determina la fecha del presupuesto como el final del ejercicio.
	gFechaBudget  =Year(Date()) & "1231"

	'Se analiza si se piedieron los datos a una fecha especifica.
	if (gFechaHasta <> "") then	gHoy = CLng(GF_DTE2FN(gFechaHasta))		
	if (gBgtParcial) then gFechaBudget = gHoy
	
	gNroHojas = 1
	gNroLinea = 0
	gCantComentarios = 0
	
	redim gVecComentarios(0)
	
	strSQL = "select * from tblbudgetobras where idobra = " & gIdObra & " order by idarea,iddetalle"
	Call executeQueryDb(DBSITE_SQL_INTRA, gRsObra, "OPEN", strSQL)
End Function
'--------------------------------------------------------------------------------------------------
Function agregarHoja(pTrim)
	gNroLinea = 0
	gNroHojas = gNroHojas + 1
	Call GF_newPage(oPDF)
End Function
'--------------------------------------------------------------------------------------------------
Function dibujarEncabezado(pTrim)
	Dim myTrim,mySubTitulo
	'recuadro
	Call GF_squareBox(oPDF,3,5,590 ,830,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_ROUND) 
	
	'logo
	Call GF_writeImage(oPDF, Server.MapPath("Images\ADMlogo2.jpg"),15, 15, 48, 48, 0)
	
	'Titulo
	Call GF_setFont(oPDF,"ARIAL", 16 , FONT_STYLE_BOLD)
	if (pTrim = RESUMEN) or (pTrim = NOTAS_REASIG) then
		Call GF_writeTextAlign(oPDF,0, 15, GF_TRADUCIR("Reporte de Mantenimiento") , 590,PDF_ALIGN_CENTER)
		Call GF_setFont(oPDF,"ARIAL", 14 , FONT_STYLE_BOLD)
	else
		select case pTrim
			case TRIMESTRE_1
				myTrim = "1er"
			case TRIMESTRE_2
				myTrim = "2do"
			case TRIMESTRE_3
				myTrim = "3er"
			case TRIMESTRE_4
				myTrim = "4to"
				
		end select
		Call GF_writeTextAlign(oPDF,0, 15, GF_TRADUCIR("Resumen "&myTrim&" Trimestre") , 590,PDF_ALIGN_CENTER)
		Call GF_setFont(oPDF,"ARIAL", 14 , FONT_STYLE_BOLD)
	end if
	'Titulo de la partida.
	Call GF_writeTextAlign(oPDF,0, 45, obraCD & "-" & obraDS , 590,PDF_ALIGN_CENTER)
	
	Call GF_horizontalLine(oPDF,5,70,585)
	
	Call GF_setFont(oPDF,"COURIER",8,FONT_STYLE_NORMAL)
	Call GF_writeTextAlign(oPDF,0, 8, GF_FN2DTE(session("MmtoSistema")) , 590,PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,0,15,session("Usuario"), 587 , PDF_ALIGN_RIGHT)
	

	Call GF_writeTextAlign(oPDF,10, 80, GF_TRADUCIR("Consumo a la fecha: ") & GF_FN2DTE(gHoy) & ", " & GF_TRADUCIR("Budget a la fecha: ") & GF_FN2DTE(gFechaBudget) , 590,PDF_ALIGN_LEFT)

	aux="NO"
	if (gChkPIC) then aux="SI"
	stringFiltro = GF_TRADUCIR("Incluir Comprometido") & ": " & aux	
	aux="NO"
	if (gChkFacturacion) then aux="SI"
	stringFiltro = stringFiltro & ", " & GF_TRADUCIR("Incluir Facturas") & ": " & aux
	aux="NO"
	if (gChkVales) then aux="SI"
	stringFiltro = stringFiltro & ", " & GF_TRADUCIR("Incluir Vales") & ": " & aux
	Call GF_writeTextAlign(oPDF,10, 88, stringFiltro, 590,PDF_ALIGN_LEFT)
	
	Call GF_writeTextAlign(oPDF,5, 88, GF_TRADUCIR("T. Cambio: ") & gTipoCambio 	, 580,PDF_ALIGN_RIGHT)
	
	'Numero de pagina
	Call GF_writeTextAlign(oPDF,0, 835, GF_TRADUCIR("Pagina") & " " & gNroHojas , 590,PDF_ALIGN_RIGHT)
		
End Function 
'--------------------------------------------------------------------------------------------------
Function dibujarLineaTrimestre(esTitulo, nroLineaLogica, idItem, desc, impMes1, impMes2, impMes3, impBudget)
	Dim myDesvPerc, myDesv, nroLinea,totalMeses
	
	nroLinea = 98+(nroLineaLogica*10)
	totalMeses = impMes1+impMes2+impMes3
	'Descripcion
	if (len(desc)>MAX_LEN_DESC) then desc = left(desc,MAX_LEN_DESC) & "..."
	
	if (esTitulo) then
		colorFondo = GRIS
		'Estructura propia de la linea de titulo del area
		Call GF_squareBox(oPDF, 10   ,nroLinea,20 ,10 ,0 ,VERDE,NEGRO ,1 ,0)
		Call GF_squareBox(oPDF, 30   ,nroLinea,175 ,10 ,0 ,colorFondo,NEGRO ,1 ,0)	
		'Se dibujan los datos.
		Call GF_setFont(oPDF,"COURIER", 6,FONT_STYLE_BOLD)			
		Call GF_setFontColor(BLANCO)
		Call GF_writeTextAlign(oPDF,10, nroLinea+2, idItem, 20,PDF_ALIGN_CENTER)				
		Call GF_setFontColor(NEGRO)
		Call GF_writeTextAlign(oPDF,35, nroLinea+2, desc, 170,PDF_ALIGN_LEFT)		
		nroLineaTitulo = nroLineaLogica 'Se salva la linea del titulo actual para completar sus importes luego de procesar todos sus detalles.
	else
		colorFondo = BLANCO
		if ( (impBudget-(totalMeses)) < 0) then colorFondo = ROJO
		'Estructura propia de un item de detalle.
		Call GF_squareBox(oPDF, 30   ,nroLinea,20 ,10 ,0 ,VERDE,NEGRO ,1 ,0)
		Call GF_squareBox(oPDF, 50   ,nroLinea,155 ,10 ,0 ,colorFondo,NEGRO ,1 ,0)
		'Se dibujan los datos.			
		Call GF_setFontColor(BLANCO)
		Call GF_setFont(oPDF,"COURIER", 6,FONT_STYLE_BOLD)
		Call GF_writeTextAlign(oPDF,30, nroLinea+2, idItem, 20,PDF_ALIGN_CENTER)
		Call GF_setFontColor(NEGRO)
		Call GF_setFont(oPDF,"COURIER", 6,FONT_STYLE_NORMAL)		
		Call GF_writeTextAlign(oPDF,55, nroLinea+2, desc, 150,PDF_ALIGN_LEFT)			
	end if 
	'Estructura común.
	Call GF_squareBox(oPDF, 205   ,nroLinea,60 ,10 ,0 ,colorFondo,NEGRO ,1 ,0) '1er mes
	Call GF_squareBox(oPDF, 265   ,nroLinea,60 ,10 ,0 ,colorFondo,NEGRO ,1 ,0) '2do mes
	Call GF_squareBox(oPDF, 325   ,nroLinea,60 ,10 ,0 ,colorFondo,NEGRO ,1 ,0) '3er mes
	Call GF_squareBox(oPDF, 385   ,nroLinea,60 ,10 ,0 ,colorFondo,NEGRO ,1 ,0) 'total
	Call GF_squareBox(oPDF, 445   ,nroLinea,70 ,10 ,0 ,colorFondo,NEGRO ,1 ,0) 'budget
	Call GF_squareBox(oPDF, 515   ,nroLinea,40 ,10 ,0 ,colorFondo,NEGRO ,1 ,0) 'desv
	Call GF_squareBox(oPDF, 555   ,nroLinea,30 ,10 ,0 ,colorFondo,NEGRO ,1 ,0) 'desv %
	'Se completan los datos de la linea. (Datos comunes: Importes y Desvios)		
	Call GF_writeTextAlign(oPDF,205, nroLinea+2, GF_EDIT_DECIMALS(impMes1,2) , 55,PDF_ALIGN_RIGHT) '1er mes
	Call GF_writeTextAlign(oPDF,265, nroLinea+2, GF_EDIT_DECIMALS(impMes2,2) , 55,PDF_ALIGN_RIGHT) '2do mes
	Call GF_writeTextAlign(oPDF,325, nroLinea+2, GF_EDIT_DECIMALS(impMes3,2) , 55,PDF_ALIGN_RIGHT) '3er mes
	
	'El total actual para el detalle.
	Call GF_writeTextAlign(oPDF,385, nroLinea+2, GF_EDIT_DECIMALS(totalMeses,2) , 55,PDF_ALIGN_RIGHT) 'Total	
			
	'El preuspuesto asignado.			
	Call GF_writeTextAlign(oPDF,445, nroLinea+2, GF_EDIT_DECIMALS(impBudget,2) , 65,PDF_ALIGN_RIGHT) 'budget
	
	'La desviacion.
	myDesv = impBudget - (impMes1+impMes2+impMes3)
	if (myDesv > 0) then myDesv = 0
	myDesvPerc  = 0
	if (impBudget > 0) then myDesvPerc = (myDesv/impBudget)*100	
	Call GF_writeTextAlign(oPDF,515, nroLinea+2, GF_EDIT_DECIMALS(abs(myDesv),2) , 35,PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,555, nroLinea+2, GF_EDIT_DECIMALS(abs(myDesvPerc),0) & " %" , 25,PDF_ALIGN_RIGHT)
	
	dibujarLineaTrimestre = nroLineaLogica+1
End Function
'--------------------------------------------------------------------------------------------------
Function dibujarLineaResumen(esTitulo, nroLineaLogica, idItem, desc, cuenta, cc, impTrim1, impTrim2, impTrim3, impTrim4, impBudget, comentarios)
	Dim myDesvPerc, myDesv, nroLinea,totalTrimestres,myDescripcion
	
	nroLinea = 98+(nroLineaLogica*10)
			
	totalTrimestres = impTrim1+impTrim2+impTrim3+impTrim4
	'Descripcion
	if (len(desc)>MAX_LEN_DESC_RESUMEN) then desc = left(desc,MAX_LEN_DESC_RESUMEN) & "..."
		
	if (esTitulo) then 
			colorFondo = GRIS
			'Estructura propia de la linea de titulo del area
			Call GF_squareBox(oPDF, 5, nroLinea, 20 ,10 ,0 ,VERDE,NEGRO ,1 ,0) 'id area			
			Call GF_squareBox(oPDF, 25,nroLinea,272 ,10 ,0 ,colorFondo,NEGRO ,1 ,0) 'Descripcion						
			'Se dibujan los datos.
			Call GF_setFontColor(BLANCO)
			Call GF_setFont(oPDF,"COURIER", 6,FONT_STYLE_BOLD)			
			Call GF_writeTextAlign(oPDF,5, nroLinea+2, idItem, 20,PDF_ALIGN_CENTER)			
			Call GF_setFontColor(NEGRO)
			Call GF_writeTextAlign(oPDF,30, nroLinea+2,desc, 270,PDF_ALIGN_LEFT)
			nroLineaTitulo = nroLineaLogica 'Se salva la linea del titulo actual para completar sus importes luego de procesar todos sus detalles.
		else			
			colorFondo = BLANCO
			if (impBudget-(totalTrimestres) < 0) then colorFondo = ROJO
			'Estructura propia de un item de detalle
			Call GF_squareBox(oPDF, 25, nroLinea,20 ,10 ,0 ,VERDE,NEGRO ,1 ,0) 'id detalle
			if(gChkContable)then
				Call GF_squareBox(oPDF, 45, nroLinea ,190,10 ,0 ,colorFondo, NEGRO ,1 ,0) 'descripcion
				Call GF_squareBox(oPDF, 165, nroLinea,40 ,10 ,0 ,colorFondo,NEGRO ,1 ,0) 'Cuenta
				Call GF_squareBox(oPDF, 205, nroLinea,22 ,10 ,0 ,colorFondo,NEGRO ,1 ,0) 'CCosto
			else 
				Call GF_squareBox(oPDF, 45, nroLinea ,252,10 ,0 ,colorFondo, NEGRO ,1 ,0) 'descripcion
			end if	
			'Se dibujan los datos.			
			Call GF_setFontColor(BLANCO)
			Call GF_setFont(oPDF,"COURIER", 6,FONT_STYLE_BOLD)
			Call GF_writeTextAlign(oPDF,25, nroLinea+2, idItem, 20,PDF_ALIGN_CENTER)			
			Call GF_setFont(oPDF,"COURIER", 6,FONT_STYLE_NORMAL)							
			Call GF_setFontColor(NEGRO)
			if(gChkContable)then
				myDescripcion = desc
				if(Len(myDescripcion) > 27 )then myDescripcion = left(myDescripcion,27) & "..."
				Call GF_writeTextAlign(oPDF,50, nroLinea+2, myDescripcion, 185,PDF_ALIGN_LEFT)
				Call GF_writeTextAlign(oPDF,165, nroLinea+2, cuenta, 40,PDF_ALIGN_CENTER)
				Call GF_writeTextAlign(oPDF,205, nroLinea+2, trim(cc), 22,PDF_ALIGN_CENTER)
			else				
				myDescripcion = desc
				if(Len(myDescripcion) > 52 )then myDescripcion = left(myDescripcion,52) & "..."
				Call GF_writeTextAlign(oPDF,50, nroLinea+2, myDescripcion , 247,PDF_ALIGN_LEFT)
			end if			
			'Si se genero un detalle sobre el item, se pone la maraca de referencia para indicar que hay una nota al pie del presupuesto.
			if (trim(comentarios)<> "") then
				gCantComentarios = gCantComentarios + 1
				redim preserve gVecComentarios(ubound(gVecComentarios)+1)
				gVecComentarios(ubound(gVecComentarios)) = comentarios		
				Call GF_setFont(oPDF,"COURIER", 4,FONT_STYLE_NORMAL)
				Call GF_writeTextAlign(oPDF,55, nroLinea+2, "("&gCantComentarios&")", 110,PDF_ALIGN_RIGHT)
				Call GF_setFont(oPDF,"COURIER", 6,FONT_STYLE_NORMAL)
			end if
		end if
		'Estructura común.
		Call GF_squareBox(oPDF, 227,nroLinea,50 ,10 ,0 ,colorFondo,NEGRO ,1 ,0) '1er Trim
		Call GF_squareBox(oPDF, 277,nroLinea,50 ,10 ,0 ,colorFondo,NEGRO ,1 ,0) '2do Trim
		Call GF_squareBox(oPDF, 327,nroLinea,50 ,10 ,0 ,colorFondo,NEGRO ,1 ,0) '3er Trim
		Call GF_squareBox(oPDF, 377,nroLinea,50 ,10 ,0 ,colorFondo,NEGRO ,1 ,0) '4to Trim
		Call GF_squareBox(oPDF, 427,nroLinea,53 ,10 ,0 ,colorFondo,NEGRO ,1 ,0) 'Total actual
		Call GF_squareBox(oPDF, 480,nroLinea,50 ,10 ,0 ,colorFondo,NEGRO ,1 ,0) 'Budget
		Call GF_squareBox(oPDF, 530,nroLinea,30 ,10 ,0 ,colorFondo,NEGRO ,1 ,0) 'Desv
		Call GF_squareBox(oPDF, 560,nroLinea,30 ,10 ,0 ,colorFondo,NEGRO ,1 ,0) 'desv %
		'Se completan los datos de la linea. (Datos comunes: Importes y Desvios)
		Call GF_writeTextAlign(oPDF,227, nroLinea+2, GF_EDIT_DECIMALS(impTrim1,2), 45,PDF_ALIGN_RIGHT)
		Call GF_writeTextAlign(oPDF,277, nroLinea+2, GF_EDIT_DECIMALS(impTrim2,2), 45,PDF_ALIGN_RIGHT)
		Call GF_writeTextAlign(oPDF,327, nroLinea+2, GF_EDIT_DECIMALS(impTrim3,2), 45,PDF_ALIGN_RIGHT)
		Call GF_writeTextAlign(oPDF,377, nroLinea+2, GF_EDIT_DECIMALS(impTrim4,2), 45,PDF_ALIGN_RIGHT)
			
		'El total actual para el detalle.
		Call GF_writeTextAlign(oPDF,427, nroLinea+2, GF_EDIT_DECIMALS(totalTrimestres,2), 50,PDF_ALIGN_RIGHT)
			
		'El preuspuesto asignado.		
		Call GF_writeTextAlign(oPDF,480, nroLinea+2, GF_EDIT_DECIMALS(impBudget,2), 45,PDF_ALIGN_RIGHT)
			
		'La desviacion.
		myDesv = impBudget - (impTrim1+impTrim2+impTrim3+impTrim4)
		if (myDesv > 0) then myDesv = 0
		myDesvPerc  = 0
		if (impBudget > 0) then myDesvPerc = (myDesv/impBudget)*100			
		Call GF_writeTextAlign(oPDF,530, nroLinea+2, GF_EDIT_DECIMALS(abs(myDesv),2) , 25,PDF_ALIGN_RIGHT)
		Call GF_writeTextAlign(oPDF,560, nroLinea+2, GF_EDIT_DECIMALS(abs(myDesvPerc),0) & " %" , 25,PDF_ALIGN_RIGHT)				
				
		dibujarLineaResumen = nroLineaLogica+1		
		
End Function
'--------------------------------------------------------------------------------------------------
'Funcion responsable de evaluar si hay datos validos y cumplen la condición de corte de los datos de gastos del resumen.
Function corteControlGastos(rsAux, pAreaActual, pDetalleActual)
	Dim ret 
	
	ret = not rsAux.eof	
	'Response.Write "<hr>a(" & CInt(pAreaActual) & "," & CInt(rsAux("IDAREA")) & ")(" & CInt(pDetalleActual) & "," & CInt(rsAux("IDDETALLE")) & ")"
	if (ret) then ret = ((CInt(pAreaActual) = CInt(rsAux("IDAREA"))) and (CInt(pDetalleActual) = CInt(rsAux("IDDETALLE"))))
	corteControlGastos = ret
End Function
'--------------------------------------------------------------------------------------------------
'Funcion responsable de evaluar si hay datos validos y cumplen la condición de corte de los datos del area en el resumen.
Function corteControlArea (rs, pAreaActual)
	Dim ret 
	ret = not rs.eof	
	if (ret) then ret = (CInt(pAreaActual) = CInt(rs("IDAREA")))	
	corteControlArea  = ret
End Function	
'--------------------------------------------------------------------------------------------------
Function dibujarTotalesTrimestre(pTotalMes1, pTotalMes2, pTotalMes3, pTotalBudget)
	Dim myDesv, myDesvPerc, nroLineaTemp
	'Se imprimen los totales generales.
	nroLineaTemp = 97+(gNroLinea*10)	
	'Estructuras
	Call GF_squareBox(oPDF, 10, nroLineaTemp,205,20 ,0 ,GRIS,NEGRO ,1 ,0)		
	Call GF_squareBox(oPDF, 205,nroLineaTemp,60, 20 ,0 ,GRIS,NEGRO ,1 ,0) '1er mes
	Call GF_squareBox(oPDF, 265,nroLineaTemp,60, 20 ,0 ,GRIS,NEGRO ,1 ,0) '2do mes
	Call GF_squareBox(oPDF, 325,nroLineaTemp,60, 20 ,0 ,GRIS,NEGRO ,1 ,0) '3er mes
	Call GF_squareBox(oPDF, 385,nroLineaTemp,60, 20 ,0 ,GRIS,NEGRO ,1 ,0) 'total
	Call GF_squareBox(oPDF, 445,nroLineaTemp,70, 20 ,0 ,GRIS,NEGRO ,1 ,0) 'budget
	Call GF_squareBox(oPDF, 515,nroLineaTemp,40, 20 ,0 ,GRIS,NEGRO ,1 ,0) 'desv
	Call GF_squareBox(oPDF, 555,nroLineaTemp,30, 20 ,0 ,GRIS,NEGRO ,1 ,0) 'desv %
	Call GF_setFont(oPDF,"COURIER", 6,FONT_STYLE_BOLD)
	Call GF_writeTextAlign(oPDF,10, 104+(gNroLinea*10), GF_TRADUCIR("Total") , 205,PDF_ALIGN_CENTER)
	'Totales
	Call GF_writeTextAlign(oPDF,205, nroLineaTemp + 8, GF_EDIT_DECIMALS(pTotalMes1,2) , 55,PDF_ALIGN_RIGHT) '1er mes
	Call GF_writeTextAlign(oPDF,265, nroLineaTemp + 8, GF_EDIT_DECIMALS(pTotalMes2,2) , 55,PDF_ALIGN_RIGHT) '2do mes
	Call GF_writeTextAlign(oPDF,325, nroLineaTemp + 8, GF_EDIT_DECIMALS(pTotalMes3,2) , 55,PDF_ALIGN_RIGHT) '3er mes
	Call GF_writeTextAlign(oPDF,385, nroLineaTemp + 8, GF_EDIT_DECIMALS(pTotalMes1+pTotalMes2+pTotalMes3,2) , 55,PDF_ALIGN_RIGHT) 'Total		
	Call GF_writeTextAlign(oPDF,445, nroLineaTemp + 8, GF_EDIT_DECIMALS(pTotalBudget,2) , 65,PDF_ALIGN_RIGHT) 'budget
	
	myDesv = pTotalBudget - (pTotalMes1+pTotalMes2+pTotalMes3)
	if (myDesv > 0) then myDesv = 0
	myDesvPerc  = 0
	if (auxBudget > 0) then myDesvPerc = (myDesv/pTotalBudget)*100
	
	Call GF_writeTextAlign(oPDF,515, nroLineaTemp + 8, GF_EDIT_DECIMALS(abs(myDesv),2) , 35,PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,555, nroLineaTemp + 8, GF_EDIT_DECIMALS(abs(myDesvPerc),0) & " %" , 25,PDF_ALIGN_RIGHT)
	
End Function
'--------------------------------------------------------------------------------------------------
Function dibujarTotalesResumen(pTotalTrim1, pTotalTrim2, pTotalTrim3, pTotalTrim4, pTotalBudget)
	Dim myDesv, myDesvPerc, nroLineaTemp
		
	'Se imprimen los totales generales.
	nroLineaTemp = 98+(gNroLinea*10)
	'Estructuras
	Call GF_squareBox(oPDF, 5  ,nroLineaTemp,292,20 ,0 ,GRIS,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, 227,nroLineaTemp,50 ,20 ,0 ,GRIS,NEGRO ,1 ,0) '1er Trim
	Call GF_squareBox(oPDF, 277,nroLineaTemp,50 ,20 ,0 ,GRIS,NEGRO ,1 ,0) '2do Trim
	Call GF_squareBox(oPDF, 327,nroLineaTemp,50 ,20 ,0 ,GRIS,NEGRO ,1 ,0) '3er Trim
	Call GF_squareBox(oPDF, 377,nroLineaTemp,50 ,20 ,0 ,GRIS,NEGRO ,1 ,0) '4to Trim
	Call GF_squareBox(oPDF, 427,nroLineaTemp,53 ,20 ,0 ,GRIS,NEGRO ,1 ,0) 'Total actual
	Call GF_squareBox(oPDF, 480,nroLineaTemp,50 ,20 ,0 ,GRIS,NEGRO ,1 ,0) 'Budget
	Call GF_squareBox(oPDF, 530,nroLineaTemp,30 ,20 ,0 ,GRIS,NEGRO ,1 ,0) 'Desv
	Call GF_squareBox(oPDF, 560,nroLineaTemp,30 ,20 ,0 ,GRIS,NEGRO ,1 ,0) 'desv %
	Call GF_setFont(oPDF,"COURIER", 6,FONT_STYLE_BOLD)	
	Call GF_writeTextAlign(oPDF,5, nroLineaTemp + 8, GF_TRADUCIR("Total") , 205,PDF_ALIGN_CENTER)
	'Datos
	Call GF_writeTextAlign(oPDF,227, nroLineaTemp + 8, GF_EDIT_DECIMALS(pTotalTrim1,2), 45,PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,277, nroLineaTemp + 8, GF_EDIT_DECIMALS(pTotalTrim2,2), 45,PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,327, nroLineaTemp + 8, GF_EDIT_DECIMALS(pTotalTrim3,2), 45,PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,377, nroLineaTemp + 8, GF_EDIT_DECIMALS(pTotalTrim4,2), 45,PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,427, nroLineaTemp + 8, GF_EDIT_DECIMALS(pTotalTrim1 + pTotalTrim2 + pTotalTrim3 + pTotalTrim4,2), 50,PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,480, nroLineaTemp + 8, GF_EDIT_DECIMALS(pTotalBudget,2), 45,PDF_ALIGN_RIGHT)
	'Desvios
	myDesv = pTotalBudget - (pTotalTrim1 + pTotalTrim2 + pTotalTrim3 + pTotalTrim4)
	if (myDesv > 0) then myDesv = 0
	myDesvPerc  = 0
	if (totalGeneralBudget > 0) then myDesvPerc = (myDesv/totalGeneralBudget)*100
	Call GF_writeTextAlign(oPDF,530, nroLineaTemp + 8, GF_EDIT_DECIMALS(abs(myDesv),2) , 25,PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,560, nroLineaTemp + 8, GF_EDIT_DECIMALS(abs(myDesvPerc),0) & " %" , 25,PDF_ALIGN_RIGHT)
	
End Function
'--------------------------------------------------------------------------------------------------
Function dibujarTrimestre(pTrim)	
	Dim descArea,auxMes1,auxMes2,auxMes3,campoImporte, rsGastos
	Dim colorFondo, areaActual, detalleActual, trimestreActual, lineaArea
	Dim totalAreaMes1, totalAreaMes2, totalAreaMes3, rsObraTrim
	Dim totalGeneralMes1, totalGeneralMes2, totalGeneralMes3, totalGeneralBudget,auxBudget

	if (not primeraHoja) then 
		Call agregarHoja(pTrim)	
	end if
	primeraHoja = False
	Call dibujarEncabezado(pTrim)
	
	Call executeQueryDb(DBSITE_SQL_INTRA, rsGastos, "OPEN", generarSQLGastos(pTrim))
	
	strSQL = "select * from tblbudgetobrasdetalle a " & _
			 " where a.idobra = "&gIdObra&" and a.periodo  = " & pTrim & _
			 " order by a.idarea,a.iddetalle"
    Call executeQueryDb(DBSITE_SQL_INTRA, rsObraTrim, "OPEN", strSQL)
			
	'Se setea el campo de importe a utilizar según la moneda en que se imprime el reporte.
	campoImporte = "dlbudget"
	if (gMoneda = MONEDA_PESO) then campoImporte = "psbudget"
	
	Call dibujarTituloTrimestre(pTrim)
	
	while (not gRsObra.EoF)
		'Recorro todos los items del presupuesto.
		areaActual = gRsObra("IDAREA")	'Se setea el area de trabajo 		
		totalAreaMes1= 0
		totalAreaMes2= 0
		totalAreaMes3= 0		
		totalAreaBudget= 0
		while (corteControlArea (gRsObra, areaActual)) 'Se procesan los gastos del area de trabajo.
			if (CInt(gRsObra("IDDETALLE")) <> 0) then	'Solo proceso si el registro corresponde a un item de detalle.
				'Es una linea de detalle.
				detalleActual = gRsObra("IDDETALLE")	'Se setea el detalle de trabajo. (con esto queda definida la clave area-detalle
				auxMes1 = 0
				auxMes2 = 0
				auxMes3 = 0
				'Se evalua la condición de corte. Se deben procesar todos los gastos del area-detalle.
				'Se hace así debido a que si el rs se quedo sin registros no deben evaluarse				
				while (corteControlGastos(rsGastos, areaActual, detalleActual)) 'Se procesan todos los gastos del area-detalle
					Select case CInt(rsGastos("MES"))
						case MES_1
							auxMes1 = auxMes1 + CDbl(rsGastos("Gasto"))
						case MES_2
							auxMes2 = auxMes2 + CDbl(rsGastos("Gasto"))
						case MES_3
							auxMes3 = auxMes3 + CDbl(rsGastos("Gasto"))
					End Select
					rsGastos.MoveNext()
				wend 
				'Se imprime el detalle que se terminó de procesar.
				auxBudget = obtenerBudgetProporcionalDetalle(gFechaBudget,areaActual,detalleActual,rsObraTrim)
				gNroLinea = dibujarLineaTrimestre(false, gNroLinea, detalleActual, gRsObra("DSBUDGET"), auxMes1, auxMes2, auxMes3, auxBudget)
				totalAreaMes1= totalAreaMes1 + auxMes1
				totalAreaMes2= totalAreaMes2 + auxMes2
				totalAreaMes3= totalAreaMes3 + auxMes3
				totalAreaBudget= totalAreaBudget + auxBudget
				Call controlPaginacion(pTrim)
				gRsObra.MoveNext()
			else		
				'Es la primera linea que solo identifica el area.				
				descArea = gRsObra("DSBUDGET")
				paginaArea = gNroHojas
				lineaArea = gNroLinea
				gRsObra.MoveNext()				
				gNroLinea = gNroLinea+1 'Dar la linea por dibujada.(OJO! se dibuja realmente al final.--.
				Call controlPaginacion(pTrim)	'														|
			end if	'																					|
		wend		'																					|
		'Dibujo la linea del area.                                                                  <---·
		Call setWorkPage(oPDF, paginaArea)		
		Call dibujarLineaTrimestre(true, lineaArea, areaActual, descArea, totalAreaMes1, totalAreaMes2, totalAreaMes3, totalAreaBudget)
		Call setWorkPage(oPDF, gNroHojas)
		totalGeneralMes1= totalGeneralMes1 + totalAreaMes1
		totalGeneralMes2= totalGeneralMes2 + totalAreaMes2
		totalGeneralMes3= totalGeneralMes3 + totalAreaMes3		
		totalGeneralBudget= totalGeneralBudget + totalAreaBudget		
	wend		
	if (gChkPIC) then				
		Call loadTrimestreCTCSinAplicacion(pTrim, totalGeneralMes1, totalGeneralMes2, totalGeneralMes3, totalGeneralBudget)
	end if
	Call dibujarTotalesTrimestre(totalGeneralMes1, totalGeneralMes2, totalGeneralMes3, totalGeneralBudget)		
	'Se deja en recordset preparado para la siguiente hoja (Trimestre)
	gRsObra.MoveFirst()
End Function
'-------------------------------------------------CNA---------------------------------------------------------------
'*******************************************************************************************************************
'-------------------------------------------------------------------------------------------------------------------
Function loadResumenCTCSinAplicacion(ByRef ptotalGeneralTrimestre1,ByRef ptotalGeneralTrimestre2,ByRef ptotalGeneralTrimestre3,ByRef ptotalGeneralTrimestre4, totalGeneralBudget)
	Dim rs,  auxTrim_1, auxTrim_2, auxTrim_3, auxTrim_4, totTrim_1, totTrim_2, totTrim_3, totTrim_4, nroLinea_old, myPago,pagina_old	
	
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", armarSQLCTC(RESUMEN))
	if not rs.eof then		
		nroLinea_old = gNroLinea		
		pagina_old = gNroHojas
		gNroLinea = gNroLinea + 1
		totTrim_1 = 0
		totTrim_2 = 0
		totTrim_3 = 0
		totTrim_4 = 0		
		while (not rs.EoF)			
			if(isnull(rs("PAGO")))then 
				myPago = 0			
			else
				myPago = CDbl(rs("PAGO"))
			end if
			auxTrim_1 = 0
			auxTrim_2 = 0
			auxTrim_3 = 0
			auxTrim_4 = 0 
			Select case CInt(rs("TRIMESTRE"))
				case TRIMESTRE_1
					auxTrim_1 = CDbl(rs("Gasto")) - myPago
					totTrim_1 = totTrim_1 + auxTrim_1
				case TRIMESTRE_2
					auxTrim_2 = CDbl(rs("Gasto")) - myPago
					totTrim_2 = totTrim_2 + auxTrim_2
				case TRIMESTRE_3
					auxTrim_3 = CDbl(rs("Gasto")) - myPago
					totTrim_3 = totTrim_3 + auxTrim_3
				case TRIMESTRE_4
					auxTrim_4 = CDbl(rs("Gasto")) - myPago
					totTrim_4 = totTrim_4 + auxTrim_4
			End Select	
			Call dibujarResumenCTCsinAplicacion(CInt(rs("IDCONTRATO")),rs("CDCONTRATO"),auxTrim_1,auxTrim_2,auxTrim_3,auxTrim_4)			
			gNroLinea = gNroLinea + 1			
			Call controlPaginacion(RESUMEN)			
			rs.MoveNext()						
		wend
		ptotalGeneralTrimestre1 = ptotalGeneralTrimestre1 + totTrim_1
		ptotalGeneralTrimestre2 = ptotalGeneralTrimestre2 + totTrim_2
		ptotalGeneralTrimestre3 = ptotalGeneralTrimestre3 + totTrim_3
		ptotalGeneralTrimestre4 = ptotalGeneralTrimestre4 + totTrim_4
		Call setWorkPage(oPDF, pagina_old)
		Call dibujarTotalResumenSinaplicacion(totTrim_1, totTrim_2, totTrim_3, totTrim_4, nroLinea_old)
		Call setWorkPage(oPDF, gNroHojas)
	end if	
End Function
'--------------------------------------------------------------------------------------------------
Function loadTrimestreCTCSinAplicacion(pTrim, ByRef totalGeneralMes1, ByRef totalGeneralMes2, ByRef totalGeneralMes3, totalGeneralBudget)
	Dim rs,  auxMes1, auxMes2, auxMes3, nroLinea_old, totMes1, totMes2, totMes3,pagina_old
	
	Dim conta
	conta = 0
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", armarSQLCTC(pTrim))
	if not rs.eof then		
		nroLinea_old = gNroLinea
		pagina_old = gNroHojas
		gNroLinea = gNroLinea + 1
		totMes1 = 0 
		totMes2 = 0 
		totMes3 = 0 
		while (not rs.EoF)
			if(isnull(rs("PAGO")))then 
				myPago = 0			
			else
				myPago = CDbl(rs("PAGO"))
			end if
			auxMes1 = 0
			auxMes2 = 0
			auxMes3 = 0			
			Select case CInt(rs("MES"))
				case MES_1
					auxMes1 = CDbl(rs("Gasto")) - myPago
					totMes1 = totMes1 + auxMes1
				case MES_2
					auxMes2 = CDbl(rs("Gasto")) - myPago
					totMes2 = totMes2 + auxMes2
				case MES_3
					auxMes3 = CDbl(rs("Gasto")) - myPago
					totMes3 = totMes3 + auxMes3
			End Select			
			Call dibujarTrimestreCTCsinAplicacion(CInt(rs("IDCONTRATO")),rs("CDCONTRATO"),auxMes1,auxMes2,auxMes3)			
			gNroLinea = gNroLinea + 1
			conta = conta + 1
			Call controlPaginacion(pTrim)
			if(conta < 40 )then rs.MoveFirst
			rs.MoveNext()
		wend
		totalGeneralMes1 = totalGeneralMes1 + totMes1
		totalGeneralMes2 = totalGeneralMes2 + totMes2
		totalGeneralMes3 = totalGeneralMes3 + totMes3
		Call setWorkPage(oPDF, pagina_old)		
		Call dibujarTotalTrimestreSinaplicacion(totMes1, totMes2, totMes3, nroLinea_old)		
		Call setWorkPage(oPDF, gNroHojas)
	end if	
End Function
'--------------------------------------------------------------------------------------------------
Function armarSQLCTC(pTrim)
	Dim myMoneda, strSQL,myFecha,filtroTrimestre, myMonedaPagos
	
	myFecha = CLng(left(session("MmtoDato"),8))	
	if (gFechaHasta <> "") then	myFecha = CLng(GF_DTE2FN(gFechaHasta))
	myMoneda = "IMPORTEDOLARES"
	myMonedaPagos = "IMPORTEDOLARES"
	if (gMoneda = MONEDA_PESO) then 
		myMoneda = "IMPORTEPESOS"
		myMonedaPagos = "IMPORTEPESOS"
	end if	
	filtroTrimestre = ""
	if (pTrim <> RESUMEN) then filtroTrimestre = " WHERE TRIMESTRE = " & pTrim 	
	strSQL =  "				SELECT TRIMESTRE, "
	if (pTrim = RESUMEN) then
		strSQL = strSQL & "    MES     ,"
	else		
		strSQL = strSQL & "	   (MES -(3 * " & pTrim & ")) MES ,"
	end if
    strSQL = strSQL & "		   GASTO,        "
    strSQL = strSQL & "		   PAGO,	     "
    strSQL = strSQL & "		   CDCONTRATO,   " 
    strSQL = strSQL & "		   IDCONTRATO    "    
	strSQL = strSQL & "		FROM (            "  
	strSQL = strSQL & "			  SELECT CAST((CAST(SUBSTRING(cast(mmtoconf as varchar(15)),5,2) AS INTEGER)-1)/3 AS INTEGER) TRIMESTRE ,  "
    strSQL = strSQL & "					 SUBSTRING(cast(mmtoconf as varchar(15)),5,2)	MES     ,  "    
    strSQL = strSQL & "                  SUM(" & myMoneda & ")	GASTO , "			
	strSQL = strSQL & "					 CDCONTRATO				CDCONTRATO ,  "
	strSQL = strSQL & "			         CTC.IDCONTRATO				IDCONTRATO ,  "
	strSQL = strSQL & "                  PAGOS.SUMAPAGOS			PAGO  "
    strSQL = strSQL & "			  FROM (Select * from TBLOBRACONTRATOS "
    strSQL = strSQL & "                 where   IDOBRA=" & gIdObra 
    strSQL = strSQL & "                         and ESTADO in (" & ESTADO_CTC_AUTORIZADO & ", " & ESTADO_CTC_FINALIZADO & ", " & ESTADO_CTC_EN_AJUSTE & ")"
    strSQL = strSQL & "					        and IDAREA = 0 AND IDDETALLE = 0    "
    strSQL = strSQL & "			                and MMTOCONF <= " & myFecha & "235959 "
    strSQL = strSQL & "                 ) CTC  "
	strSQL = strSQL & "			  LEFT JOIN ( SELECT IDCONTRATO,"
	strSQL = strSQL & "								 SUM(" & myMonedaPagos & ") AS SUMAPAGOS	"
	strSQL = strSQL & "						  FROM   TBLCTZCABECERA CTZ "
	strSQL = strSQL & "						  WHERE CTZ.ESTADO IN ('" & CTZ_EN_FIRMA & "', '" & CTZ_FIRMADA & "') and IDCONTRATO <> 0 and IDOBRA=" & gIdObra
	strSQL = strSQL & "						  GROUP BY IDCONTRATO "
	strSQL = strSQL & "						  ) AS PAGOS " 
	strSQL = strSQL & "				 ON PAGOS.IDCONTRATO = CTC.IDCONTRATO " 
    strSQL = strSQL & "			  GROUP BY CAST((CAST(SUBSTRING(cast(mmtoconf as varchar(15)),5,2) AS INTEGER)-1)/3 AS INTEGER),  "
    strSQL = strSQL & "			          SUBSTRING(cast(mmtoconf as varchar(15)),5,2),  "
    strSQL = strSQL & "			  	      CDCONTRATO ,  "
    strSQL = strSQL & "			  	      CTC.IDCONTRATO ,   "
    strSQL = strSQL & "			  	      PAGOS.SUMAPAGOS   "    
    strSQL = strSQL & "			) AS T1  "
    strSQL = strSQL &			filtroTrimestre
	strSQL = strSQL & "	   ORDER BY TRIMESTRE, "
    strSQL = strSQL & "				MES, "
    strSQL = strSQL & "		        GASTO, "
	strSQL = strSQL & "			    CDCONTRATO, "
    strSQL = strSQL & "		        IDCONTRATO, "
    strSQL = strSQL & "		        PAGO "
    
	'Response.Write strSQL
	'	Response.End	
	armarSQLCTC = strSQL
End Function
'------------------------------------------------------------------------------------------------------------
Function dibujarTotalResumenSinaplicacion(pTotTrim_1, pTotTrim_2, pTotTrim_3, pTotTrim_4, pNroLinea)
	Dim nroLineaTemp
	nroLineaTemp = 98+(pNroLinea*10)		
	Call GF_setFont(oPDF,"COURIER", 6,FONT_STYLE_BOLD)			
	Call GF_squareBox(oPDF, 5   ,nroLineaTemp ,20  ,11 ,0 ,VERDE,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, 25  ,nroLineaTemp ,202 ,11 ,0 ,GRIS,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, 227 ,nroLineaTemp ,50  ,11 ,0 ,GRIS,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, 277 ,nroLineaTemp ,50  ,11 ,0 ,GRIS,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, 327 ,nroLineaTemp ,50  ,11 ,0 ,GRIS,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, 377 ,nroLineaTemp ,50  ,11 ,0 ,GRIS,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, 427 ,nroLineaTemp ,53  ,11 ,0 ,GRIS,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, 480 ,nroLineaTemp ,50  ,11 ,0 ,GRIS,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, 530 ,nroLineaTemp ,30  ,11 ,0 ,GRIS,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, 560 ,nroLineaTemp ,30  ,11 ,0 ,GRIS,NEGRO ,1 ,0)					
	Call GF_setFontColor(BLANCO)	
	Call GF_writeTextAlign(oPDF,5 , nroLineaTemp + 2 , GF_TRADUCIR("CTC"), 20,PDF_ALIGN_CENTER)
	Call GF_setFontColor(NEGRO)	
	Call GF_writeTextAlign(oPDF,30 , nroLineaTemp + 2 , GF_TRADUCIR("SALDO CONTRATO SIN APLICACION"), 202,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF,227 , nroLineaTemp + 2 ,  GF_EDIT_DECIMALS(pTotTrim_1,2), 45,PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,277 , nroLineaTemp + 2  , GF_EDIT_DECIMALS(pTotTrim_2,2), 45,PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,327 , nroLineaTemp + 2  , GF_EDIT_DECIMALS(pTotTrim_3,2), 45,PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,377 , nroLineaTemp + 2  , GF_EDIT_DECIMALS(pTotTrim_4,2), 45,PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,427 , nroLineaTemp + 2  , GF_EDIT_DECIMALS(0,2), 48,PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,480 , nroLineaTemp + 2  , GF_EDIT_DECIMALS(0,2), 45,PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,530 , nroLineaTemp + 2  , GF_EDIT_DECIMALS(0,2), 25,PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,560 , nroLineaTemp + 2  , "0 %" , 25,PDF_ALIGN_RIGHT)
	Call GF_setFontColor(NEGRO)	
end function
'-------------------------------------------------------------------------------------------------------------------
Function dibujarTotalTrimestreSinaplicacion(pTotMes_1, pTotMes_2, pTotMes_3, pNroLinea)
	Dim nroLineaTemp
	nroLineaTemp = 98+(pNroLinea*10)	
	Call GF_setFont(oPDF,"COURIER", 6,FONT_STYLE_BOLD)
	Call GF_squareBox(oPDF, 10   ,nroLineaTemp ,20  ,10 ,0 ,VERDE,NEGRO ,1 ,0)	
	Call GF_squareBox(oPDF, 30  ,nroLineaTemp ,175 ,10 ,0 ,GRIS,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, 205 ,nroLineaTemp ,60  ,10 ,0 ,GRIS,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, 265 ,nroLineaTemp ,60  ,10 ,0 ,GRIS,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, 325 ,nroLineaTemp ,60  ,10 ,0 ,GRIS,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, 385 ,nroLineaTemp ,60  ,10 ,0 ,GRIS,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, 445 ,nroLineaTemp ,70  ,10 ,0 ,GRIS,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, 515 ,nroLineaTemp ,40  ,10 ,0 ,GRIS,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, 555 ,nroLineaTemp ,30  ,10 ,0 ,GRIS,NEGRO ,1 ,0)	
	Call GF_setFontColor(BLANCO)	
	Call GF_writeTextAlign(oPDF,10 , nroLineaTemp + 2 , GF_TRADUCIR("CTC"), 20,PDF_ALIGN_CENTER)
	Call GF_setFontColor(NEGRO)
	Call GF_writeTextAlign(oPDF,35 , nroLineaTemp + 2 , GF_TRADUCIR("SALDO CONTRATO SIN APLICACION"), 175,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF,205 , nroLineaTemp + 2 , GF_EDIT_DECIMALS(pTotMes_1,2), 55,PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,265 , nroLineaTemp + 2 , GF_EDIT_DECIMALS(pTotMes_2,2), 55,PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,325 , nroLineaTemp + 2, GF_EDIT_DECIMALS(pTotMes_3,2), 55,PDF_ALIGN_RIGHT)			
	Call GF_writeTextAlign(oPDF,385 , nroLineaTemp + 2, GF_EDIT_DECIMALS(0,2), 55,PDF_ALIGN_RIGHT)			
	Call GF_writeTextAlign(oPDF,445 , nroLineaTemp + 2, GF_EDIT_DECIMALS(0,2), 65,PDF_ALIGN_RIGHT)			
	Call GF_writeTextAlign(oPDF,515 , nroLineaTemp + 2, GF_EDIT_DECIMALS(0,2), 35,PDF_ALIGN_RIGHT)			
	Call GF_writeTextAlign(oPDF,555 , nroLineaTemp + 2, "0 %", 25,PDF_ALIGN_RIGHT)					
	Call GF_setFontColor(NEGRO)	
end function
'-------------------------------------------------------------------------------------------------------------------
Function dibujarResumenCTCsinAplicacion(pIdContrato, pCdContrato, pImpTrim1, pImpTrim2, pImpTrim3, pImpTrim4)
	Dim nroLinea
	nroLinea = 98+(gNroLinea*10)
	Call GF_squareBox(oPDF, 25  ,  nroLinea,20  ,10 ,0 ,VERDE ,NEGRO ,1 ,0)	'IdContrato
	Call GF_squareBox(oPDF, 45 ,  nroLinea,182 ,10 ,0 ,BLANCO  ,NEGRO ,1 ,0)	'CdContrato	
	Call GF_squareBox(oPDF, 227,  nroLinea,50  ,10 ,0 ,BLANCO ,NEGRO ,1 ,0) '1er Trim
	Call GF_squareBox(oPDF, 277,  nroLinea,50  ,10 ,0 ,BLANCO ,NEGRO ,1 ,0) '2do Trim
	Call GF_squareBox(oPDF, 327,  nroLinea,50  ,10 ,0 ,BLANCO ,NEGRO ,1 ,0) '3er Trim
	Call GF_squareBox(oPDF, 377,  nroLinea,50  ,10 ,0 ,BLANCO ,NEGRO ,1 ,0) '4to Trim	
	Call GF_squareBox(oPDF, 427,  nroLinea,53  ,10 ,0 ,BLANCO ,NEGRO ,1 ,0) 'Total actual
	Call GF_squareBox(oPDF, 480,  nroLinea,50  ,10 ,0 ,BLANCO ,NEGRO ,1 ,0) 'Budget
	Call GF_squareBox(oPDF, 530,  nroLinea,30  ,10 ,0 ,BLANCO ,NEGRO ,1 ,0) 'Desv
	Call GF_squareBox(oPDF, 560,  nroLinea,30  ,10 ,0 ,BLANCO ,NEGRO ,1 ,0) 'desv %
	Call GF_setFont(oPDF,"COURIER", 6,FONT_STYLE_BOLD)
	Call GF_setFontColor(BLANCO)
	Call GF_writeTextAlign(oPDF,25, nroLinea + 2 , pIdContrato, 20,PDF_ALIGN_CENTER)
	Call GF_setFontColor(NEGRO)
	Call GF_writeTextAlign(oPDF,50, nroLinea + 2 , pCdContrato, 182,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF,227, nroLinea + 2, GF_EDIT_DECIMALS(pImpTrim1,2), 45,PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,277, nroLinea + 2, GF_EDIT_DECIMALS(pImpTrim2,2), 45,PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,327, nroLinea + 2, GF_EDIT_DECIMALS(pImpTrim3,2), 45,PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,377, nroLinea + 2, GF_EDIT_DECIMALS(pImpTrim4,2), 45,PDF_ALIGN_RIGHT)	
	Call GF_writeTextAlign(oPDF,427, nroLinea + 2, GF_EDIT_DECIMALS(0,2), 48,PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,480, nroLinea + 2, GF_EDIT_DECIMALS(0,2), 45,PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,530, nroLinea + 2, GF_EDIT_DECIMALS(0,2), 25,PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,560, nroLinea + 2, "0 %", 25,PDF_ALIGN_RIGHT)	
End Function
'------------------------------------------------------------------------------------------------------------------
Function dibujarTrimestreCTCsinAplicacion(pIdContrato, pCdContrato, pImpMes1, pImpMes2, pImpMes3)
	Dim nroLinea 
	nroLinea = 98+(gNroLinea*10)
	Call GF_squareBox(oPDF, 30 ,  nroLinea,20  ,10 ,0 ,VERDE ,NEGRO ,1 ,0)	'IdContrato
	Call GF_squareBox(oPDF, 50 ,  nroLinea,155 ,10 ,0 ,BLANCO  ,NEGRO ,1 ,0)	'CdContrato	
	Call GF_squareBox(oPDF, 205,  nroLinea,60  ,10 ,0 ,BLANCO ,NEGRO ,1 ,0) '1er Mes
	Call GF_squareBox(oPDF, 265,  nroLinea,60  ,10 ,0 ,BLANCO ,NEGRO ,1 ,0) '2do Mes
	Call GF_squareBox(oPDF, 325,  nroLinea,60  ,10 ,0 ,BLANCO ,NEGRO ,1 ,0) '3er Mes	
	Call GF_squareBox(oPDF, 385,  nroLinea,60  ,10 ,0 ,BLANCO ,NEGRO ,1 ,0) 'Total actual
	Call GF_squareBox(oPDF, 445,  nroLinea,70  ,10 ,0 ,BLANCO ,NEGRO ,1 ,0) 'Budget
	Call GF_squareBox(oPDF, 515,  nroLinea,40  ,10 ,0 ,BLANCO ,NEGRO ,1 ,0) 'Desv
	Call GF_squareBox(oPDF, 555,  nroLinea,30  ,10 ,0 ,BLANCO ,NEGRO ,1 ,0) 'desv %	
	Call GF_setFont(oPDF,"COURIER", 6,FONT_STYLE_BOLD)
	Call GF_setFontColor(BLANCO)
	Call GF_writeTextAlign(oPDF,30, nroLinea + 2, pIdContrato, 20,PDF_ALIGN_CENTER)
	Call GF_setFontColor(NEGRO)
	Call GF_writeTextAlign(oPDF,55, nroLinea + 2, pCdContrato, 155,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF,205, nroLinea + 2 , GF_EDIT_DECIMALS(pImpMes1,2), 55,PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,265, nroLinea + 2 , GF_EDIT_DECIMALS(pImpMes2,2), 55,PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,325, nroLinea + 2 , GF_EDIT_DECIMALS(pImpMes3,2), 55,PDF_ALIGN_RIGHT)	
	Call GF_writeTextAlign(oPDF,385, nroLinea + 2 , GF_EDIT_DECIMALS(0,2), 55,PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,445, nroLinea + 2 , GF_EDIT_DECIMALS(0,2), 65,PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,515, nroLinea + 2 , GF_EDIT_DECIMALS(0,2), 35,PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,555, nroLinea + 2 , "0 %", 25,PDF_ALIGN_RIGHT)
End Function
'-------------------------------------------------------------------------------------------------------------------
'*******************************************************************************************************************
'-------------------------------------------------CNA---------------------------------------------------------------
Function dibujarResumen()	
	Dim descArea,auxTrim1,auxTrim2,auxTrim3,auxTrim4,auxBudget,campoImporte,hayDatos
	Dim colorFondo, areaActual, detalleActual, trimestreActual, lineaArea
	Dim totalAreaTrimestre1, totalAreaTrimestre2, totalAreaTrimestre3, totalAreaTrimestre4
	Dim totalGeneralTrimestre1, totalGeneralTrimestre2, totalGeneralTrimestre3, totalGeneralTrimestre4, totalGeneralBudget
	
	primeraHoja = FALSE
	
	Call executeQueryDb(DBSITE_SQL_INTRA, rsGastos, "OPEN", generarSQLGastos(RESUMEN))
	
	Call dibujarEncabezado(RESUMEN)
	
	'Se setea el campo de importe a utilizar según la moneda en que se imprime el reporte.
	campoImporte = "dlbudget"
	if (gMoneda = MONEDA_PESO) then campoImporte = "psbudget"
	
	Call dibujarTitulosResumen()	
		
	while (not gRsObra.EoF)
		'Recorro todos los items del presupuesto.
		areaActual = gRsObra("IDAREA")	'Se setea el area de trabajo 		
		totalAreaTrimestre1= 0
		totalAreaTrimestre2= 0
		totalAreaTrimestre3= 0
		totalAreaTrimestre4= 0
		totalAreaBudget= 0
		
		strSQL ="select * from tblbudgetobrasdetalle a " & _
					" where a.idobra = "&gIdObra & _
					" and a.idarea = "&areaActual & _
					" order by a.idarea,a.iddetalle"
            Call executeQueryDb(DBSITE_SQL_INTRA, gRsResu, "OPEN", strSQL)
		
		while (corteControlArea(gRsObra, areaActual)) 'Se procesan los gastos del area de trabajo.			
			
			if (CInt(gRsObra("IDDETALLE")) <> 0) then	'Solo proceso si el registro corresponde a un item de detalle.
				'Es una linea de detalle.
				detalleActual = gRsObra("IDDETALLE")	'Se setea el detalle de trabajo. (con esto queda definida la clave area-detalle
				auxTrim1 = 0
				auxTrim2 = 0
				auxTrim3 = 0
				auxTrim4 = 0
				'Se evalua la condición de corte. Se procesan los gastos del area-detalle de trabajo.
				'Se hace así debido a que si el rs se quedo sin registros no deben evaluarse								
				while (corteControlGastos(rsGastos, areaActual, detalleActual)) 'Se procesan todos los gastos del area-detalle
					Select case rsGastos("TRIMESTRE")
						case TRIMESTRE_1
							auxTrim1 = auxTrim1 + CDbl(rsGastos("Gasto"))
						case TRIMESTRE_2
							auxTrim2 = auxTrim2 + CDbl(rsGastos("Gasto"))
						case TRIMESTRE_3
							auxTrim3 = auxTrim3 + CDbl(rsGastos("Gasto"))
						case TRIMESTRE_4
							auxTrim4 = auxTrim4 + CDbl(rsGastos("Gasto"))
					End Select
					rsGastos.MoveNext()
				wend 				
				'Se imprime el detalle que se terminó de procesar.
				auxBudget = obtenerBudgetProporcionalResumen(gRsResu,gFechaBudget,detalleActual)
				gNroLinea = dibujarLineaResumen(false, gNroLinea, detalleActual, gRsObra("DSBUDGET"), gRsObra("CDCUENTA"), gRsObra("CCOSTOS"), auxTrim1, auxTrim2, auxTrim3, auxTrim4, CDbl(auxBudget), gRsObra("DSDETALLE"))
				totalAreaTrimestre1= totalAreaTrimestre1 + auxTrim1
				totalAreaTrimestre2= totalAreaTrimestre2 + auxTrim2
				totalAreaTrimestre3= totalAreaTrimestre3 + auxTrim3
				totalAreaTrimestre4= totalAreaTrimestre4 + auxTrim4
				totalAreaBudget= totalAreaBudget + CDbl(auxBudget)
				Call controlPaginacion(RESUMEN)
				gRsObra.MoveNext()
			else		
				'Es la primera linea que solo identifica el area.				
				descArea = gRsObra("DSBUDGET")
				paginaArea = gNroHojas
				lineaArea = gNroLinea
				gRsObra.MoveNext()				
				gNroLinea = gNroLinea+1 'Dar la linea por dibujada.(OJO! se dibuja realmente al final.--.
				Call controlPaginacion(RESUMEN)	'														|
			end if	'																					|
		wend		'																					|
		
		'Dibujo la linea del area.                                                                  <---·
		Call setWorkPage(oPDF, paginaArea)		
		Call dibujarLineaResumen(true, lineaArea, areaActual, descArea, "", "", totalAreaTrimestre1, totalAreaTrimestre2, totalAreaTrimestre3, totalAreaTrimestre4, totalAreaBudget, "")
		Call setWorkPage(oPDF, gNroHojas)
		totalGeneralTrimestre1= totalGeneralTrimestre1 + totalAreaTrimestre1
		totalGeneralTrimestre2= totalGeneralTrimestre2 + totalAreaTrimestre2
		totalGeneralTrimestre3= totalGeneralTrimestre3 + totalAreaTrimestre3
		totalGeneralTrimestre4= totalGeneralTrimestre4 + totalAreaTrimestre4
		totalGeneralBudget= totalGeneralBudget + totalAreaBudget		
	wend	
	if (gChkPIC) then				
		Call loadResumenCTCSinAplicacion(totalGeneralTrimestre1, totalGeneralTrimestre2, totalGeneralTrimestre3, totalGeneralTrimestre4, totalGeneralBudget)
	end if	
	Call dibujarTotalesResumen(totalGeneralTrimestre1, totalGeneralTrimestre2, totalGeneralTrimestre3, totalGeneralTrimestre4, totalGeneralBudget)	
	gNroLinea = gNroLinea + 2 	
	Call dibujarNotas()
	Call DibujarReasignaciones()
    Call dibujarAjustesPartidaPresupuestaria()
	'Se deja en recordset preparado para la siguiente hoja (Trimestre)
	gRsObra.MoveFirst()
End Function
'--------------------------------------------------------------------------------------------------
Function dibujarNotas()
	if (gCantComentarios > 0) then
	
		Call GF_setFont(oPDF,"COURIER", 6,FONT_STYLE_BOLD)
		Call GF_writeTextAlign(oPDF,10, 98+(gNroLinea*10), GF_TRADUCIR("NOTAS") & ":" , 580,PDF_ALIGN_LEFT)
		gNroLinea = gNroLinea+1
	
		Call GF_setFont(oPDF,"COURIER", 6,FONT_STYLE_NORMAL)	
		for i = 1 to ubound(gVecComentarios)
			Call GF_writeTextAlign(oPDF,10, 98+(gNroLinea*10), i&": "& gVecComentarios(i) , 580,PDF_ALIGN_LEFT)
			gNroLinea = gNroLinea +1
			Call controlPaginacion(NOTAS_REASIG)
		next
		
	end if
End Function
'--------------------------------------------------------------------------------------------------
' Autor: 	GFG - Guido Fonticelli
' Fecha: 	01/01/11
' Objetivo:	
'			Dibuja las reasignaciones que haya tenido la obra
' Devuelve:
'			Nada
' Modificaciones:
'			06/01/11 - GFG
'--------------------------------------------------------------------------------------------------
Function DibujarReasignaciones()
	Dim myReasig,rsFirmas
	
	set myReasig = obtenerReasignaciones(gIdObra,gHoy)
	
	if (not myReasig.EoF)	then
		
		Call GF_squareBox(oPDF, 5  ,100+(gNroLinea*10),585 ,15 ,0 ,VERDE,NEGRO ,1 ,0)
		Call GF_setFontColor(BLANCO)
		Call GF_setFont(oPDF,"ARIAL", 8 , FONT_STYLE_BOLD)
		Call GF_writeTextAlign(oPDF,5 , 103+(gNroLinea*10), GF_TRADUCIR("REASIGNACIONES") , 585,PDF_ALIGN_CENTER)
		
		gNroLinea = gNroLinea +1
		
		
		while not myReasig.EoF
			Call GF_setFont(oPDF,"ARIAL", 6 , FONT_STYLE_BOLD)
			Call GF_setFontColor(BLANCO)
			Call GF_squareBox(oPDF, 5  ,105+(gNroLinea*10),75 ,10 ,0 ,VERDE,NEGRO ,1 ,0)
			Call GF_writeTextAlign(oPDF,5 , 107+(gNroLinea*10), GF_TRADUCIR("FECHA") , 75,PDF_ALIGN_CENTER)
			Call GF_squareBox(oPDF, 80  ,105+(gNroLinea*10),200 ,10 ,0 ,VERDE,NEGRO ,1 ,0)
			Call GF_writeTextAlign(oPDF,80 , 107+(gNroLinea*10), GF_TRADUCIR("AREA") , 200,PDF_ALIGN_CENTER)
			Call GF_squareBox(oPDF, 280  ,105+(gNroLinea*10),200 ,10 ,0 ,VERDE,NEGRO ,1 ,0)
			Call GF_writeTextAlign(oPDF,280 , 107+(gNroLinea*10), GF_TRADUCIR("PERIODO") , 200,PDF_ALIGN_CENTER)
			Call GF_squareBox(oPDF, 480  ,105+(gNroLinea*10),110 ,10 ,0 ,VERDE,NEGRO ,1 ,0)
			Call GF_writeTextAlign(oPDF,480 , 107+(gNroLinea*10), GF_TRADUCIR("MONTO") , 110,PDF_ALIGN_CENTER)
				
			gNroLinea = gNroLinea +1
			
			Call GF_setFontColor(NEGRO)
			Call GF_squareBox(oPDF, 5  ,105+(gNroLinea*10),75 ,10 ,0 ,BLANCO,NEGRO ,1 ,0)
			Call GF_writeTextAlign(oPDF,5 , 107+(gNroLinea*10), GF_FN2DTE(myReasig("fecha")) , 75,PDF_ALIGN_CENTER)
			Call GF_squareBox(oPDF, 80  ,105+(gNroLinea*10),200 ,10 ,0 ,BLANCO,NEGRO ,1 ,0)
			Call GF_writeTextAlign(oPDF,85 , 107+(gNroLinea*10), myReasig("idAreaOrigen") & " - " & myReasig("dsArea") , 190,PDF_ALIGN_LEFT)
			Call GF_squareBox(oPDF, 280  ,105+(gNroLinea*10),200 ,10 ,0 ,BLANCO,NEGRO ,1 ,0)
			if (not isnull(myReasig("periodo"))) then
				Call GF_writeTextAlign(oPDF,285 , 107+(gNroLinea*10), v_trim(myReasig("periodo")) , 190,PDF_ALIGN_LEFT)
			end if
			Call GF_squareBox(oPDF, 480  ,105+(gNroLinea*10),110 ,10 ,0 ,BLANCO,NEGRO ,1 ,0)
			if (gMoneda = MONEDA_DOLAR) then
				Call GF_writeTextAlign(oPDF,480 , 107+(gNroLinea*10), UCASE(getSimboloMoneda(gMoneda)) & " " & GF_EDIT_DECIMALS(myReasig("importedolares"),2) , 105,PDF_ALIGN_RIGHT)
			else
				Call GF_writeTextAlign(oPDF,480 , 107+(gNroLinea*10), UCASE(getSimboloMoneda(gMoneda)) & " " & GF_EDIT_DECIMALS(myReasig("importepesos"),2) , 105,PDF_ALIGN_RIGHT)
			end if
			
			gNroLinea = gNroLinea +1
			
			Call GF_setFontColor(BLANCO)
			Call GF_squareBox(oPDF, 5  ,105+(gNroLinea*10),45 ,10 ,0 ,VERDE,NEGRO ,1 ,0)
			Call GF_writeTextAlign(oPDF,5 , 107+(gNroLinea*10), GF_TRADUCIR("DESDE") , 45,PDF_ALIGN_CENTER)
			Call GF_squareBox(oPDF, 295  ,105+(gNroLinea*10),45 ,10 ,0 ,VERDE,NEGRO ,1 ,0)
			Call GF_writeTextAlign(oPDF,295 , 107+(gNroLinea*10), GF_TRADUCIR("HASTA") , 45,PDF_ALIGN_CENTER)
			
			Call GF_setFontColor(NEGRO)
			
			Call GF_squareBox(oPDF, 50  ,105+(gNroLinea*10),245 ,10 ,0 ,BLANCO,NEGRO ,1 ,0)
			Call GF_writeTextAlign(oPDF,55 , 107+(gNroLinea*10), myReasig("idDetaOrigen") & " - " & myReasig("detaOrigen") , 245,PDF_ALIGN_LEFT)
			Call GF_squareBox(oPDF, 340  ,105+(gNroLinea*10),250 ,10 ,0 ,BLANCO,NEGRO ,1 ,0)
			Call GF_writeTextAlign(oPDF,345 , 107+(gNroLinea*10), myReasig("idDetaDestino") & " - " & myReasig("detaDestino") , 245,PDF_ALIGN_LEFT)
				
			gNroLinea = gNroLinea +1
			
			Call GF_setFontColor(BLANCO)
			Call GF_squareBox(oPDF, 5  ,105+(gNroLinea*10),585 ,10 ,0 ,VERDE,NEGRO ,1 ,0)
			Call GF_writeTextAlign(oPDF,5 , 107+(gNroLinea*10), GF_TRADUCIR("MOTIVO") , 585,PDF_ALIGN_CENTER)
			
			gNroLinea = gNroLinea +1
			
			Call GF_setFontColor(NEGRO)
			Call GF_squareBox(oPDF, 5  ,105+(gNroLinea*10),585 ,20 ,0 ,BLANCO,NEGRO ,1 ,0)
			Call GF_writeTextPlus(oPDF,10, 107+(gNroLinea*10), myReasig("motivo"), 565, 8, PDF_ALIGN_LEFT)

            gNroLinea = gNroLinea +2
			
			Call GF_setFontColor(BLANCO)
			Call GF_squareBox(oPDF, 5  ,105+(gNroLinea*10),585 ,10 ,0 ,VERDE,NEGRO ,1 ,0)
			Call GF_writeTextAlign(oPDF,5 , 107+(gNroLinea*10), GF_TRADUCIR("AUTORIZACIONES") , 585,PDF_ALIGN_CENTER)
				
            gNroLinea = gNroLinea +1
			
            Call GF_setFontColor(NEGRO)
            Call executeProcedureDb(DBSITE_SQL_INTRA, rsFirmas, "TBLBUDGETREASIGNACIONFIRMAS_GET_BY_IDREASIGNACION", myReasig("idReasignacion"))
            Call GF_squareBox(oPDF, 5  ,105+(gNroLinea*10),292,10 ,0 ,BLANCO,NEGRO ,1 ,0)
            Call GF_squareBox(oPDF, 297  ,105+(gNroLinea*10),293,10 ,0 ,BLANCO,NEGRO ,1 ,0)
            if (not rsFirmas.Eof) then
                Call GF_writeTextPlus(oPDF,10, 107+(gNroLinea*10),getUserDescription(rsFirmas("CDUSUARIO")) & " - " & armarTextoPlanoFirma(rsFirmas("HKEY"), rsFirmas("MOMENTO")), 290, 8, PDF_ALIGN_LEFT)                
                rsFirmas.MoveNext()
            end if
            if (not rsFirmas.Eof) then
                Call GF_writeTextPlus(oPDF,302, 107+(gNroLinea*10),getUserDescription(rsFirmas("CDUSUARIO")) & " - " & armarTextoPlanoFirma(rsFirmas("HKEY"), rsFirmas("MOMENTO")), 290, 8, PDF_ALIGN_LEFT)
                rsFirmas.MoveNext()
            end if

			gNroLinea = gNroLinea + 4
			
			Call controlPaginacion(NOTAS_REASIG) 

			myReasig.MoveNext
		wend
	end if
	
End Function
'--------------------------------------------------------------------------------------------------
Function dibujarAjustesPartidaPresupuestaria()
    Dim rsAjs,rsFirmas
	set rsAjs = obtenerAjustePartidaPresupuestaria(gIdObra,gHoy)
	
    if (not rsAjs.EoF)	then
		
		Call GF_squareBox(oPDF, 5  ,100+(gNroLinea*10),585 ,15 ,0 ,VERDE,NEGRO ,1 ,0)
		Call GF_setFontColor(BLANCO)
		Call GF_setFont(oPDF,"ARIAL", 8 , FONT_STYLE_BOLD)
		Call GF_writeTextAlign(oPDF,5 , 103+(gNroLinea*10), GF_TRADUCIR("AJUSTES") , 585,PDF_ALIGN_CENTER)
		
	
    	gNroLinea = gNroLinea +1

		while not rsAjs.EoF
			Call GF_setFont(oPDF,"ARIAL", 6 , FONT_STYLE_BOLD)
			Call GF_setFontColor(BLANCO)
			Call GF_squareBox(oPDF, 5  ,105+(gNroLinea*10),75 ,10 ,0 ,VERDE,NEGRO ,1 ,0)
			Call GF_writeTextAlign(oPDF,5 , 107+(gNroLinea*10), GF_TRADUCIR("FECHA") , 75,PDF_ALIGN_CENTER)
			Call GF_squareBox(oPDF, 80  ,105+(gNroLinea*10),200 ,10 ,0 ,VERDE,NEGRO ,1 ,0)
			Call GF_writeTextAlign(oPDF,80 , 107+(gNroLinea*10), GF_TRADUCIR("AREA") , 200,PDF_ALIGN_CENTER)
			Call GF_squareBox(oPDF, 280  ,105+(gNroLinea*10),200 ,10 ,0 ,VERDE,NEGRO ,1 ,0)
			Call GF_writeTextAlign(oPDF,280 , 107+(gNroLinea*10), GF_TRADUCIR("PERIODO") , 200,PDF_ALIGN_CENTER)
			Call GF_squareBox(oPDF, 480  ,105+(gNroLinea*10),110 ,10 ,0 ,VERDE,NEGRO ,1 ,0)
			Call GF_writeTextAlign(oPDF,480 , 107+(gNroLinea*10), GF_TRADUCIR("MONTO") , 110,PDF_ALIGN_CENTER)
				
			gNroLinea = gNroLinea +1
			
			Call GF_setFontColor(NEGRO)
			Call GF_squareBox(oPDF, 5  ,105+(gNroLinea*10),75 ,10 ,0 ,BLANCO,NEGRO ,1 ,0)
			Call GF_writeTextAlign(oPDF,5 , 107+(gNroLinea*10), GF_FN2DTE(rsAjs("fecha")) , 75,PDF_ALIGN_CENTER)
			Call GF_squareBox(oPDF, 80  ,105+(gNroLinea*10),200 ,10 ,0 ,BLANCO,NEGRO ,1 ,0)
			Call GF_writeTextAlign(oPDF,85 , 107+(gNroLinea*10), rsAjs("idAreaDestino") & " - " & rsAjs("dsArea") , 190,PDF_ALIGN_LEFT)
			Call GF_squareBox(oPDF, 280  ,105+(gNroLinea*10),200 ,10 ,0 ,BLANCO,NEGRO ,1 ,0)
			if (not isnull(rsAjs("periodo"))) then
				Call GF_writeTextAlign(oPDF,285 , 107+(gNroLinea*10), v_trim(rsAjs("periodo")) , 190,PDF_ALIGN_LEFT)
			end if
			Call GF_squareBox(oPDF, 480  ,105+(gNroLinea*10),110 ,10 ,0 ,BLANCO,NEGRO ,1 ,0)
			if (gMoneda = MONEDA_DOLAR) then
				Call GF_writeTextAlign(oPDF,480 , 107+(gNroLinea*10), UCASE(getSimboloMoneda(gMoneda)) & " " & GF_EDIT_DECIMALS(rsAjs("importedolares"),2) , 105,PDF_ALIGN_RIGHT)
			else
				Call GF_writeTextAlign(oPDF,480 , 107+(gNroLinea*10), UCASE(getSimboloMoneda(gMoneda)) & " " & GF_EDIT_DECIMALS(rsAjs("importepesos"),2) , 105,PDF_ALIGN_RIGHT)
			end if
			
			gNroLinea = gNroLinea +1
			
			Call GF_setFontColor(BLANCO)
			Call GF_squareBox(oPDF, 5  ,105+(gNroLinea*10),45 ,10 ,0 ,VERDE,NEGRO ,1 ,0)
			Call GF_writeTextAlign(oPDF,5 , 107+(gNroLinea*10), GF_TRADUCIR("DETALLE") , 45,PDF_ALIGN_CENTER)
			
			Call GF_setFontColor(NEGRO)
			
			Call GF_squareBox(oPDF, 50  ,105+(gNroLinea*10),540,10 ,0 ,BLANCO,NEGRO ,1 ,0)
			Call GF_writeTextAlign(oPDF,55 , 107+(gNroLinea*10), rsAjs("idDetaDestino") & " - " & rsAjs("detaDestino") , 500,PDF_ALIGN_LEFT)
				
			gNroLinea = gNroLinea +1
			
			Call GF_setFontColor(BLANCO)
			Call GF_squareBox(oPDF, 5  ,105+(gNroLinea*10),585 ,10 ,0 ,VERDE,NEGRO ,1 ,0)
			Call GF_writeTextAlign(oPDF,5 , 107+(gNroLinea*10), GF_TRADUCIR("MOTIVO") , 585,PDF_ALIGN_CENTER)
			
			gNroLinea = gNroLinea +1
			
			Call GF_setFontColor(NEGRO)
			Call GF_squareBox(oPDF, 5  ,105+(gNroLinea*10),585 ,40 ,0 ,BLANCO,NEGRO ,1 ,0)
			Call GF_writeTextPlus(oPDF,10, 107+(gNroLinea*10), rsAjs("motivo"), 565, 8, PDF_ALIGN_LEFT)
	        
             gNroLinea = gNroLinea +2
			
			Call GF_setFontColor(BLANCO)
			Call GF_squareBox(oPDF, 5  ,105+(gNroLinea*10),585 ,10 ,0 ,VERDE,NEGRO ,1 ,0)
			Call GF_writeTextAlign(oPDF,5 , 107+(gNroLinea*10), GF_TRADUCIR("AUTORIZACIONES") , 585,PDF_ALIGN_CENTER)
				
            gNroLinea = gNroLinea +1
			
            Call GF_setFontColor(NEGRO)
            Call executeProcedureDb(DBSITE_SQL_INTRA, rsFirmas, "TBLBUDGETREASIGNACIONFIRMAS_GET_BY_IDREASIGNACION", rsAjs("idReasignacion"))
            Call GF_squareBox(oPDF, 5  ,105+(gNroLinea*10),292,10 ,0 ,BLANCO,NEGRO ,1 ,0)
            Call GF_squareBox(oPDF, 297  ,105+(gNroLinea*10),293,10 ,0 ,BLANCO,NEGRO ,1 ,0)
            if (not rsFirmas.Eof) then
                Call GF_writeTextPlus(oPDF,10, 107+(gNroLinea*10),getUserDescription(rsFirmas("CDUSUARIO")) & " - " & armarTextoPlanoFirma(rsFirmas("HKEY"), rsFirmas("MOMENTO")), 290, 8, PDF_ALIGN_LEFT)                
                rsFirmas.MoveNext()
            end if
            if (not rsFirmas.Eof) then
                Call GF_writeTextPlus(oPDF,302, 107+(gNroLinea*10),getUserDescription(rsFirmas("CDUSUARIO")) & " - " & armarTextoPlanoFirma(rsFirmas("HKEY"), rsFirmas("MOMENTO")), 290, 8, PDF_ALIGN_LEFT)
                rsFirmas.MoveNext()
            end if
			


			gNroLinea = gNroLinea + 4
			
			Call controlPaginacion(NOTAS_REASIG) 

			rsAjs.MoveNext
		wend
	end if
End Function
'--------------------------------------------------------------------------------------------------
Function controlPaginacion(pTrim)
	if (pTrim = NOTAS_REASIG) then gNroLinea = gNroLinea + 5
	if (gNroLinea >= MAX_PAGE_LINES) then
		Call agregarHoja(pTrim)		
		Call dibujarEncabezado(pTrim)

		Select case pTrim 
			case RESUMEN
				Call dibujarTitulosResumen()
			case else		
				if (pTrim <> NOTAS_REASIG) then	Call dibujarTituloTrimestre(pTrim)					
		End Select 
	else
		if (pTrim = NOTAS_REASIG) then gNroLinea = gNroLinea - 5
	end if
End Function
'--------------------------------------------------------------------------------------------------
Function dibujarReporte()
	
	if (gMostrarResumen) then call dibujarResumen()
	if (gMostrar1erTrim) then call dibujarTrimestre(TRIMESTRE_1)
	if (gMostrar2doTrim) then call dibujarTrimestre(TRIMESTRE_2)
	if (gMostrar3erTrim) then call dibujarTrimestre(TRIMESTRE_3)
	if (gMostrar4toTrim) then call dibujarTrimestre(TRIMESTRE_4)
	
End Function
'--------------------------------------------------------------------------------------------------
Function dibujarTitulosResumen()
		Dim nroLineaTemp
		nroLineaTemp = 98+(gNroLinea*10)
		Call GF_setFont(oPDF,"ARIAL", 7 , FONT_STYLE_BOLD)
		if(gChkContable)then			
			Call GF_squareBox(oPDF, 5  ,nroLineaTemp,160 ,20 ,0 ,VERDE,NEGRO ,1 ,0) 'Detalle
			Call GF_squareBox(oPDF, 165,nroLineaTemp,40 ,20 ,0 ,VERDE,NEGRO ,1 ,0) 'Cuenta
			Call GF_squareBox(oPDF, 205,nroLineaTemp,22 ,20 ,0 ,VERDE,NEGRO ,1 ,0) 'CCosto
		else
			Call GF_squareBox(oPDF, 5  ,nroLineaTemp,222 ,20 ,0 ,VERDE,NEGRO ,1 ,0) 'Detalle
		end if
		Call GF_squareBox(oPDF, 227,nroLineaTemp,50 ,20 ,0 ,VERDE,NEGRO ,1 ,0) '1er Trim
		Call GF_squareBox(oPDF, 277,nroLineaTemp,50 ,20 ,0 ,VERDE,NEGRO ,1 ,0) '2do Trim
		Call GF_squareBox(oPDF, 327,nroLineaTemp,50 ,20 ,0 ,VERDE,NEGRO ,1 ,0) '3er Trim
		Call GF_squareBox(oPDF, 377,nroLineaTemp,50 ,20 ,0 ,VERDE,NEGRO ,1 ,0) '4to Trim
		Call GF_squareBox(oPDF, 427,nroLineaTemp,53 ,20 ,0 ,VERDE,NEGRO ,1 ,0) 'Total actual
		Call GF_squareBox(oPDF, 480,nroLineaTemp,50 ,20 ,0 ,VERDE,NEGRO ,1 ,0) 'Budget
		Call GF_squareBox(oPDF, 530,nroLineaTemp,30 ,20 ,0 ,VERDE,NEGRO ,1 ,0) 'Desv
		Call GF_squareBox(oPDF, 560,nroLineaTemp,30 ,20 ,0 ,VERDE,NEGRO ,1 ,0) 'desv %
		
		Call GF_setFontColor(BLANCO)
		if(gChkContable)then
			'ESTA TILDADO , QUIERE QUE SE MUESTRE LA INFORMACION CONTABLE
			Call GF_writeTextAlign(oPDF,10 , 103+(gNroLinea*10), GF_TRADUCIR("Detalle")		, 220,PDF_ALIGN_LEFT)
			Call GF_writeTextAlign(oPDF,170, 103+(gNroLinea*10), GF_TRADUCIR("Cuenta")		, 30 ,PDF_ALIGN_CENTER)
			Call GF_writeTextAlign(oPDF,205, 103+(gNroLinea*10), GF_TRADUCIR("CC")			, 20 ,PDF_ALIGN_CENTER)
		else
			'NO ESTA TILDADO , SOLO SE AGRANDA EL CUADRO DE LA DESCRIPCION
			Call GF_writeTextAlign(oPDF,10 , 103+(gNroLinea*10), GF_TRADUCIR("Detalle")		, 270,PDF_ALIGN_LEFT)
		end if
		Call GF_writeTextAlign(oPDF,227, 101+(gNroLinea*10), GF_TRADUCIR("1er")			, 50 ,PDF_ALIGN_CENTER)
		Call GF_writeTextAlign(oPDF,227, 106+(gNroLinea*10), GF_TRADUCIR("Trimestre")	, 50 ,PDF_ALIGN_CENTER)
		Call GF_writeTextAlign(oPDF,277, 101+(gNroLinea*10), GF_TRADUCIR("2do")			, 50 ,PDF_ALIGN_CENTER)
		Call GF_writeTextAlign(oPDF,277, 106+(gNroLinea*10), GF_TRADUCIR("Trimestre")	, 50 ,PDF_ALIGN_CENTER)
		Call GF_writeTextAlign(oPDF,327, 101+(gNroLinea*10), GF_TRADUCIR("3er")			, 50 ,PDF_ALIGN_CENTER)
		Call GF_writeTextAlign(oPDF,327, 106+(gNroLinea*10), GF_TRADUCIR("Trimestre")	, 50 ,PDF_ALIGN_CENTER)
		Call GF_writeTextAlign(oPDF,377, 101+(gNroLinea*10), GF_TRADUCIR("4to")			, 50 ,PDF_ALIGN_CENTER)
		Call GF_writeTextAlign(oPDF,377, 106+(gNroLinea*10), GF_TRADUCIR("Trimestre")	, 50 ,PDF_ALIGN_CENTER)
		Call GF_writeTextAlign(oPDF,427, 101+(gNroLinea*10), GF_TRADUCIR("Total")		, 53 ,PDF_ALIGN_CENTER)
		Call GF_writeTextAlign(oPDF,427, 108+(gNroLinea*10), GF_TRADUCIR("Actual")		, 53 ,PDF_ALIGN_CENTER)
		Call GF_writeTextAlign(oPDF,479, 101+(gNroLinea*10), GF_TRADUCIR("Budget")		, 50 ,PDF_ALIGN_CENTER)
		Call GF_writeTextAlign(oPDF,479,108+(gNroLinea*10), UCASE(getSimboloMonedaLetras(gMoneda)) , 50,PDF_ALIGN_CENTER)
		Call GF_writeTextAlign(oPDF,529, 101+(gNroLinea*10), GF_TRADUCIR("Desv")			, 30 ,PDF_ALIGN_CENTER)
		Call GF_writeTextAlign(oPDF,529,108+(gNroLinea*10), UCASE(getSimboloMonedaLetras(gMoneda)) , 30,PDF_ALIGN_CENTER)
		Call GF_writeTextAlign(oPDF,559, 101+(gNroLinea*10), GF_TRADUCIR("Desv") , 30 ,PDF_ALIGN_CENTER)
		Call GF_writeTextAlign(oPDF,559, 108+(gNroLinea*10), GF_TRADUCIR("Desv") & "%"	, 30 ,PDF_ALIGN_CENTER)
		
		Call GF_setFontColor(NEGRO)
		gNroLinea = gNroLinea+2
end function
'--------------------------------------------------------------------------------------------------
' Autor: 	GFG - Guido Fonticelli
' Fecha: 	01/01/11
' Objetivo:	
'			Dibuja el titulo de los trimestres
' Parametros:
'			pTrim	[int] 
' Devuelve:
'			Nada
' Modificaciones:
'			06/01/11 - GFG
'--------------------------------------------------------------------------------------------------
Function dibujarTituloTrimestre(pTrim)
	Dim nroLineaTemp
	nroLineaTemp = 98+(gNroLinea*10)
	Call GF_squareBox(oPDF, 10   ,nroLineaTemp,195 ,20 ,0 ,VERDE,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, 205  ,nroLineaTemp,60  ,20 ,0 ,VERDE,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, 265  ,nroLineaTemp,60  ,20 ,0 ,VERDE,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, 325  ,nroLineaTemp,60  ,20 ,0 ,VERDE,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, 385  ,nroLineaTemp,60  ,20 ,0 ,VERDE,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, 445  ,nroLineaTemp,70  ,20 ,0 ,VERDE,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, 515  ,nroLineaTemp,40  ,20 ,0 ,VERDE,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, 555  ,nroLineaTemp,30  ,20 ,0 ,VERDE,NEGRO ,1 ,0)
	
	Call GF_setFontColor(BLANCO)
	Call GF_setFont(oPDF,"ARIAL", 8,FONT_STYLE_BOLD)
	Call GF_writeTextAlign(oPDF,10  ,103+(gNroLinea*10) , GF_TRADUCIR("Detalle")	, 195,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,385 ,99+(gNroLinea*10) , GF_TRADUCIR("Total")	, 60,PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(oPDF,385 ,108+(gNroLinea*10), GF_TRADUCIR("Actual")	, 60,PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(oPDF,445 ,99+(gNroLinea*10) , GF_TRADUCIR("Budget")	, 70,PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(oPDF,445 ,108+(gNroLinea*10), UCASE(getSimboloMonedaLetras(gMoneda))	, 70,PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(oPDF,515 ,99+(gNroLinea*10) , GF_TRADUCIR("Desv.")	, 40,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,515 ,108+(gNroLinea*10), UCASE(getSimboloMonedaLetras(gMoneda))	, 40,PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(oPDF,555 ,99+(gNroLinea*10) , GF_TRADUCIR("Desv.")	, 30,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,555 ,108+(gNroLinea*10), "%"	, 30,PDF_ALIGN_CENTER)
	
	
	for i = 0 to 2
		Call GF_writeTextAlign(oPDF,205+(i*60) ,103+(gNroLinea*10), cstr( V_MESES(pTrim,i,0) )	, 60,PDF_ALIGN_CENTER)
	next
	
	gNroLinea = gNroLinea + 2
	Call GF_setFontColor(NEGRO)
	Call GF_setFont(oPDF,"ARIAL", 10,FONT_STYLE_NORMAL)
	
End Function
'--------------------------------------------------------------------------------------------------
Function obtenerBudgetProporcionalResumen(byref pRs,pFecha,pDetalle)
	Dim rtrn,aux,trimFecha,fechaInicioTrim,difDias,auxImportes,detalleActual,cantDias
	
	'parte entera ( (mes fecha-1) / 3 )
	trimFecha = int(((mid( pFecha,5,2))-1) /3)
	
	if (not pRs.EoF) then
		detalleActual = pDetalle
		
		campoMoneda = "DLBUDGET"
		if (gMoneda = MONEDA_PESO) then	campoMoneda = "PSBUDGET"
	end if
	
	while (corteControlResumen(pRs,detalleActual))
		if (gBgtParcial) then			
			'Hay que sacar proporcional
			if (cdbl(trimFecha) = cdbl(pRs("PERIODO"))) then
				'Hay que sacar proporcional y es el trimestre limite
				
				'obtengo la cantidad de dias que tiene el trimestre
				fechaInicioTrim = left(pFecha,4) & V_INICIO_TRIM(trimFecha)
				cantDias = GF_DTEDIFF( fechaInicioTrim ,left(pFecha,4) & V_INICIO_TRIM(trimFecha+1),"D")				
				
				difDias = GF_DTEDIFF(fechaInicioTrim,pFecha,"D") + 1 ' +1 porque la diferencia de dias no es inclusiva				
				temp = rtrn
				rtrn = rtrn + int( (difDias * cdbl(pRs(campoMoneda)))/cantDias)								
			else			
				if (cdbl(trimFecha) > cdbl(pRs("PERIODO"))) then					
					'No es el trimestre limite, sumo el importe completo
					rtrn = rtrn + cdbl(pRs(campoMoneda))
				end if
			end if
		else
			'No es el trimestre limite, sumo el importe completo
			rtrn = rtrn + cdbl(pRs(campoMoneda))
		end if
		
		pRs.MoveNext
	wend
	'si se pide mostrar trimestres superiores a la fecha estimada del budget
	'los resultados daran negativos
	if (rtrn < 0) then rtrn = 0
	
	obtenerBudgetProporcionalResumen = rtrn
	
End Function
'--------------------------------------------------------------------------------------------------
Function corteControlResumen(rs,pDetalle)
	Dim rtrn 
	rtrn = not rs.eof
	
	if (rtrn) then rtrn = (CInt(pDetalle) = CInt(rs("IDDETALLE")))
	
	corteControlResumen = rtrn
End Function
'--------------------------------------------------------------------------------------------------
Function obtenerBudgetProporcionalDetalle(pFecha,pArea,pDetalle,pRs)
	Dim rtrn,aux,trimFecha,fechaInicioTrim,difDias,auxImportes,areaActual,detalleActual
	Dim esTrimLimite
	
	'parte entera ( (mes fecha-1) / 3 )
	trimFecha = int(((mid( pFecha,5,2))-1) /3)
	
	esTrimLimite = false
	if (not pRs.EoF) then
		areaActual = pArea
		detalleActual = pDetalle

		campoMoneda = "DLBUDGET"
		if (gMoneda = MONEDA_PESO) then	campoMoneda = "PSBUDGET"
		
		if (cdbl(trimFecha) = cdbl(pRs("PERIODO"))) then esTrimLimite = true
	end if
	
	rtrn = 0
	while (corteControlDetalle(pRs,areaActual,detalleActual))
		
		if (gBgtParcial) then
			'Hay que sacar proporcional
			if (esTrimLimite) then
				'Hay que sacar proporcional y es el trimestre limite
				
				'obtengo la cantidad de dias que tiene el trimestre
				cantDias = GF_DTEDIFF( left(pFecha,4) & V_INICIO_TRIM(trimFecha) ,left(pFecha,4) & V_INICIO_TRIM(trimFecha+1),"D")
				
				fechaInicioTrim = left(pFecha,4) & V_INICIO_TRIM(trimFecha)
				difDias = GF_DTEDIFF(fechaInicioTrim,pFecha,"D") + 1 ' +1 porque la diferencia de dias no es inclusiva
				rtrn = rtrn + int( (difDias * cdbl(pRs(campoMoneda)))/cantDias )
			else
				'No es el trimestre limite
				if (cdbl(trimFecha) > cdbl(pRs("PERIODO"))) then
					'Sumo solo si el trimestre no es superior a la fecha estimada del budget
					rtrn = rtrn + cdbl(pRs(campoMoneda))
				end if
			end if
		else
			'No es el trimestre limite, sumo el importe completo
			rtrn = rtrn + cdbl(pRs(campoMoneda))
		end if
		
		pRs.MoveNext
	wend
	
	'si se pide mostrar trimestres superiores a la fecha estimada del budget
	'los resultados daran negativos
	if (rtrn < 0) then rtrn = 0
	
	obtenerBudgetProporcionalDetalle = rtrn
	
End Function
'--------------------------------------------------------------------------------------------------
Function corteControlDetalle(rs,pArea,pDetalle)
	Dim rtrn 
	rtrn = not rs.eof
	
	if (rtrn) then rtrn = ((CInt(pArea) = CInt(rs("IDAREA"))) and (CInt(pDetalle) = CInt(rs("IDDETALLE"))))
	
	corteControlDetalle = rtrn
End Function

'**************************************************************************************************
'*                                                                                                *
'*                                   INICIO DE PAGINA                                             *
'*                                                                                                *
'**************************************************************************************************

Dim gIdObra,gFechaHasta, gFechaBudget,gMoneda,gTrimestre,gMostrarResumen,gMostrar1erTrim,gMostrar2doTrim
Dim gMostrar3erTrim,gMostrar4toTrim,gHoy,gNroHojas,gNroLinea,gRsObra,gRsDeta,gRsResu
Dim gTipoCambio,primeraHoja,gCantComentarios, gVecComentarios(),gBgtParcial
Dim obraCD, obraDS, obraDivID, obraDivDS, obraImorte, obraFechaBudget, obraMonedaID, gObraFechaInicio, gObraFechaFin, obraFechaAjustada, obraRespCD, obraRespDS
Dim gChkPIC, gChkVales, gChkFacturacion,gChkContable

primeraHoja = TRUE
Call GP_ConfigurarMomentos()

gIdObra 		= GF_Parametros7("idObra", 0, 6)	
gFechaHasta     = GF_Parametros7("hasta", "", 6)
gMoneda  	    = GF_Parametros7("moneda", "", 6)
gTrimestre  	= GF_Parametros7("trimestre", "", 6)
gBgtParcial  	= GF_Parametros7("bgtParcial", "", 6)
gChkFacturacion = false
if (GF_Parametros7("chkFacturacion", "", 6) <> "") then	gChkFacturacion = true
gChkPIC = false
if (GF_Parametros7("chkPIC", "", 6) <> "") then	gChkPIC = true
gChkVales = false
if (GF_Parametros7("chkVales", "", 6) <> "") then gChkVales = true
gChkContable = GF_Parametros7("chkContable", "", 6)

Call cargaDatos()

filename = "test.pdf"
Set oPDF = GF_createPDF(Server.MapPath("temp\" & filename))
call GF_setPDFMode(PDF_STREAM_MODE)

gTipoCambio = getTipoCambioBudget(gIdObra)

Call dibujarReporte()

Call GF_closePDF(oPDF)	
%>