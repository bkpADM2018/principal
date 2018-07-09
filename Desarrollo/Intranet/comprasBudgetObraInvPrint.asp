<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosPDF.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->

<%
CONST ANCHO_COL = 50

Const COL_AREA       = 690
CONST COL_DETALLE    = 30

CONST ENERO      = 0
CONST FEBRERO    = 1
CONST MARZO      = 2
CONST ABRIL      = 3
CONST MAYO       = 4
CONST JUNIO      = 5
CONST JULIO      = 6
CONST AGOSTO     = 7
CONST SEPTIEMBRE = 8
CONST OCTUBRE    = 9
CONST NOVIEMBRE  = 10
CONST DICIEMBRE  = 11
CONST TOTAL      = 12

CONST AREA_BUDGET = 12
CONST DETALLE_BUDGET = 13
CONST DSBUDGET = 14

Const SECCION_CTOS_SIN_ASIGNAR = 9999

Dim COL_ENERO, COL_FEBRERO, COL_MARZO, COL_ABRIL, COL_MAYO, COL_JUNIO, COL_JULIO, COL_AGOSTO, COL_SEPTIEMBRE, COL_OCTUBRE, COL_NOVIEMBRE, COL_DICIEMBRE,COL_TOTAL_PARCIAL

COL_ENERO      = COL_AREA - (ANCHO_COL*1)
COL_FEBRERO    = COL_AREA - (ANCHO_COL*2)
COL_MARZO      = COL_AREA - (ANCHO_COL*3)
COL_ABRIL      = COL_AREA - (ANCHO_COL*4)
COL_MAYO       = COL_AREA - (ANCHO_COL*5)
COL_JUNIO      = COL_AREA - (ANCHO_COL*6)
COL_JULIO      = COL_AREA - (ANCHO_COL*7)
COL_AGOSTO     = COL_AREA - (ANCHO_COL*8)
COL_SEPTIEMBRE = COL_AREA - (ANCHO_COL*9)
COL_OCTUBRE    = COL_AREA - (ANCHO_COL*10)
COL_NOVIEMBRE  = COL_AREA - (ANCHO_COL*11)
COL_DICIEMBRE  = COL_AREA - (ANCHO_COL*12)

COL_TOTAL_PARCIAL = COL_AREA - ((ANCHO_COL)*13)

Const VERDE  = "#396E8F"
Const VERDE2 = "#000000"
Const ROJO   = "#FFAA99"
Const NEGRO  = "#000000"
Const GRIS   = "#ADAFA7"
Const BLANCO = "#FFFFFF"

'----------------------------------------------------------------------------------------
Function obtenerImporteBudget(pArea,pDetalle)
	Dim strSQL,rs,conn,rtrn,auxWhere 
		
	auxWhere = ""
	Call mkWhere(auxWhere, "idobra", idObra,"=", 1)
	if (pArea <> 0) then Call mkWhere(auxWhere, "idarea", pArea,"=", 1)
	if (pDetalle <> 0) then Call mkWhere(auxWhere, "iddetalle", pDetalle,"=", 1)
	
	if (gMoneda = MONEDA_DOLAR) then
		strSQL = "Select Sum(dlbudget) suma from TBLBUDGETOBRAS " & auxWhere
	else
		strSQL = "Select Sum(PSBUDGET) suma from TBLBUDGETOBRAS " & auxWhere
	end if	
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	
	rtrn = cdbl(rs("suma"))
	
	
	obtenerImporteBudget = rtrn
	
End Function
'----------------------------------------------------------------------------------------
Function TitulosResumen()
	Call GF_squareBox(oPDF, 10 ,lineaActual ,330 ,10,0 ,GRIS,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, 340,lineaActual ,70  ,10,0 ,GRIS,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, 410,lineaActual ,70  ,10,0 ,GRIS,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, 480,lineaActual ,50  ,10,0 ,GRIS,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, 530,lineaActual ,50  ,10,0 ,GRIS,NEGRO ,1 ,0)
	
	'titulos
	Call GF_setFontColor(VERDE2)
	Call GF_setFont(oPDF,"ARIAL", 8 , FONT_STYLE_BOLD)
	Call GF_writeTextAlign(oPDF,10 , lineaActual+1, GF_TRADUCIR("Detalle") , 150,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,340, lineaActual+1, GF_TRADUCIR("Total")   , 70 ,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,410, lineaActual+1, GF_TRADUCIR("Budget")  , 70 ,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,480, lineaActual+1, GF_TRADUCIR("Desv")    , 50 ,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,530, lineaActual+1, GF_TRADUCIR("Desv %")  , 50 ,PDF_ALIGN_CENTER)
	Call GF_setFontColor(NEGRO)
	lineaActual = lineaActual +12
End Function
'----------------------------------------------------------------------------------------
Function dibujarResumen()
	Dim auxTotal,auxbudget,auxDesv,parametros(), aux
	redim parametros(9)
	
	if (nroHojas = 1) then
						
		'parametros de filtrado
		parametros(0)= "Responsable" & "|" & obraRespDS & " ("&obraRespCD&")"	
		parametros(1)= "Division" & "|" &GF_TRADUCIR(obraDivDS) 
		parametros(2)= "Inicio" & "|" & GF_FN2DTE(obraFechaInicio)
		parametros(3)= "Fin" & "|" & GF_FN2DTE(obraFechaFin)
		parametros(4)= "Moneda" & "|" & "Dolares"
		if (gMoneda = MONEDA_PESO) then
			parametros(4)= "Moneda" & "|" & "Pesos"
		end if
		parametros(5)= "Tipo de cambio" & "|" & gTipoCambio
		parametros(6)= "Fecha consulta" & "|Consumos al " & GF_FN2DTE(fechaHasta) & ", Budget al " & GF_FN2DTE(gFechaColumnaBudget)
		aux="NO"
		if (gChkPIC) then aux="SI"
		parametros(7)= "Incluir Comprometido" & "|" & aux
		aux="NO"
		if (gChkFacturacion) then aux="SI"
		parametros(8)= "Incluir Facturas" & "|" & aux
		aux="NO"
		if (gChkVales) then aux="SI"
		parametros(9)= "Incluir Vales" & "|" &  aux
		aux="NO"

		lineaActual = dibujarFiltros(parametros)
	end if
	
	Call TitulosResumen()
	
	
	'budgets
	for y = 0 to gCantBudgets-1
		Call GF_setFont(oPDF,"COURIER", 7 , FONT_STYLE_NORMAL)
		Call GF_setFontColor(NEGRO)
		if (CStr(gValores(DETALLE_BUDGET, y, 0)) = "0") then
			
			Call GF_setFont(oPDF,"COURIER", 7 , FONT_STYLE_BOLD)
			Call GF_setFontColor(BLANCO)
			Call GF_squareBox(oPDF, 10 ,lineaActual ,570 ,10,0 ,VERDE,GRIS ,1 ,0)
			
			Call GF_writeTextAlign(oPDF,10, lineaActual+1,gValores(AREA_BUDGET,y,0) , 15 ,PDF_ALIGN_RIGHT)
			Call GF_writeTextAlign(oPDF,30, lineaActual+1,gValores(DSBUDGET, y, 0)  , 150 ,PDF_ALIGN_LEFT)
			
			'total
			auxTotal = obtenerTotalGeneral(gValores(AREA_BUDGET,y,0),0)
			if (auxTotal > 0) then Call GF_writeTextAlign(oPDF,345, lineaActual+1, GF_EDIT_DECIMALS(auxTotal,2)  , 60,PDF_ALIGN_RIGHT)
			
			'Budget
			if (gValores(AREA_BUDGET,y,0) = SECCION_CTOS_SIN_ASIGNAR) then
			    'Solo para la primera linea del resumen donde AREA=0 and DETALLE=0 (ctos sin asignacion)
			    auxBudget = 0
			else
			    auxBudget = obtenerProporcionalBudget(obtenerImporteBudget(gValores(AREA_BUDGET,y,0),0))
            end if			    
			Call GF_writeTextAlign(oPDF,415, lineaActual+1, GF_EDIT_DECIMALS(auxBudget ,2)  , 60,PDF_ALIGN_RIGHT)
			
			auxDesv  = auxBudget - auxTotal
			if (auxDesv > 0 ) then auxDesv = 0
			Call GF_writeTextAlign(oPDF,485, lineaActual+1, GF_EDIT_DECIMALS(auxDesv*-1,2)  , 40,PDF_ALIGN_RIGHT)
			
			if ((auxDesv= 0) or (auxBudget=0)) then
				desvPerc = 0
			else
				desvPerc = (auxDesv/auxBudget)*100
			end if
			Call GF_writeTextAlign(oPDF,535, lineaActual+1, round(desvPerc,2) & " %"  , 40,PDF_ALIGN_RIGHT)
			
			Call GF_setFontColor(NEGRO)
			lineaActual = lineaActual + 10
			
			
		else
			
			Call GF_squareBox(oPDF, 10 ,lineaActual ,330 ,10,0 ,BLANCO,GRIS ,1 ,0)
			Call GF_squareBox(oPDF, 340,lineaActual ,70  ,10,0 ,BLANCO,GRIS ,1 ,0)
			Call GF_squareBox(oPDF, 410,lineaActual ,70  ,10,0 ,BLANCO,GRIS ,1 ,0)
			Call GF_squareBox(oPDF, 480,lineaActual ,50  ,10,0 ,BLANCO,GRIS ,1 ,0)
			Call GF_squareBox(oPDF, 530,lineaActual ,50  ,10,0 ,BLANCO,GRIS ,1 ,0)

	
		
			Call GF_writeTextAlign(oPDF,20, lineaActual+1,  gValores(DETALLE_BUDGET, y, 0)  ,  15,PDF_ALIGN_RIGHT)
			Call GF_writeTextAlign(oPDF,40, lineaActual+1,  gValores(DSBUDGET, y, 0)  , 150 ,PDF_ALIGN_LEFT)
			
			'Total
			auxTotal = obtenerTotalGeneral(gValores(AREA_BUDGET,y,0),gValores(DETALLE_BUDGET,y,0))
			if (auxTotal > 0) then Call GF_writeTextAlign(oPDF,345, lineaActual+1, GF_EDIT_DECIMALS(auxTotal,2)  , 60,PDF_ALIGN_RIGHT)
			
			
			'Budget
			if (gValores(AREA_BUDGET,y,0) = SECCION_CTOS_SIN_ASIGNAR) then
			    'Solo para la primera linea del resumen donde AREA=0 and DETALLE=0 (ctos sin asignacion)
			    auxBudget = 0
			else			    
			    auxBudget = obtenerProporcionalBudget(obtenerImporteBudget(gValores(AREA_BUDGET,y,0),gValores(DETALLE_BUDGET,y,0)) )
            end if	
			
			Call GF_writeTextAlign(oPDF,415, lineaActual+1, GF_EDIT_DECIMALS(auxBudget ,2)  , 60,PDF_ALIGN_RIGHT)
	
			auxDesv  = auxBudget - auxTotal
			if (auxDesv > 0 ) then auxDesv = 0
			Call GF_writeTextAlign(oPDF,485, lineaActual+1, GF_EDIT_DECIMALS(auxDesv*-1,2)  , 40,PDF_ALIGN_RIGHT)
			
			if ((auxDesv= 0) or (auxBudget=0)) then
				desvPerc = 0
			else
				desvPerc = (auxDesv/auxBudget)*10000
			end if
			Call GF_writeTextAlign(oPDF,535, lineaActual+1, GF_EDIT_DECIMALS(desvPerc,2) & " %"  , 40,PDF_ALIGN_RIGHT)
			
			lineaActual = lineaActual + 10
		end if
		
		if ( lineaActual => 800 ) then
			nroHojas = nroHojas +1
			Call GF_newPage(oPDF)
			Call DibujarEncabezadoVertical()
			lineaActual = 90
			Call TitulosResumen()
		end if
	next
	
	Call GF_squareBox(oPDF, 10 ,lineaActual ,330 ,10,0 ,VERDE,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, 340,lineaActual ,70  ,10,0 ,VERDE,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, 410,lineaActual ,70  ,10,0 ,VERDE,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, 480,lineaActual ,50  ,10,0 ,VERDE,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, 530,lineaActual ,50  ,10,0 ,VERDE,NEGRO ,1 ,0)

	Call GF_setFont(oPDF,"COURIER", 7 , FONT_STYLE_BOLD)
	Call GF_setFontColor(BLANCO)
	Call GF_writeTextAlign(oPDF,15, lineaActual+1,GF_TRADUCIR("TOTAL") , 15 ,PDF_ALIGN_LEFT)

	'total
	auxTotal = obtenerTotal()
	if (auxTotal > 0) then Call GF_writeTextAlign(oPDF,345, lineaActual+1, GF_EDIT_DECIMALS(auxTotal,2)  , 60,PDF_ALIGN_RIGHT)

	'total budget	
	auxBudget = obtenerProporcionalBudget(obtenerImporteBudget(0,0))
	Call GF_writeTextAlign(oPDF,415, lineaActual+1, GF_EDIT_DECIMALS(auxBudget ,2)  , 60,PDF_ALIGN_RIGHT)

	auxDesv  = auxBudget - auxTotal
	if (auxDesv > 0 ) then auxDesv = 0
	Call GF_writeTextAlign(oPDF,485, lineaActual+1, GF_EDIT_DECIMALS(auxDesv*-1, 2)  , 40,PDF_ALIGN_RIGHT)
	
	if ((auxDesv= 0) or (auxBudget=0)) then
		desvPerc = 0
	else
		desvPerc = (auxDesv/auxBudget)*10000
	end if
	Call GF_writeTextAlign(oPDF,535, lineaActual+1, GF_EDIT_DECIMALS(desvPerc, 2) & " %"  , 40,PDF_ALIGN_RIGHT)
			
	Call GF_setFontColor(NEGRO)
End Function
'----------------------------------------------------------------------------------------
Function DibujarEncabezadoVertical()
	
	'recuadro
	Call GF_squareBox(oPDF,3,5,590 ,830,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_ROUND) 
	
	'logo
	Call GF_writeImage(oPDF, Server.MapPath("Images\ADMlogo2.jpg"),15, 15, 48, 48, 0)
	
	'Titulo
	Call GF_setFont(oPDF,"ARIAL", 16 , FONT_STYLE_BOLD)
	Call GF_writeTextAlign(oPDF,0, 15, obraDS & " (" & obraCD & ")" , 590,PDF_ALIGN_CENTER)
	Call GF_setFont(oPDF,"ARIAL", 14 , FONT_STYLE_BOLD)
	Call GF_writeTextAlign(oPDF,0, 35, GF_TRADUCIR("Resumen") , 590,PDF_ALIGN_CENTER)
	
	Call GF_horizontalLine(oPDF,5,70,585)
	
	Call GF_setFont(oPDF,"COURIER",8,FONT_STYLE_NORMAL)
	Call GF_writeTextAlign(oPDF,0, 8, GF_FN2DTE(session("MmtoSistema")) , 590,PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,0,15,session("Usuario"), 587 , PDF_ALIGN_RIGHT)
	
	'Numero de pagina
	Call GF_writeTextAlign(oPDF,0, 835, GF_TRADUCIR("Pagina") & " " & nroHojas , 590,PDF_ALIGN_RIGHT)
	
end Function 
'-----------------------------------------------------------------------------------------
Function obtenerTotal()
	Dim rtrn
	
	rtrn = 0
	for z = 0 to gCantAnios
		for y = 0 to gCantBudgets-1
			for x = ENERO to DICIEMBRE
				if (gValores(x,y,z) <> "") then rtrn = rtrn + cdbl(gValores(x,y,z))
			next
		next
	next
	
	obtenerTotal = rtrn
End Function
'-----------------------------------------------------------------------------------------
Function obtenerTotalGeneral(pArea,pDetalle)
	Dim rtrn
	
	rtrn = 0	
	for y = 0 to gCantBudgets-1	        
			if ( pDetalle <> 0 ) then
				if (cdbl(gValores(AREA_BUDGET,y,0)) = cdbl(pArea) AND CStr(gValores(DETALLE_BUDGET,y,0)) = CStr(pDetalle) ) then
					for meses = ENERO to DICIEMBRE
						for anios = 0 to gCantAnios
							if (gValores(meses,y,anios) <> "") then rtrn = rtrn + cdbl(gValores(meses,y,anios))
						next
					next
				end if
			else
				if (cdbl(gValores(AREA_BUDGET,y,0)) = cdbl(pArea)) then
					for meses = ENERO to DICIEMBRE
						for anios = 0 to gCantAnios
							if (gValores(meses,y,anios) <> "") then rtrn = rtrn + cdbl(gValores(meses,y,anios))							                                
						next
					next
				end if
			end if
	next

	obtenerTotalGeneral = rtrn
	
End Function
'-----------------------------------------------------------------------------------------
Function dibujarTotalesMensuales(pAnio)
	Dim myAnio,totalesMensuales()
	
	redim totalesMensuales(13)
	
	myAnio = pAnio - gAnioInicio
	Call GF_squareBox(oPDF, lineaActual,10 ,pdf_currentFontSize,830       ,0 ,GRIS,NEGRO, 1 ,0)
	Call dibujarLineasSeparacion(NEGRO)
	
	Call GF_setFontColor(VERDE2)
	Call GF_setFont(oPDF,"COURIER", 7 , FONT_STYLE_BOLD)
	Call GF_writeVerticalText(oPDF, lineaActual, 835, GF_TRADUCIR("TOTALES"), 140, PDF_ALIGN_CENTER)

	totalesMensuales(ENERO)      = obtenerTotalMes(ENERO     , myAnio)
	totalesMensuales(FEBRERO)    = obtenerTotalMes(FEBRERO   , myAnio)	
	totalesMensuales(MARZO)      = obtenerTotalMes(MARZO     , myAnio)
	totalesMensuales(ABRIL)      = obtenerTotalMes(ABRIL     , myAnio)
	totalesMensuales(MAYO)       = obtenerTotalMes(MAYO      , myAnio)
	totalesMensuales(JUNIO)      = obtenerTotalMes(JUNIO     , myAnio)
	totalesMensuales(JULIO)      = obtenerTotalMes(JULIO     , myAnio)
	totalesMensuales(AGOSTO)     = obtenerTotalMes(AGOSTO    , myAnio)
	totalesMensuales(SEPTIEMBRE) = obtenerTotalMes(SEPTIEMBRE, myAnio)
	totalesMensuales(OCTUBRE)    = obtenerTotalMes(OCTUBRE   , myAnio)
	totalesMensuales(NOVIEMBRE)  = obtenerTotalMes(NOVIEMBRE , myAnio)
	totalesMensuales(DICIEMBRE)  = obtenerTotalMes(DICIEMBRE , myAnio)
	
	for mes = ENERO to DICIEMBRE
		totalesMensuales(TOTAL)  = totalesMensuales(TOTAL) + totalesMensuales(mes)
	next
	

	if (totalesMensuales(ENERO)      <> 0) then Call GF_writeVerticalText(oPDF, lineaActual, COL_AREA-5       , GF_EDIT_DECIMALS(totalesMensuales(ENERO)     , 2) ,ANCHO_COL-10 ,PDF_ALIGN_RIGHT)
	if (totalesMensuales(FEBRERO)    <> 0) then Call GF_writeVerticalText(oPDF, lineaActual, COL_ENERO-5      , GF_EDIT_DECIMALS(totalesMensuales(FEBRERO)   , 2) ,ANCHO_COL-10 ,PDF_ALIGN_RIGHT)
	if (totalesMensuales(MARZO)      <> 0) then Call GF_writeVerticalText(oPDF, lineaActual, COL_FEBRERO-5    , GF_EDIT_DECIMALS(totalesMensuales(MARZO)     , 2) ,ANCHO_COL-10 ,PDF_ALIGN_RIGHT)
	if (totalesMensuales(ABRIL)      <> 0) then Call GF_writeVerticalText(oPDF, lineaActual, COL_MARZO-5      , GF_EDIT_DECIMALS(totalesMensuales(ABRIL)     , 2) ,ANCHO_COL-10 ,PDF_ALIGN_RIGHT)
	if (totalesMensuales(MAYO)       <> 0) then Call GF_writeVerticalText(oPDF, lineaActual, COL_ABRIL-5      , GF_EDIT_DECIMALS(totalesMensuales(MAYO)      , 2) ,ANCHO_COL-10 ,PDF_ALIGN_RIGHT)
	if (totalesMensuales(JUNIO)      <> 0) then Call GF_writeVerticalText(oPDF, lineaActual, COL_MAYO-5       , GF_EDIT_DECIMALS(totalesMensuales(JUNIO)     , 2) ,ANCHO_COL-10 ,PDF_ALIGN_RIGHT)
	if (totalesMensuales(JULIO)      <> 0) then Call GF_writeVerticalText(oPDF, lineaActual, COL_JUNIO-5      , GF_EDIT_DECIMALS(totalesMensuales(JULIO)     , 2) ,ANCHO_COL-10 ,PDF_ALIGN_RIGHT)
	if (totalesMensuales(AGOSTO)     <> 0) then Call GF_writeVerticalText(oPDF, lineaActual, COL_JULIO-5      , GF_EDIT_DECIMALS(totalesMensuales(AGOSTO)    , 2) ,ANCHO_COL-10 ,PDF_ALIGN_RIGHT)
	if (totalesMensuales(SEPTIEMBRE) <> 0) then Call GF_writeVerticalText(oPDF, lineaActual, COL_AGOSTO-5     , GF_EDIT_DECIMALS(totalesMensuales(SEPTIEMBRE), 2) ,ANCHO_COL-10 ,PDF_ALIGN_RIGHT)
	if (totalesMensuales(OCTUBRE)    <> 0) then Call GF_writeVerticalText(oPDF, lineaActual, COL_SEPTIEMBRE-5 , GF_EDIT_DECIMALS(totalesMensuales(OCTUBRE)   , 2) ,ANCHO_COL-10 ,PDF_ALIGN_RIGHT)
	if (totalesMensuales(NOVIEMBRE)  <> 0) then Call GF_writeVerticalText(oPDF, lineaActual, COL_OCTUBRE-5    , GF_EDIT_DECIMALS(totalesMensuales(NOVIEMBRE) , 2) ,ANCHO_COL-10 ,PDF_ALIGN_RIGHT)
	if (totalesMensuales(DICIEMBRE)  <> 0) then Call GF_writeVerticalText(oPDF, lineaActual, COL_NOVIEMBRE-5  , GF_EDIT_DECIMALS(totalesMensuales(DICIEMBRE) , 2) ,ANCHO_COL-10 ,PDF_ALIGN_RIGHT)
	if (totalesMensuales(TOTAL)      <> 0) then Call GF_writeVerticalText(oPDF, lineaActual, COL_DICIEMBRE-5  , GF_EDIT_DECIMALS(totalesMensuales(TOTAL)     , 2) ,ANCHO_COL+20 ,PDF_ALIGN_RIGHT)
	Call GF_setFont(oPDF,"COURIER", 7 , FONT_STYLE_NORMAL)
	Call GF_setFontColor(NEGRO)
	
	lineaActual = lineaActual + 12
	
End Function
'-----------------------------------------------------------------------------------------
Function obtenerTotalMes(pMes,pAnio)
	Dim rtrn

	rtrn = 0
	for y = 0 to gCantBudgets-1
		if (gValores(pMes,y,pAnio) <> "") then rtrn = rtrn + cdbl(gValores(pMes,y,pAnio))		
	next	
	obtenerTotalMes = rtrn
End Function
'-----------------------------------------------------------------------------------------
Function obtenerTotalAreaMensual(pMes,pArea,pAnio)
	Dim rtrn
	
	rtrn = 0
	for y = 0 to gCantBudgets-1
		if (CDbl(gValores(AREA_BUDGET, y, 0)) = CDbl(pArea)) then
			if (gValores(pMes,y,pAnio) <> "") then rtrn = rtrn + CDbl(gValores(pMes,y,pAnio))			
		end if
	next	
	obtenerTotalAreaMensual = rtrn
End Function
'-----------------------------------------------------------------------------------------
Function obtenerTotalAreaDetalle(pArea,pDetalle,pAnio)
	Dim rtrn
	
	rtrn = 0	
	for y = 0 to gCantBudgets-1
		if ( pDetalle <> 0 ) then
			if ( cdbl(gValores(AREA_BUDGET, y, 0)) = cdbl(pArea) AND CStr(gValores(DETALLE_BUDGET, y, 0)) = CStr(pDetalle) ) then
				for mes = ENERO to DICIEMBRE
					if (gValores(mes,y,pAnio) <> "") then
						rtrn = rtrn + gValores(mes,y,pAnio)
					end if
				next
			end if
		else
			if ( cdbl(gValores(AREA_BUDGET, y, 0)) = cdbl(pArea) ) then
				for mes = ENERO to DICIEMBRE
					if (gValores(mes,y,pAnio) <> "") then
						rtrn = rtrn + gValores(mes,y,pAnio)
					end if
				next
			end if
		end if
	next	
	
	obtenerTotalAreaDetalle = rtrn
	
End Function
'-----------------------------------------------------------------------------------------
Function completarEspacios(p_palabra,p_len)
	Dim rtrn
	rtrn = p_palabra
	for i = len(p_palabra) to p_len
		rtrn = rtrn & "."
	next
	completarEspacios = rtrn
End Function
'-----------------------------------------------------------------------------------------
Function dibujarFiltros(p_parametros)
'Funcion que dibuja los parametros de busqueda 
'recibe como parametro un vector con la siguiente estructura en cada posicion:
'	"nombreBusqueda,valorBusqueda"
'Devuelve la posicion donde se puede seguir escribiendo
	Dim aux,x_inicial,y_inicial,font_size,mySplitChar
	Dim max_len
	
	Call GF_squareBox(oPDF,10,75,570 ,97,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_ROUND) 
	
	my_font_size = 8
	x_inicial = 20
	y_inicial = 80
	mySplitChar = "|"
	
	max_len = 0
	for i = 0 to ubound(p_parametros)
		aux = split(p_parametros(i),mySplitChar)
		if (len(aux(0)) > max_len) then
			max_len = len(aux(0))
		end if
	next
	
	for i = 0 to ubound(p_parametros)
		aux = split(p_parametros(i),mySplitChar)
		Call GF_setFont(oPDF,"COURIER",my_font_size,0)
		Call GF_writeTextAlign(oPDF,x_inicial,y_inicial+(i*my_font_size), completarEspacios(aux(0),max_len) & ": " & aux(1) ,580, PDF_ALIGN_LEFT)
	next 
	
	dibujarFiltros = y_inicial+(i*my_font_size)+my_font_size
End Function
'-------------------------------------------------------------------------------
Function dibujarHoja()
	Dim parametros()
	redim parametros(1)
	
	if (nroHojas = 1) then
		Call DibujarEncabezadoVertical()
		Call dibujarResumen()
		Call AgregarHoja()
		Call PDFGirarHoja(90)
	else		
		'recuadro
		Call GF_squareBox(oPDF,3,5,575 ,840,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_ROUND) 
	
		'logo		
		Call GF_writeImage(oPDF, Server.MapPath("Images\ADMlogo2.jpg"),15, 840, 48, 48, 90)
		
		
		'Titulo
		Call GF_setFont(oPDF,"ARIAL", 16 , FONT_STYLE_BOLD)
		Call GF_writeVerticalText(oPDF,10, 840, obraDS & " (" & obraCD & ")", 840, PDF_ALIGN_CENTER)
		Call GF_setFont(oPDF,"ARIAL", 14 , FONT_STYLE_BOLD)
		Call GF_writeVerticalText(oPDF,35, 840, GF_TRADUCIR("Detalle consumo"), 840, PDF_ALIGN_CENTER)
	
		Call GF_verticalLine(oPDF, 70, 10, 830)
		
		Call GF_setFont(oPDF,"COURIER",8,FONT_STYLE_NORMAL)
	
		Call GF_writeVerticalText(oPDF,5, 840, GF_FN2DTE(session("MmtoSistema")), 830, PDF_ALIGN_RIGHT)
		Call GF_writeVerticalText(oPDF,15, 840, session("Usuario"), 830, PDF_ALIGN_RIGHT)
		
		Call GF_setFont(oPDF,"COURIER",10,FONT_STYLE_NORMAL)
		
		'Numero de pagina
		Call GF_writeVerticalText(oPDF,580, 840, GF_TRADUCIR("Pagina") & " " & nroHojas, 830, PDF_ALIGN_RIGHT)
	
		
	end if
End Function
'-------------------------------------------------------------------------------
Function dibujarTitulos(pAnio)

	Call GF_setFontColor(VERDE2)
	Call GF_setFont(oPDF,"ARIAL",8,FONT_STYLE_BOLD)
	
	Call GF_squareBox(oPDF, lineaActual,10 ,pdf_currentFontSize ,830       ,0 ,GRIS,NEGRO, 1 ,0)
	Call GF_writeVerticalText(oPDF, lineaActual, 830, pAnio , 830, PDF_ALIGN_CENTER)
	lineaActual = lineaActual +pdf_currentFontSize

	'Recuadros titulo
	Call GF_squareBox(oPDF, lineaActual-pdf_currentFontSize,COL_AREA       ,18 ,150       ,0 ,GRIS,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, lineaActual,COL_ENERO      ,10 ,ANCHO_COL ,0 ,GRIS,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, lineaActual,COL_FEBRERO    ,10 ,ANCHO_COL ,0 ,GRIS,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, lineaActual,COL_MARZO      ,10 ,ANCHO_COL ,0 ,GRIS,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, lineaActual,COL_ABRIL      ,10 ,ANCHO_COL ,0 ,GRIS,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, lineaActual,COL_MAYO       ,10 ,ANCHO_COL ,0 ,GRIS,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, lineaActual,COL_JUNIO      ,10 ,ANCHO_COL ,0 ,GRIS,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, lineaActual,COL_JULIO      ,10 ,ANCHO_COL ,0 ,GRIS,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, lineaActual,COL_AGOSTO     ,10 ,ANCHO_COL ,0 ,GRIS,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, lineaActual,COL_SEPTIEMBRE ,10 ,ANCHO_COL ,0 ,GRIS,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, lineaActual,COL_OCTUBRE    ,10 ,ANCHO_COL ,0 ,GRIS,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, lineaActual,COL_NOVIEMBRE  ,10 ,ANCHO_COL ,0 ,GRIS,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, lineaActual,COL_DICIEMBRE  ,10 ,ANCHO_COL ,0 ,GRIS,NEGRO ,1 ,0)
	Call GF_squareBox(oPDF, lineaActual,COL_TOTAL_PARCIAL-30,10,ANCHO_COL+30,0,GRIS,NEGRO ,1 ,0)
	
	'Titulos	
	Call GF_setFont(oPDF,"ARIAL",8,FONT_STYLE_BOLD)
	
	Call GF_writeVerticalText(oPDF, lineaActual-5, 835, GF_TRADUCIR("Detalle"), 140, PDF_ALIGN_CENTER)
	Call GF_writeVerticalText(oPDF, lineaActual+1, COL_AREA-5       , GF_TRADUCIR("Enero")      ,ANCHO_COL-10 ,PDF_ALIGN_CENTER)
	Call GF_writeVerticalText(oPDF, lineaActual+1, COL_ENERO-5      , GF_TRADUCIR("Febrero")    ,ANCHO_COL-10 ,PDF_ALIGN_CENTER)
	Call GF_writeVerticalText(oPDF, lineaActual+1, COL_FEBRERO-5    , GF_TRADUCIR("Marzo")      ,ANCHO_COL-10 ,PDF_ALIGN_CENTER)
	Call GF_writeVerticalText(oPDF, lineaActual+1, COL_MARZO-5      , GF_TRADUCIR("Abril")      ,ANCHO_COL-10 ,PDF_ALIGN_CENTER)
	Call GF_writeVerticalText(oPDF, lineaActual+1, COL_ABRIL-5      , GF_TRADUCIR("Mayo")       ,ANCHO_COL-10 ,PDF_ALIGN_CENTER)
	Call GF_writeVerticalText(oPDF, lineaActual+1, COL_MAYO-5       , GF_TRADUCIR("Junio")      ,ANCHO_COL-10 ,PDF_ALIGN_CENTER)
	Call GF_writeVerticalText(oPDF, lineaActual+1, COL_JUNIO-5      , GF_TRADUCIR("Julio")      ,ANCHO_COL-10 ,PDF_ALIGN_CENTER)
	Call GF_writeVerticalText(oPDF, lineaActual+1, COL_JULIO-5      , GF_TRADUCIR("Agosto")     ,ANCHO_COL-10 ,PDF_ALIGN_CENTER)
	Call GF_writeVerticalText(oPDF, lineaActual+1, COL_AGOSTO-5     , GF_TRADUCIR("Septiembre") ,ANCHO_COL-10 ,PDF_ALIGN_CENTER)
	Call GF_writeVerticalText(oPDF, lineaActual+1, COL_SEPTIEMBRE-5 , GF_TRADUCIR("Octubre")    ,ANCHO_COL-10 ,PDF_ALIGN_CENTER)
	Call GF_writeVerticalText(oPDF, lineaActual+1, COL_OCTUBRE-5    , GF_TRADUCIR("Noviembre")  ,ANCHO_COL-10 ,PDF_ALIGN_CENTER)
	Call GF_writeVerticalText(oPDF, lineaActual+1, COL_NOVIEMBRE-5  , GF_TRADUCIR("Diciembre")  ,ANCHO_COL-10 ,PDF_ALIGN_CENTER)
	Call GF_writeVerticalText(oPDF, lineaActual+1, COL_DICIEMBRE-5  , GF_TRADUCIR("Total") & " " & pAnio      ,ANCHO_COL+20 ,PDF_ALIGN_CENTER)
	Call GF_setFont(oPDF,"ARIAL",8,FONT_STYLE_NORMAL)
	GF_setFontColor(NEGRO)
	
	
End Function
'-------------------------------------------------------------------------------
Function dibujarLineasSeparacion(pColor)
	GF_setFontColor(pColor)
	Call GF_horizontalLine(oPDF, lineaActual, COL_AREA, pdf_currentFontSize)
	Call GF_horizontalLine(oPDF, lineaActual, COL_ENERO, pdf_currentFontSize)
	Call GF_horizontalLine(oPDF, lineaActual, COL_FEBRERO, pdf_currentFontSize)
	Call GF_horizontalLine(oPDF, lineaActual, COL_MARZO, pdf_currentFontSize)
	Call GF_horizontalLine(oPDF, lineaActual, COL_ABRIL, pdf_currentFontSize)
	Call GF_horizontalLine(oPDF, lineaActual, COL_MAYO, pdf_currentFontSize)
	Call GF_horizontalLine(oPDF, lineaActual, COL_JUNIO, pdf_currentFontSize)
	Call GF_horizontalLine(oPDF, lineaActual, COL_JULIO, pdf_currentFontSize)
	Call GF_horizontalLine(oPDF, lineaActual, COL_AGOSTO, pdf_currentFontSize)
	Call GF_horizontalLine(oPDF, lineaActual, COL_SEPTIEMBRE, pdf_currentFontSize)
	Call GF_horizontalLine(oPDF, lineaActual, COL_OCTUBRE, pdf_currentFontSize)
	Call GF_horizontalLine(oPDF, lineaActual, COL_NOVIEMBRE, pdf_currentFontSize)
	Call GF_horizontalLine(oPDF, lineaActual, COL_DICIEMBRE, pdf_currentFontSize)
	GF_setFontColor(NEGRO)
End Function
'-------------------------------------------------------------------------------
Function DibujarCuerpo()
	Dim cantLineas,limitePaginacion,myFontSize,auxTotales(), anioActual
	cantLineas = 1
	Redim auxTotales(13)
	
	lineaActual = 75
	limitePaginacion = 66
	
	myFontSize = 7
		
	
	Call dibujarTitulos(gAnioInicio)
	
	lineaActual = lineaActual +11
	
	for z = 0 to gCantAnios-1
		anioActual = gAnioInicio + z
		'para que no cree paginas de años posteriores a la fechaHasta
		if ( CDbl(left(fechaHasta,4)) < anioActual) then exit function
		
		if (z > 0) then	
			'Se imprimen los totales del año anterior.
			Call dibujarTotalesMensuales(anioActual-1)
			'Se agregan las estructuras del nuevo año.
			Call AgregarHoja()
			Call PDFGirarHoja(90)
			Call dibujarTitulos(anioActual)
			lineaActual = lineaActual +10
			cantLineas = 0
		end if
			
		Call GF_setFontColor(NEGRO)
		
		for y = 0 to gCantBudgets-1
			
			Call GF_setFont(oPDF,"COURIER",myFontSize,FONT_STYLE_NORMAL)
			
			
			if (CStr(gValores(DETALLE_BUDGET, y, 0)) = "0") then
				Call GF_setFontColor(BLANCO)
				Call GF_setFont(oPDF,"COURIER",myFontSize,FONT_STYLE_BOLD)
				Call GF_squareBox(oPDF, lineaActual,10 ,pdf_currentFontSize ,830       ,0 ,VERDE,VERDE, 1 ,0)
				Call GF_writeVerticalText(oPDF, lineaActual, 840  , gValores(AREA_BUDGET, y, 0)     ,12 ,PDF_ALIGN_RIGHT)
				
				if (len(gValores(DSBUDGET, y, 0)) > 27) then
					Call GF_writeVerticalText(oPDF, lineaActual, 820  , left(gValores(DSBUDGET, y, 0),27) & "..."      ,140 ,PDF_ALIGN_LEFT)
				else
					Call GF_writeVerticalText(oPDF, lineaActual, 820  , gValores(DSBUDGET, y, 0)      ,140 ,PDF_ALIGN_LEFT)
				end if
				
				
				auxTotales(ENERO)      = obtenerTotalAreaMensual(ENERO		,gValores(AREA_BUDGET, y, 0), z)
				auxTotales(FEBRERO)    = obtenerTotalAreaMensual(FEBRERO	,gValores(AREA_BUDGET, y, 0), z)
				auxTotales(MARZO)      = obtenerTotalAreaMensual(MARZO		,gValores(AREA_BUDGET, y, 0), z)  
				auxTotales(ABRIL)      = obtenerTotalAreaMensual(ABRIL		,gValores(AREA_BUDGET, y, 0), z)
				auxTotales(MAYO)       = obtenerTotalAreaMensual(MAYO		,gValores(AREA_BUDGET, y, 0), z) 
				auxTotales(JUNIO)      = obtenerTotalAreaMensual(JUNIO		,gValores(AREA_BUDGET, y, 0), z) 
				auxTotales(JULIO)      = obtenerTotalAreaMensual(JULIO		,gValores(AREA_BUDGET, y, 0), z) 
				auxTotales(AGOSTO)     = obtenerTotalAreaMensual(AGOSTO		,gValores(AREA_BUDGET, y, 0), z)
				auxTotales(SEPTIEMBRE) = obtenerTotalAreaMensual(SEPTIEMBRE	,gValores(AREA_BUDGET, y, 0), z)
				auxTotales(OCTUBRE)    = obtenerTotalAreaMensual(OCTUBRE	,gValores(AREA_BUDGET, y, 0), z) 
				auxTotales(NOVIEMBRE)  = obtenerTotalAreaMensual(NOVIEMBRE	,gValores(AREA_BUDGET, y, 0), z) 
				auxTotales(DICIEMBRE)  = obtenerTotalAreaMensual(DICIEMBRE	,gValores(AREA_BUDGET, y, 0), z)				
								
				auxTotales(TOTAL) = 0
				for i = ENERO to DICIEMBRE
					auxTotales(TOTAL) = auxTotales(TOTAL) + auxTotales(i)
				next
				
				if ( auxTotales(ENERO)		> 0 ) then	Call GF_writeVerticalText(oPDF, lineaActual, COL_AREA-5       , GF_EDIT_DECIMALS( auxTotales(ENERO)     , 2) ,ANCHO_COL-10 ,PDF_ALIGN_RIGHT)
				if ( auxTotales(FEBRERO)	> 0 ) then  Call GF_writeVerticalText(oPDF, lineaActual, COL_ENERO-5      , GF_EDIT_DECIMALS( auxTotales(FEBRERO)   , 2) ,ANCHO_COL-10 ,PDF_ALIGN_RIGHT)
				if ( auxTotales(MARZO)		> 0 ) then  Call GF_writeVerticalText(oPDF, lineaActual, COL_FEBRERO-5    , GF_EDIT_DECIMALS( auxTotales(MARZO)     , 2) ,ANCHO_COL-10 ,PDF_ALIGN_RIGHT)
				if ( auxTotales(ABRIL)		> 0 ) then  Call GF_writeVerticalText(oPDF, lineaActual, COL_MARZO-5      , GF_EDIT_DECIMALS( auxTotales(ABRIL)     , 2) ,ANCHO_COL-10 ,PDF_ALIGN_RIGHT)
				if ( auxTotales(MAYO)		> 0 ) then  Call GF_writeVerticalText(oPDF, lineaActual, COL_ABRIL-5      , GF_EDIT_DECIMALS( auxTotales(MAYO)      , 2) ,ANCHO_COL-10 ,PDF_ALIGN_RIGHT)
				if ( auxTotales(JUNIO)		> 0 ) then  Call GF_writeVerticalText(oPDF, lineaActual, COL_MAYO-5       , GF_EDIT_DECIMALS( auxTotales(JUNIO)     , 2) ,ANCHO_COL-10 ,PDF_ALIGN_RIGHT)
				if ( auxTotales(JULIO)		> 0 ) then  Call GF_writeVerticalText(oPDF, lineaActual, COL_JUNIO-5      , GF_EDIT_DECIMALS( auxTotales(JULIO)     , 2) ,ANCHO_COL-10 ,PDF_ALIGN_RIGHT)
				if ( auxTotales(AGOSTO)		> 0 ) then  Call GF_writeVerticalText(oPDF, lineaActual, COL_JULIO-5      , GF_EDIT_DECIMALS( auxTotales(AGOSTO)    , 2) ,ANCHO_COL-10 ,PDF_ALIGN_RIGHT)
				if ( auxTotales(SEPTIEMBRE) > 0 ) then	Call GF_writeVerticalText(oPDF, lineaActual, COL_AGOSTO-5     , GF_EDIT_DECIMALS( auxTotales(SEPTIEMBRE), 2) ,ANCHO_COL-10 ,PDF_ALIGN_RIGHT)
				if ( auxTotales(OCTUBRE)	> 0 ) then  Call GF_writeVerticalText(oPDF, lineaActual, COL_SEPTIEMBRE-5 , GF_EDIT_DECIMALS( auxTotales(OCTUBRE)   , 2) ,ANCHO_COL-10 ,PDF_ALIGN_RIGHT)
				if ( auxTotales(NOVIEMBRE)	> 0 ) then  Call GF_writeVerticalText(oPDF, lineaActual, COL_OCTUBRE-5    , GF_EDIT_DECIMALS( auxTotales(NOVIEMBRE) , 2) ,ANCHO_COL-10 ,PDF_ALIGN_RIGHT)
				if ( auxTotales(DICIEMBRE)	> 0 ) then  Call GF_writeVerticalText(oPDF, lineaActual, COL_NOVIEMBRE-5  , GF_EDIT_DECIMALS( auxTotales(DICIEMBRE) , 2) ,ANCHO_COL-10 ,PDF_ALIGN_RIGHT)
				if ( auxTotales(TOTAL)		> 0 ) then  Call GF_writeVerticalText(oPDF, lineaActual, COL_DICIEMBRE-5  , GF_EDIT_DECIMALS( auxTotales(TOTAL)     , 2) ,ANCHO_COL+20 ,PDF_ALIGN_RIGHT)
				
				
				Call GF_setFont(oPDF,"COURIER",myFontSize,FONT_STYLE_NORMAL)
				Call GF_setFontColor(NEGRO)
			else
				Call GF_squareBox(oPDF, lineaActual,10 ,pdf_currentFontSize ,830       ,0 ,BLANCO,GRIS ,1 ,0)
				
				Call GF_writeVerticalText(oPDF, lineaActual, 835  , gValores(DETALLE_BUDGET, y, 0)  ,15 ,PDF_ALIGN_RIGHT)
				
				if (len(gValores(DSBUDGET, y, 0)) > 25) then
					Call GF_writeVerticalText(oPDF, lineaActual, 815  , left(gValores(DSBUDGET, y, 0),25) & "..."      ,140 ,PDF_ALIGN_LEFT)
				else
					Call GF_writeVerticalText(oPDF, lineaActual, 815  , gValores(DSBUDGET, y, 0)      ,140 ,PDF_ALIGN_LEFT)
				end if

				if (gValores(ENERO     ,y ,z)<> "" ) then Call GF_writeVerticalText(oPDF, lineaActual, COL_AREA-5       , GF_EDIT_DECIMALS(gValores(ENERO      ,y ,z), 2) ,ANCHO_COL-10 ,PDF_ALIGN_RIGHT)
				if (gValores(FEBRERO   ,y ,z)<> "" ) then Call GF_writeVerticalText(oPDF, lineaActual, COL_ENERO-5      , GF_EDIT_DECIMALS(gValores(FEBRERO    ,y ,z), 2) ,ANCHO_COL-10 ,PDF_ALIGN_RIGHT)
				if (gValores(MARZO     ,y ,z)<> "" ) then Call GF_writeVerticalText(oPDF, lineaActual, COL_FEBRERO-5    , GF_EDIT_DECIMALS(gValores(MARZO      ,y ,z), 2) ,ANCHO_COL-10 ,PDF_ALIGN_RIGHT)
				if (gValores(ABRIL     ,y ,z)<> "" ) then Call GF_writeVerticalText(oPDF, lineaActual, COL_MARZO-5      , GF_EDIT_DECIMALS(gValores(ABRIL      ,y ,z), 2) ,ANCHO_COL-10 ,PDF_ALIGN_RIGHT)
				if (gValores(MAYO      ,y ,z)<> "" ) then Call GF_writeVerticalText(oPDF, lineaActual, COL_ABRIL-5      , GF_EDIT_DECIMALS(gValores(MAYO       ,y ,z), 2) ,ANCHO_COL-10 ,PDF_ALIGN_RIGHT)
				if (gValores(JUNIO     ,y ,z)<> "" ) then Call GF_writeVerticalText(oPDF, lineaActual, COL_MAYO-5       , GF_EDIT_DECIMALS(gValores(JUNIO      ,y ,z), 2) ,ANCHO_COL-10 ,PDF_ALIGN_RIGHT)
				if (gValores(JULIO     ,y ,z)<> "" ) then Call GF_writeVerticalText(oPDF, lineaActual, COL_JUNIO-5      , GF_EDIT_DECIMALS(gValores(JULIO      ,y ,z), 2) ,ANCHO_COL-10 ,PDF_ALIGN_RIGHT)
				if (gValores(AGOSTO    ,y ,z)<> "" ) then Call GF_writeVerticalText(oPDF, lineaActual, COL_JULIO-5      , GF_EDIT_DECIMALS(gValores(AGOSTO     ,y ,z), 2) ,ANCHO_COL-10 ,PDF_ALIGN_RIGHT)
				if (gValores(SEPTIEMBRE,y ,z)<> "" ) then Call GF_writeVerticalText(oPDF, lineaActual, COL_AGOSTO-5     , GF_EDIT_DECIMALS(gValores(SEPTIEMBRE ,y ,z), 2) ,ANCHO_COL-10 ,PDF_ALIGN_RIGHT)
				if (gValores(OCTUBRE   ,y ,z)<> "" ) then Call GF_writeVerticalText(oPDF, lineaActual, COL_SEPTIEMBRE-5 , GF_EDIT_DECIMALS(gValores(OCTUBRE    ,y ,z), 2) ,ANCHO_COL-10 ,PDF_ALIGN_RIGHT)
				if (gValores(NOVIEMBRE ,y ,z)<> "" ) then Call GF_writeVerticalText(oPDF, lineaActual, COL_OCTUBRE-5    , GF_EDIT_DECIMALS(gValores(NOVIEMBRE  ,y ,z), 2) ,ANCHO_COL-10 ,PDF_ALIGN_RIGHT)
				if (gValores(DICIEMBRE ,y ,z)<> "" ) then Call GF_writeVerticalText(oPDF, lineaActual, COL_NOVIEMBRE-5  , GF_EDIT_DECIMALS(gValores(DICIEMBRE  ,y, z), 2) ,ANCHO_COL-10 ,PDF_ALIGN_RIGHT)
				auxTotalDetalle =obtenerTotalAreaDetalle(gValores(AREA_BUDGET, y, 0),gValores(DETALLE_BUDGET, y, 0),z)
				if (auxTotalDetalle <> 0) then Call GF_writeVerticalText(oPDF, lineaActual, COL_DICIEMBRE-5  , GF_EDIT_DECIMALS(auxTotalDetalle, 2)      ,ANCHO_COL+20 ,PDF_ALIGN_RIGHT)
				
				Call dibujarLineasSeparacion(GRIS)
			end if
			
			
			lineaActual = lineaActual +pdf_currentFontSize
			cantLineas = cantLineas + 1
			
			if (cantLineas > limitePaginacion) then 
				Call AgregarHoja()
				Call PDFGirarHoja(90)
				Call dibujarTitulos(anioActual)
				lineaActual = lineaActual +11
				cantLineas = 0
			end if
		next
	next
	'Se imprimen los totales del último año
	Call dibujarTotalesMensuales(anioActual)

		
End Function
'-------------------------------------------------------------------------------
Function AgregarHoja()
	nroHojas = nroHojas +1
	Call GF_newPage(oPDF)
	Call dibujarHoja()
	lineaActual = 75
end Function
'-------------------------------------------------------------------------------
Function cargarDatos()
	Dim strSQL,con,rs,x,y,z,auxAnio, myFechaHasta,campoImporte,campoImporte2, myFechaAMD, campoImporte31, campoImporte32,campoImporte4
		
	'###########################################
	' Cargo la matriz con los budgets	
	'###########################################
	strSQL = "Select * from TBLBUDGETObras where idobra = " & idObra & " Order by IDAREA, IDDETALLE"
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	
	gCantBudgets = rs.RecordCount
		
	'La dimención Z (Anios) se inicializa en 1 pero luego se ira creciendo en forma dinamica según se necesite.
	redim gValores(15,gCantBudgets ,1)
	
	'gTipoCambio = rs("TIPOCAMBIO")
	gTipoCambio = getTipoCambio(MONEDA_DOLAR,"")
	y=0	
    'Todo el resto del budget.
	While not rs.EoF						
		gValores(AREA_BUDGET	,y,0)= CLng(rs("IDAREA"))
		gValores(DETALLE_BUDGET	,y,0)= CStr(rs("IDDETALLE"))
		gValores(DSBUDGET		,y,0)= CStr(rs("DSBUDGET"))							
		y=y+1
		rs.MoveNext
	Wend	
	

	'#############################################
	'Cargo todos los valores mensuales donde corresponda
	'#############################################
	strSQL = ""
	if ((gChkVales) or (gChkPIC) or (gChkFacturacion))then	
		campoImporte = "acd7.IMPORTEDOLARES"	
		campoImporte2= "det.vluDOLARES"
		campoImporte31= "((det.IMPORTEPESOS-det.IMPORTEPESOSFACTURADO)/" & gTipoCambio & ")"
		campoImporte32= "(det.IMPORTEDOLARES-det.IMPORTEDOLARESFACTURADO)"
		campoImporte41 = "((IMPORTEASIGNADO-IMPORTEGASTADO)/" & gTipoCambio & ")"
		campoImporte42= "(IMPORTEASIGNADO-IMPORTEGASTADO)"
		if (gMoneda=MONEDA_PESO) then 
			campoImporte = "acd7.IMPORTEPESOS"
			campoImporte2= "det.vlupesos"
			campoImporte31= "(det.IMPORTEPESOS-det.IMPORTEPESOSFACTURADO)"
			campoImporte32= "((det.IMPORTEDOLARES-det.IMPORTEDOLARESFACTURADO)*" & gTipoCambio & ")"
			campoImporte41= "(IMPORTEASIGNADO-IMPORTEGASTADO)"
		    campoImporte42= "(IMPORTEASIGNADO-IMPORTEGASTADO)*" & gTipoCambio & ")"
		end if
	
		myFechaHasta = GF_FN2DTCONTABLE(fechaHasta)
		myFechaAMD = fechaHasta
	
		strSQL = ""
		strSQL = strSQL & "SELECT   anio   , "
		strSQL = strSQL & "         area   , "
		strSQL = strSQL & "         detalle, "
		strSQL = strSQL & "         mes    , "
		strSQL = strSQL & "			case when sum(gasto) is null then 0 "
		strSQL = strSQL & "			else sum(gasto) end as gasto "		
		strSQL = strSQL & "FROM     ( "		
		' sql para las facturas que corresponden a la inversion.	
		if ((gChkPIC) or (gChkFacturacion)) then		
			strSQL = strSQL & "			SELECT  1 nro, acd7.anio         anio, "
			strSQL = strSQL & "                  acd7.IDAREA         Area   , "
			strSQL = strSQL & "                  acd7.IDDETALLE      detalle, "
			strSQL = strSQL & "                  acd7.mes            mes    , "			
			strSQL = strSQL & "                  SUM( "&campoImporte&")*100 gasto "			
			strSQL = strSQL & "         FROM     VWMEP001C acd7 "
			strSQL = strSQL & "                  INNER JOIN VWCOMPROBANTES acds ON acd7.NroInt = acds.NroInt AND ACD7.anio=acds.anio AND ACD7.mes=acds.mes "
			strSQL = strSQL & "					 INNER JOIN tblarticulos art on art.idarticulo=acd7.IDARTICULO "
			strSQL = strSQL & "					 INNER JOIN tblartcategorias cat on art.idcategoria=cat.idcategoria "	
			strSQL = strSQL & "         WHERE    acd7.IDOBRA = " & idObra
			strSQL = strSQL & "			AND		 acd7.IDARTICULO NOT IN (" & ITEM_FONDO_REPARO_ARS & "," & ITEM_FONDO_REPARO_USD & ", " & ITEM_FONDO_REPARO_ARS_IVA & "," & ITEM_FONDO_REPARO_USD_IVA & ")" 
			strSQL = strSQL & "			and		 cat.tipocategoria	  <> '" & TIPO_CAT_IMPUESTOS & "'"
			'Se pone una fecha desde inicial par que no tome documentos del milenio pasado
			strSQL = strSQL & "         AND      acds.feccbt <= '" & myFechaHasta & "' and acds.feccbt >= '2000-01-01'"
			strSQL = strSQL & "         GROUP BY acd7.anio, "
			strSQL = strSQL & "                  acd7.mes, "
			strSQL = strSQL & "                  acd7.IDAREA       , "
			strSQL = strSQL & "                  acd7.IDDETALLE "
			strSQL = strSQL & "          "		
		end if
		' sql para obtener los vales que correspondan a la inversion.
		if (gChkVales) then			
		    if ((gChkPIC) or (gChkFacturacion)) then strSQL = strSQL & " UNION "					
			strSQL = strSQL & "         SELECT   2 nro, SUBSTRING(convert(varchar, cab.fecha),3,2)         anio   , "
			strSQL = strSQL & "                  cab.idbudgetarea                 area   , "
			strSQL = strSQL & "                  cab.idbudgetdetalle              detalle, "
			strSQL = strSQL & "                  SUBSTRING(convert(varchar,cab.fecha),5,2)            mes    , "
			strSQL = strSQL & "                  SUM(det.existencia*"&campoImporte2&") gasto "
			strSQL = strSQL & "         FROM     tblvalescabecera cab "
			strSQL = strSQL & "                  INNER JOIN tblvalesdetalle det "
			strSQL = strSQL & "                  ON       cab.idvale = det.idvale "
			strSQL = strSQL & "         WHERE    cab.idobra          = " & idObra
			strSQL = strSQL & "         AND      cab.estado = "& ESTADO_ACTIVO
			strSQL = strSQL & "         AND      cab.fecha <= " & myFechaAMD
			strSQL = strSQL & "         GROUP BY SUBSTRING(convert(varchar,cab.fecha),3,2), "
			strSQL = strSQL & "                  SUBSTRING(convert(varchar,cab.fecha),5,2), "
			strSQL = strSQL & "                  cab.idbudgetarea     , "
			strSQL = strSQL & "                  cab.idbudgetdetalle "	
		end if
		'sub sql para los pic que correspondan a la inversion.
		if (gChkPIC) then
			strSQL = strSQL & " UNION "
			strSQL = strSQL & "         SELECT   31 nro, SUBSTRING(convert(varchar,cab.momento),3,2)	anio	, "
			strSQL = strSQL & "                  det.idarea                 area	, "
			strSQL = strSQL & "                  det.iddetalle              detalle , "
			strSQL = strSQL & "                  SUBSTRING(convert(varchar,cab.momento),5,2)	mes     , "
			strSQL = strSQL & "                  SUM("&campoImporte31&")		gasto "
			strSQL = strSQL & "         FROM     tblctzcabecera cab "
			strSQL = strSQL & "                  INNER JOIN tblctzdetalle det "
			strSQL = strSQL & "                  ON       cab.idcotizacion = det.idcotizacion "
			strSQL = strSQL & "         WHERE    cab.idobra                = " & idObra
			strSQL = strSQL & "					 AND cab.idcontrato=0 "
			strSQL = strSQL & "					 AND cab.estado <> '" & CTZ_ANULADA & "' "
			strSQL = strSQL & "                  AND cab.momento <= " & myFechaAMD & "595959"									
			strSQL = strSQL & "                  AND  cab.CDMONEDA = '" & MONEDA_PESO & "'"
			strSQL = strSQL & "         GROUP BY SUBSTRING(convert(varchar,cab.momento),3,2), "
			strSQL = strSQL & "                  SUBSTRING(convert(varchar,cab.momento),5,2), "	
			strSQL = strSQL & "					 det.idarea             , "
			strSQL = strSQL & "                  det.iddetalle            "
			strSQL = strSQL & " UNION "
			strSQL = strSQL & "         SELECT   32 nro, SUBSTRING(convert(varchar,cab.momento),3,2)	anio	, "
			strSQL = strSQL & "                  det.idarea                 area	, "
			strSQL = strSQL & "                  det.iddetalle              detalle , "
			strSQL = strSQL & "                  SUBSTRING(convert(varchar,cab.momento),5,2)	mes     , "
			strSQL = strSQL & "                  SUM("&campoImporte32&")	gasto "
			strSQL = strSQL & "         FROM     tblctzcabecera cab "
			strSQL = strSQL & "                  INNER JOIN tblctzdetalle det "
			strSQL = strSQL & "                  ON       cab.idcotizacion = det.idcotizacion "
			strSQL = strSQL & "         WHERE    cab.idobra                = " & idObra
			strSQL = strSQL & "					 AND cab.idcontrato=0 "
			strSQL = strSQL & "					 AND cab.estado <> '" & CTZ_ANULADA & "' "
			strSQL = strSQL & "                  AND cab.momento <= " & myFechaAMD & "595959"			
			strSQL = strSQL & "                  AND  cab.CDMONEDA = '" & MONEDA_DOLAR & "'"
			strSQL = strSQL & "         GROUP BY SUBSTRING(convert(varchar,cab.momento),3,2), "
			strSQL = strSQL & "                  SUBSTRING(convert(varchar,cab.momento),5,2), "	
			strSQL = strSQL & "					 det.idarea             , "
			strSQL = strSQL & "                  det.iddetalle            "				
			strSQL = strSQL & " UNION "
			'Se adicionan a lo comprometido los contratos con AREA-DETALLE.
			strSQL = strSQL & "         SELECT 41 nro, ANIO, AREA, DETALLE, MES, Sum(SALDOCTC) GASTO    "
	        strSQL = strSQL & "         FROM    "
	        strSQL = strSQL & "             (   SELECT			"
            strSQL = strSQL & "                  CASE WHEN SUBSTRING(convert(varchar,P.FECHAINICIO),1,8) >= " & obraFechaInicio & " THEN SUBSTRING(convert(varchar,P.FECHAINICIO),3,2) ELSE " & Mid(obraFechaInicio, 3, 2)  & " END	ANIO	, "
			strSQL = strSQL & "                  P.IDAREA                 AREA	, "
			strSQL = strSQL & "                  P.IDDETALLE              DETALLE , "
			strSQL = strSQL & "                  CASE WHEN SUBSTRING(convert(varchar,P.FECHAINICIO),1,8) >= " & obraFechaInicio & " THEN SUBSTRING(convert(varchar,P.FECHAINICIO),5,2) ELSE " & Mid(obraFechaInicio, 5, 2)  & " END 	MES     , "			
			strSQL = strSQL & "                  " & campoImporte41 & "	SALDOCTC "						
			strSQL = strSQL & "                 FROM TBLOBRACONTRATOS CTC "
			strSQL = strSQL & "                 INNER JOIN TBLCTCPARTIDAS P on P.IDCONTRATO=CTC.IDCONTRATO "
			strSQL = strSQL & "                 WHERE P.CDMONEDA='" & MONEDA_PESO & "' and CTC.ESTADO not in (" & ESTADO_CTC_CANCELADO & ") AND P.IDOBRA=" & idObra
			strSQL = strSQL & "                 AND P.IDAREA<>0 and P.IDDETALLE<>0 "			
			strSQL = strSQL & "             ) T"
			strSQL = strSQL & "         GROUP BY T.ANIO, "
			strSQL = strSQL & "                  T.MES, "	
			strSQL = strSQL & "					 T.AREA             , "
			strSQL = strSQL & "                  T.DETALLE            "
			strSQL = strSQL & " UNION "
			'Se adicionan a lo comprometido los contratos con AREA-DETALLE.
			strSQL = strSQL & "         SELECT 42 nro ,ANIO, AREA, DETALLE, MES, Sum(SALDOCTC) GASTO    "
	        strSQL = strSQL & "         FROM    "
	        strSQL = strSQL & "             (   SELECT			"
            strSQL = strSQL & "                  CASE WHEN SUBSTRING(convert(varchar,P.FECHAINICIO),1,8) >= " & obraFechaInicio & " THEN SUBSTRING(convert(varchar,P.FECHAINICIO),3,2) ELSE " & Mid(obraFechaInicio, 3, 2)  & " END	ANIO	, "
			strSQL = strSQL & "                  P.IDAREA                 AREA	, "
			strSQL = strSQL & "                  P.IDDETALLE              DETALLE , "
			strSQL = strSQL & "                  CASE WHEN SUBSTRING(convert(varchar,P.FECHAINICIO),1,8) >= " & obraFechaInicio & " THEN SUBSTRING(convert(varchar,P.FECHAINICIO),5,2) ELSE " & Mid(obraFechaInicio, 5, 2)  & " END 	MES     , "			
			strSQL = strSQL & "                  " & campoImporte42 & "	SALDOCTC "			
			strSQL = strSQL & "                 FROM TBLOBRACONTRATOS CTC "
            strSQL = strSQL & "                 INNER JOIN TBLCTCPARTIDAS P on P.IDCONTRATO=CTC.IDCONTRATO "
			strSQL = strSQL & "                 WHERE P.CDMONEDA='" & MONEDA_DOLAR & "' and CTC.ESTADO in (" & ESTADO_CTC_CANCELADO & ") AND P.IDOBRA=" & idObra
			strSQL = strSQL & "                 AND P.IDAREA<>0 and P.IDDETALLE<>0 "						
			strSQL = strSQL & "             ) T"
			strSQL = strSQL & "         GROUP BY T.ANIO, "
			strSQL = strSQL & "                  T.MES, "	
			strSQL = strSQL & "					 T.AREA             , "
			strSQL = strSQL & "                  T.DETALLE            "
		end if
		 
		strSQL = strSQL & "         ) "
		strSQL = strSQL & "         aux "
		strSQL = strSQL & "GROUP BY anio   , "
		strSQL = strSQL & "         area   , "
		strSQL = strSQL & "         detalle, "
		strSQL = strSQL & "         mes "
		strSQL = strSQL & "ORDER BY anio   , "
		strSQL = strSQL & "         area   , "
		strSQL = strSQL & "         detalle, "
		strSQL = strSQL & "         mes"	
	else
		'No se selecciono ningun tipo de consumo, se crea una SQL solo para evitar 
		strSQL = "Select * from TBLBUDGETOBRAS where 1=2"
	end if
	'Response.Write strSQL
	'response.end
	'    
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)

	gAnioInicio = left(obraFechaInicio, 4)	
	if (not rs.EoF) then gAnioInicio =  2000 + CInt(rs("ANIO")) 		
	
	while not rs.EoF			
		
		'Calculo el subindice de la dimensión anio en donde corresponde poner el dato.
		auxAnio = ( 2000 + CInt(rs("ANIO")) ) - gAnioInicio
		'Si el sub Indice no existe, lo creo.
		if (UBound(gValores, 3) < auxAnio) then
			'Se tiene un registro con un nuevo anio, se agrega la dimension.	
			Redim Preserve gValores(15,gCantBudgets ,auxAnio)
		end if
		call insertarValor(rs("AREA"),rs("DETALLE"),CInt(rs("MES")) - 1,auxAnio,rs("GASTO"))
		rs.MoveNext
	wend
	
	gCantAnios = UBound(gValores, 3)
End Function
'-------------------------------------------------------------------------------
Function insertarValor(pArea,pDetalle,pMes,pAnio,pValor)
	
	for i = 0 to gCantBudgets 
		if ( cdbl(gValores(AREA_BUDGET,i, 0)) = cdbl(pArea) AND CStr(gValores(DETALLE_BUDGET,i,0)) = CStr(pDetalle) ) then
			gValores(pMes,i,pAnio) = cdbl(pValor)
		end if
	next
	 
End function
'-------------------------------------------------------------------------------
Function obtenerProporcionalBudget(pBudget)
	Dim rtrn,importeDiario,aux
	
	if (gBgtParcial) then
		importeDiario = cdbl(pBudget) / gCantDiasObra	
		rtrn = pBudget	
		if (gDifDias < gCantDiasObra) then
			aux = cstr(importeDiario * gDifDias)
			aux = split(aux,".")
			rtrn = aux(0)
			rtrn = round(cdbl(rtrn),2)
		end if
	else
		rtrn = pBudget
	end if		
	obtenerProporcionalBudget = rtrn
	
End Function
'-------------------------------------------------------------------------------
Dim gValores(),gCantBudgets,gCantAnios,gTipoCambio,gDifDias,gCantDiasObra
Dim oPDF,fechaHasta,gMoneda,gHoy
Dim idObra, obraCD, obraDS, obraDivID, obraDivDS, obraImorte, obraFechaBudget, obraMonedaID, obraFechaInicio, obraFechaFin, obraFechaAjustada, obraRespCD, obraRespDS
Dim nroHojas,rsBudget,fileName,lineaActual, gFechaColumnaBudget, gBgtParcial, gAnioInicio
Dim gChkPIC, gChkVales, gChkFacturacion

	idObra = GF_Parametros7("idObra", 0, 6)	
	fechaHasta  = GF_Parametros7("hasta", "", 6)
	gMoneda  = GF_Parametros7("moneda", "", 6)	
	gBgtParcial = false
	if (GF_Parametros7("bgtParcial", "", 6) <> "") then	gBgtParcial = true
	gChkFacturacion = false
	if (GF_Parametros7("chkFacturacion", "", 6) <> "") then	gChkFacturacion = true
	gChkPIC = false
	if (GF_Parametros7("chkPIC", "", 6) <> "") then	gChkPIC = true
	gChkVales = false
	if (GF_Parametros7("chkVales", "", 6) <> "") then gChkVales = true

	
	if (gMoneda = "") then
		gMoneda = MONEDA_DOLAR
	end if

	Call GP_ConfigurarMomentos()
	Call loadDatosObra(IdObra, obraCD, obraDS, obraDivID, obraDivDS, obraImorte, obraFechaBudget, obraMonedaID, obraFechaInicio, obraFechaFin, obraFechaAjustada, obraRespCD, obraRespDS)

	fechaFinlizacion = obraFechaFin
	if (CDbl(obraFechaAjustada) <> 0) then fechaFinlizacion = obraFechaAjustada
	
	if (fechaHasta = "") then
		Call GP_ConfigurarMomentos()
		fechaHasta = Cdbl(left(session("MmtoSistema"),8))
	else
		fechaHasta = Cdbl(GF_DTE2FN(fechaHasta))
	end if
	
	gDifDias = GF_DTEDIFF(obraFechaInicio,fechaHasta,"D")
	gCantDiasObra = GF_DTEDIFF(obraFechaInicio,fechaFinlizacion,"D")

	gFechaColumnaBudget = fechaFinlizacion
	if (gBgtParcial) then gFechaColumnaBudget = fechaHasta
	
	filename = "test.pdf"
	Set oPDF = GF_createPDF(Server.MapPath("temp\" & filename))
	call GF_setPDFMode(PDF_STREAM_MODE)
		
	'Set rsBudget = leerBudget(idObra)
	
	nroHojas = 1

	Call cargarDatos()
	
	lineaActual = dibujarHoja()	
	
	Call dibujarCuerpo

	Call GF_closePDF(oPDF)
%>