<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosPDF.asp"-->
<!--#include file="Includes/procedimientosLog.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosMantenimiento.asp"-->
<%
dim oPDF, rs, conn, strSQL, nroPagina, RPT_accion, ultimaLinea
dim current_line, auxText, listOfPM, lote, loteArr

Const PAGINA_CABECERA  = 0
Const PAGINA_SIGUIENTE = 1

Const LINEA_POST_ENCABEZADO = 80
Const PAGE_HEIGHT_SIZE = 750
Const X_COL_CAB_1  = 40
Const X_COL_CAB_2  = 300
Const X_COL_TASK_1 = 30
Const X_COL_TASK_2 = 50
Const X_COL_TASK_3 = 540

Const X_COL_ITEM_NRO = 30
Const X_COL_ITEM_ARTICULO = 50
Const X_COL_ITEM_CANTIDAD = 430
Const X_COL_ITEM_PM = 530
Const X_COL_ITEM_ARTID = 50
Const X_COL_ITEM_ARTDS = 80
Const X_COL_ITEM_CANTPROG = 430
Const X_COL_ITEM_CANTREAL = 480
Const ANCHO_COL_ITEM_NRO = 20
Const ANCHO_COL_ITEM_ARTICULO = 380
Const ANCHO_COL_ITEM_CANTIDAD = 100
Const ANCHO_COL_ITEM_PM = 30
Const ANCHO_COL_ITEM_ARTID = 30
Const ANCHO_COL_ITEM_ARTDS = 350
Const ANCHO_COL_ITEM_CANTPROG = 50
Const ANCHO_COL_ITEM_CANTREAL = 50

Const ANCHO_COL_CAB	   = 50
Const ANCHO_COL_CAB_1  = 100
Const ANCHO_COL_CAB_2  = 100
Const ANCHO_TOTAL      = 550
Const ANCHO_COL_TASK_1 = 20
Const ANCHO_COL_TASK_2 = 490
Const ANCHO_COL_TASK_3 = 20

Const MAX_Y_PAGINA      = 760
CONST SEC_ITEMS = "ITEMS"
CONST SEC_TASKS = "TASKS"
CONST SEC_OBVS = "OBVS"
CONST SEC_EXE = "EXE"
Const LARGO_TITULOS = 13
Const INICIO_PAGINA = 100
Const MARGEN = 20
Const SEPARACION = 15

'**************************************************************************
'**************************** INICIO PAGINA *******************************
'**************************************************************************

    lote= GF_Parametros7("idOT", "", 6)    
    loteArr = Split(lote,",")

    momento = session("MmtoSistema")	
	Set oPDF = GF_createPDF("PDFTemp")
	Call GF_setPDFMODE(PDF_STREAM_MODE)
	
	Call armadoPDF(oPDF, loteArr)	

	Call GF_closePDF(oPDF)
'-----------------------------------------------------------------------------------------
Function armadoPDF(oPDF, loteArr)
    Dim i
    ',518,1041
    for i = LBound(loteArr) to UBound(loteArr)
        nroPagina = 1
        SM_idOrder = CLng(loteArr(i))
        call readHeaderOT(SM_idOrder)
	    Call dibujarPagina(oPDF, PAGINA_CABECERA)	    
	    'Imprimir marca de agua cancelado!
	    if SM_cdState = STATE_CANCELED then Call GF_CreateWaterMark(oPDF, 30, 520, GF_Traducir("CANCELADA  "), 100, "#FF0000", 315, 0.5)	    
	    'Si hay mas ordenes, creo una nueva pagina.
	    if (i <> UBound(loteArr)) then Call GF_newPage(oPDF)
    Next	    
End Function
'-----------------------------------------------------------------------------------------
Function dibujarPagina(oPDF, pTipoPagina)
	ultimaLinea = dibujarEncabezado(oPDF)	
	ultimaLinea = dibujarOrden(oPDF)
End Function
'-----------------------------------------------------------------------------------------
'Devuelve la posicion donde se puede seguir escribiendo
Function dibujarEncabezado(oPDF)
	Dim titulo
	Call GF_squareBox(oPDF, 2, 2, 590, 833, 0, "", "#0B3B0B", 2, PDF_SQUARE_ROUND)
	Call GF_squareBox(oPDF, 261, 17, 70, 50, 0, "", "#0B3B0B", 2, PDF_SQUARE_ROUND)
	Call GF_writeImage(oPDF, Server.MapPath("images\ADMlogo2.jpg"), 10, 15, 48, 48, 0)
	Call GF_setFont(oPDF,"ARIAL",28,FONT_STYLE_BOLD)
	Call GF_writeTextAlign(oPDF,20,25,GF_TRADUCIR("OT"), 550 , PDF_ALIGN_CENTER)
	Call GF_setFont(oPDF,"ARIAL",8,FONT_STYLE_NORMAL)
	Call GF_writeTextAlign(oPDF,10,80,GF_TRADUCIR("Departamento de Mantenimiento"), 550 , PDF_ALIGN_LEFT)
	Call GF_setFont(oPDF,"ARIAL",8,FONT_STYLE_NORMAL)
	Call GF_writeTextAlign(oPDF,30,80,GF_TRADUCIR(getDivisionDS(SM_idDivision)), 550 , PDF_ALIGN_RIGHT)
	GP_CONFIGURARMOMENTOS
	Call GF_setFont(oPDF,"COURIER",8,FONT_STYLE_NORMAL)
	Call GF_writeTextAlign(oPDF,5,5,GF_FN2DTE(momento), 580 , PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,5,5+pdf_currentFontSize,session("Usuario"), 580 , PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF, 10 , 840, "Pagina  " & nroPagina		 , 580 , PDF_ALIGN_RIGHT)
	Call GF_horizontalLine(oPDF,2,75,590)
	Call GF_setFont(oPDF,"COURIER",20,FONT_STYLE_NORMAL)
	Call GF_writeTextAlign(oPDF,20,50, SM_nroOrder, 560 , PDF_ALIGN_RIGHT)
	dibujarEncabezado = LINEA_POST_ENCABEZADO
End Function
'-----------------------------------------------------------------------------------------
Function dibujarOrden(oPDF)
call executeProcedureDb(DBSITE_SQL_INTRA, rsList, "TBLSMACTIVEEQUIPMENT_GET_FULL_BY_ID", SM_idActiveEquipment & "||0|| ||0|| || || ||1|| ")
if not rsList.eof then
	SM_cdActivation = trim(rsList("CDACTIVATION"))
	SM_dsActivation = trim(rsList("DSACTIVATION"))
	SM_dsSector = trim(rsList("DSSECTOR"))
	SM_activeCode = trim(rsList("CDACTIVECODE"))
end if	
	current_line = INICIO_PAGINA
	Call GF_setFont(oPDF,"ARIAL",12,FONT_STYLE_BOLD)
	Call GF_squareBox(oPDF, MARGEN, current_line, ANCHO_TOTAL, 17, 0, "#396E8F", "#000000", 1, PDF_SQUARE_ROUND)	
	call GF_setFontColor("FFFFFF")
	Call GF_setFont(oPDF,"ARIAL",12,FONT_STYLE_BOLD)
	Call GF_writeTextAlign(oPDF,MARGEN+2,current_line+2,GF_TRADUCIR(" Título: "), 550 , PDF_ALIGN_LEFT)
	auxText = SM_dsOrder
	if len(auxText) > 85 then auxText = left(auxText,85) & "..."
	Call GF_writeTextAlign(oPDF,MARGEN+45,current_line+2, auxText, 550 , PDF_ALIGN_LEFT)

	call GF_setFontColor("000000")
	Call GF_setFont(oPDF,"ARIAL",12,FONT_STYLE_BOLD)
	current_line = current_line + (SEPARACION) + 10
	Call GF_setFont(oPDF,"COURIER",10,FONT_STYLE_BOLD)
	Call GF_writeTextAlign(oPDF,X_COL_CAB_1,current_line,GF_TRADUCIR("Equipo.....: "), ANCHO_COL_CAB , PDF_ALIGN_LEFT)
	Call GF_setFont(oPDF,"COURIER",10,FONT_STYLE_NORMAL)
	Call GF_writeTextAlign(oPDF,X_COL_CAB_1 + 75,current_line, SM_cdActivation & "-" & SM_dsActivation, ANCHO_COL_CAB , PDF_ALIGN_LEFT)
	Call GF_setFont(oPDF,"COURIER",10,FONT_STYLE_BOLD)
	Call GF_writeTextAlign(oPDF,X_COL_CAB_2,current_line,GF_TRADUCIR("Solicitante: "), ANCHO_COL_CAB , PDF_ALIGN_LEFT)
	Call GF_setFont(oPDF,"COURIER",10,FONT_STYLE_NORMAL)
	Call GF_writeTextAlign(oPDF,X_COL_CAB_2 + 75,current_line, SM_dsApplicant, ANCHO_COL_CAB , PDF_ALIGN_LEFT)
	current_line = current_line + (SEPARACION)
	Call GF_setFont(oPDF,"COURIER",10,FONT_STYLE_BOLD)	
	Call GF_writeTextAlign(oPDF,X_COL_CAB_1,current_line,GF_TRADUCIR("Tipo Mant..: "), ANCHO_COL_CAB , PDF_ALIGN_LEFT)
	Call GF_setFont(oPDF,"COURIER",10,FONT_STYLE_NORMAL)
	Call GF_writeTextAlign(oPDF,X_COL_CAB_1 + 75,current_line,getDsMaintenanceType(SM_maintenanceType), ANCHO_COL_CAB , PDF_ALIGN_LEFT)
	Call GF_setFont(oPDF,"COURIER",10,FONT_STYLE_BOLD)	
	Call GF_writeTextAlign(oPDF,X_COL_CAB_2,current_line,GF_TRADUCIR("Tipo Orden.: "), ANCHO_COL_CAB , PDF_ALIGN_LEFT)
	Call GF_setFont(oPDF,"COURIER",10,FONT_STYLE_NORMAL)
	Call GF_writeTextAlign(oPDF,X_COL_CAB_2 + 75,current_line,getDsOrderType(SM_orderType), ANCHO_COL_CAB , PDF_ALIGN_LEFT)
	current_line = current_line + (SEPARACION)
	Call GF_setFont(oPDF,"COURIER",10,FONT_STYLE_BOLD)
	Call GF_writeTextAlign(oPDF,X_COL_CAB_1,current_line,GF_TRADUCIR("Sector.....: "), ANCHO_COL_CAB , PDF_ALIGN_LEFT)
	Call GF_setFont(oPDF,"COURIER",10,FONT_STYLE_NORMAL)
	Call GF_writeTextAlign(oPDF,X_COL_CAB_1 + 75,current_line,SM_dsSector, ANCHO_COL_CAB , PDF_ALIGN_LEFT)
	Call GF_setFont(oPDF,"COURIER",10,FONT_STYLE_BOLD)	
	Call GF_writeTextAlign(oPDF,X_COL_CAB_2,current_line,GF_TRADUCIR("Mano Obra..: "), ANCHO_COL_CAB , PDF_ALIGN_LEFT)
	Call GF_setFont(oPDF,"COURIER",10,FONT_STYLE_NORMAL)
	Call GF_writeTextAlign(oPDF,X_COL_CAB_2 + 75,current_line,left(SM_dsResponsableCompany,32), ANCHO_COL_CAB , PDF_ALIGN_LEFT)
	current_line = current_line + (SEPARACION)
	Call GF_setFont(oPDF,"COURIER",10,FONT_STYLE_BOLD)	
	Call GF_writeTextAlign(oPDF,X_COL_CAB_1,current_line,GF_TRADUCIR("Fec. Prog..: "), ANCHO_COL_CAB , PDF_ALIGN_LEFT)
	Call GF_setFont(oPDF,"COURIER",10,FONT_STYLE_NORMAL)
	Call GF_writeTextAlign(oPDF,X_COL_CAB_1 + 75,current_line,SM_scheduledDate, ANCHO_COL_CAB , PDF_ALIGN_LEFT)
	Call GF_setFont(oPDF,"COURIER",10,FONT_STYLE_BOLD)	
	Call GF_writeTextAlign(oPDF,X_COL_CAB_2,current_line,GF_TRADUCIR("Estado.....: "), ANCHO_COL_CAB , PDF_ALIGN_LEFT)
	Call GF_setFont(oPDF,"COURIER",10,FONT_STYLE_NORMAL)
	Call GF_writeTextAlign(oPDF,X_COL_CAB_2 + 75,current_line,getDsState(SM_cdState), ANCHO_COL_CAB , PDF_ALIGN_LEFT)

	current_line = current_line + (SEPARACION)
	Call GF_setFont(oPDF,"COURIER",10,FONT_STYLE_BOLD)
	Call GF_writeTextAlign(oPDF,X_COL_CAB_1,current_line,GF_TRADUCIR("Part. Pres.: "), ANCHO_COL_CAB , PDF_ALIGN_LEFT)
	Call GF_setFont(oPDF,"COURIER",10,FONT_STYLE_NORMAL)
		set rsObra = obtenerDescripcionCompletaDetalle(SM_idObra, SM_idBudgetArea, SM_idBudgetDetalle)
		if not rsObra.eof then
			myText = rsObra("DSOBRA") & ": " & rsObra("DSAREA") & "-" & rsObra("DSDETALLE")
			if len(myText)>75 then myText = left(mytext,72) & "..."
			Call GF_writeTextAlign(oPDF,X_COL_CAB_1 + 75,current_line,myText, ANCHO_COL_CAB , PDF_ALIGN_LEFT)
			'Response.write rsObra("DSOBRA") & ": " & rsObra("DSAREA") & "-" & rsObra("DSDETALLE")
		end if
	'Call GF_writeTextAlign(oPDF,X_COL_CAB_2 + 75,current_line,SM_IdObra & ": " & SM_idBudgetArea & "-" & SM_idBudgetDetalle, ANCHO_COL_CAB , PDF_ALIGN_LEFT)
	current_line = current_line + (SEPARACION) 
	anchoMArco = 80	
	if SM_cdState = STATE_FINISHED then
		Call GF_setFont(oPDF,"COURIER",10,FONT_STYLE_BOLD)	
		Call GF_writeTextAlign(oPDF,X_COL_CAB_1,current_line,GF_TRADUCIR("Fec. Inicio: "), ANCHO_COL_CAB , PDF_ALIGN_LEFT)
		Call GF_setFont(oPDF,"COURIER",10,FONT_STYLE_NORMAL)
		Call GF_writeTextAlign(oPDF,X_COL_CAB_1 + 75,current_line,SM_startDate, ANCHO_COL_CAB , PDF_ALIGN_LEFT)
		Call GF_setFont(oPDF,"COURIER",10,FONT_STYLE_BOLD)	
		Call GF_writeTextAlign(oPDF,X_COL_CAB_2,current_line,GF_TRADUCIR("Fec. Fin...: "), ANCHO_COL_CAB , PDF_ALIGN_LEFT)
		Call GF_setFont(oPDF,"COURIER",10,FONT_STYLE_NORMAL)
		Call GF_writeTextAlign(oPDF,X_COL_CAB_2 + 75,current_line,SM_finishedDate, ANCHO_COL_CAB , PDF_ALIGN_LEFT)

		current_line = current_line + (SEPARACION)	
		anchoMArco = anchoMArco + 15
	end if
	if SM_OTFrequencyUnit <> ORDER_FREQ_UNIQUE then
		Call GF_setFont(oPDF,"COURIER",10,FONT_STYLE_BOLD)	
		Call GF_writeTextAlign(oPDF,X_COL_CAB_1,current_line,GF_TRADUCIR("Frecuencia.: "), ANCHO_COL_CAB , PDF_ALIGN_LEFT)
		Call GF_setFont(oPDF,"COURIER",10,FONT_STYLE_NORMAL)
		Call GF_writeTextAlign(oPDF,X_COL_CAB_1 + 75,current_line,getFrequency(SM_OTFrequencyUnit,SM_OTFrequencyQuantity), ANCHO_COL_CAB , PDF_ALIGN_LEFT)
		current_line = current_line + (SEPARACION)	
		anchoMArco = anchoMArco + 15
	end if
	Call GF_squareBoxTransparent(oPDF, MARGEN, INICIO_PAGINA + SEPARACION+6, 550, anchoMArco, 0, "", "#0B3B0B", 1, PDF_SQUARE_ROUND)	

	Call GF_setFont(oPDF,"ARIAL",10,FONT_STYLE_BOLD)
	if SM_cdState = STATE_FINISHED then
		auxText = " Realizado"
	else
		auxText = " a Realizar"
	end if	
	'Call GF_horizontalLine(oPDF,MARGEN,current_line,550)
	current_line = current_line + (SEPARACION) 
	Call GF_writeTextAlign(oPDF,MARGEN,current_line,GF_TRADUCIR("Descripción del Trabajo" & auxText) , ANCHO_COL_CAB , PDF_ALIGN_LEFT)
	current_line = current_line + (SEPARACION) 
	Call GF_setFont(oPDF,"ARIAL",10,FONT_STYLE_NORMAL)
	Call GF_writeTextAlign(oPDF,MARGEN+8,current_line,GF_TRADUCIR("Tareas"), ANCHO_COL_CAB , PDF_ALIGN_LEFT)
	current_line = current_line + (SEPARACION) 
	'Leer Detalles
	call dibujarTitulosTasks(current_line)
	call initTasksOT()	
	while readNextTaskOt()
		Call GF_setFont(oPDF,"COURIER",8,FONT_STYLE_NORMAL)
		if i mod 2  then
			Call GF_squareBox(oPDF, X_COL_TASK_1, current_line-2, ANCHO_TOTAL-20, 15, 0, "#d3d3d3", "#000000", 0, PDF_SQUARE_NORMAL)	
		end if	
		Call GF_writeTextAlign(oPDF,X_COL_TASK_1,current_line+2,SM_NROTASK,  ANCHO_COL_TASK_1 , PDF_ALIGN_CENTER)
		Call GF_writeTextAlign(oPDF,X_COL_TASK_2+2,current_line+2,SM_DSTASK,   ANCHO_COL_TASK_2 , PDF_ALIGN_LEFT)
		'Call GF_writeTextAlign(oPDF,X_COL_TASK_3,current_line+2,SM_DONETASK, ANCHO_COL_TASK_3 , PDF_ALIGN_CENTER)
		if SM_DONETASK = SM_TASK_DONE_YES then 
			Call GF_writeImage(oPDF, Server.MapPath("images\mantenimiento\checked.gif"), X_COL_TASK_3+5, current_line-1, 12, 12, 0)
		else
			Call GF_writeImage(oPDF, Server.MapPath("images\mantenimiento\unchecked.gif"), X_COL_TASK_3+5, current_line-1, 12, 12, 0)
		end if	
		call nuevaLinea(SEC_TASKS)
		i = i + 1
	wend	
	current_line = current_line + 8
	Call GF_setFont(oPDF,"ARIAL",10,FONT_STYLE_NORMAL)
	Call GF_writeTextAlign(oPDF,MARGEN+8,current_line,GF_TRADUCIR("Piezas/Repuestos"), ANCHO_COL_CAB , PDF_ALIGN_LEFT)
	current_line = current_line + SEPARACION 
	'Leer Repuestos
	call dibujarTitulosItems(current_line)
	Call GF_setFont(oPDF,"COURIER",8,FONT_STYLE_NORMAL)
	call initItemsOT()	
	while readNextItemOt()
		Call GF_setFont(oPDF,"COURIER",8,FONT_STYLE_NORMAL)
		if i mod 2  then
			Call GF_squareBox(oPDF, X_COL_TASK_1, current_line-2, ANCHO_TOTAL-20, 12, 0, "#d3d3d3", "#000000", 0, PDF_SQUARE_NORMAL)	
		end if	
		Call GF_writeTextAlign(oPDF,X_COL_ITEM_NRO		,current_line,SM_NROITEM			, ANCHO_COL_ITEM_NRO		, PDF_ALIGN_CENTER)
		Call GF_writeTextAlign(oPDF,X_COL_ITEM_ARTID	,current_line,SM_IDITEM				, ANCHO_COL_ITEM_ARTID		, PDF_ALIGN_CENTER)
		Call GF_writeTextAlign(oPDF,X_COL_ITEM_ARTDS+2	,current_line,SM_DSITEM				, ANCHO_COL_ITEM_ARTDS		, PDF_ALIGN_LEFT)
		Call GF_writeTextAlign(oPDF,X_COL_ITEM_CANTPROG-2,current_line,GF_EDIT_DECIMALS(Cdbl(SM_PROGRAMQUANTITYITEM)*100,2), ANCHO_COL_ITEM_CANTPROG	, PDF_ALIGN_RIGHT)
		Call GF_writeTextAlign(oPDF,X_COL_ITEM_CANTREAL-2,current_line,GF_EDIT_DECIMALS(Cdbl(SM_REALQUANTITYITEM)*100,2)	, ANCHO_COL_ITEM_CANTREAL	, PDF_ALIGN_RIGHT)
		myText = SM_IDPMITEM
		if myText = 0 then myText = "-"
		Call GF_writeTextAlign(oPDF,X_COL_ITEM_PM		,current_line,myText			, ANCHO_COL_ITEM_PM			, PDF_ALIGN_CENTER)
		call nuevaLinea(SEC_ITEMS) 
		i = i + 1
	wend	
	call dibujarPie(current_line)	
End Function
'-----------------------------------------------------------------------------------------
function dibujarTitulosTasks(pY)
	Call GF_squareBox(oPDF, X_COL_TASK_1, pY, ANCHO_COL_TASK_1	, 13, 0, "#396E8F", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, X_COL_TASK_2, pY, ANCHO_COL_TASK_2	, 13, 0, "#396E8F", "#000000", 1, PDF_SQUARE_NORMAL)	
	Call GF_squareBox(oPDF, X_COL_TASK_3, pY, ANCHO_COL_TASK_3	, 13, 0, "#396E8F", "#000000", 1, PDF_SQUARE_NORMAL)
	call GF_setFont(oPDF,"ARIAL",8,8)
	call GF_setFontColor("FFFFFF")
	Call GF_writeTextAlign(oPDF, X_COL_TASK_1-1, pY+2	, GF_TRADUCIR("Nro")			, ANCHO_COL_TASK_1	, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, X_COL_TASK_2, pY+2	, GF_TRADUCIR("Descripción")	, ANCHO_COL_TASK_2	, PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(oPDF, X_COL_TASK_3, pY+2	, GF_TRADUCIR("OK")			, ANCHO_COL_TASK_3	, PDF_ALIGN_CENTER)	
	call GF_setFontColor("000000")
	call nuevaLinea("")
end function
'-----------------------------------------------------------------------------------------
function dibujarTitulosItems(pY)
	Call GF_squareBox(oPDF, X_COL_ITEM_NRO		, pY, ANCHO_COL_ITEM_NRO		, 26, 0, "#396E8F", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, X_COL_ITEM_ARTICULO	, pY, ANCHO_COL_ITEM_ARTICULO	, 13, 0, "#396E8F", "#000000", 1, PDF_SQUARE_NORMAL)	
	Call GF_squareBox(oPDF, X_COL_ITEM_CANTIDAD	, pY, ANCHO_COL_ITEM_CANTIDAD	, 13, 0, "#396E8F", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, X_COL_ITEM_PM		, pY, ANCHO_COL_ITEM_PM			, 26, 0, "#396E8F", "#000000", 1, PDF_SQUARE_NORMAL)
	
	Call GF_squareBox(oPDF, X_COL_ITEM_ARTID	, pY+13, ANCHO_COL_ITEM_ARTID	, 13, 0, "#396E8F", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, X_COL_ITEM_ARTDS	, pY+13, ANCHO_COL_ITEM_ARTDS	, 13, 0, "#396E8F", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, X_COL_ITEM_CANTPROG	, pY+13, ANCHO_COL_ITEM_CANTPROG, 13, 0, "#396E8F", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, X_COL_ITEM_CANTREAL	, pY+13, ANCHO_COL_ITEM_CANTREAL, 13, 0, "#396E8F", "#000000", 1, PDF_SQUARE_NORMAL)
	
	call GF_setFont(oPDF,"ARIAL",8,8)
	call GF_setFontColor("FFFFFF")
	Call GF_writeTextAlign(oPDF, X_COL_ITEM_NRO-1	, pY+8, GF_TRADUCIR("Nro")				, ANCHO_COL_ITEM_NRO		, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, X_COL_ITEM_ARTICULO, pY+2, GF_TRADUCIR("Pieza/Repuesto")	, ANCHO_COL_ITEM_ARTICULO	, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, X_COL_ITEM_CANTIDAD, pY+2, GF_TRADUCIR("Cantidad")			, ANCHO_COL_ITEM_CANTIDAD	, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, X_COL_ITEM_PM		, pY+8, GF_TRADUCIR("PM")				, ANCHO_COL_ITEM_PM			, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, X_COL_ITEM_ARTID	, pY+15, GF_TRADUCIR("Id")				, ANCHO_COL_ITEM_ARTID		, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, X_COL_ITEM_ARTDS	, pY+15, GF_TRADUCIR("Descripción")		, ANCHO_COL_ITEM_ARTDS		, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, X_COL_ITEM_CANTPROG, pY+15, GF_TRADUCIR("Prog.")			, ANCHO_COL_ITEM_CANTPROG	, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, X_COL_ITEM_CANTREAL, pY+15, GF_TRADUCIR("Real")			, ANCHO_COL_ITEM_CANTREAL	, PDF_ALIGN_CENTER)						
	call GF_setFontColor("000000")
	call nuevaLinea("")
	call nuevaLinea("")
end function
'-----------------------------------------------------------------------------------------
function dibujarTitulosTitulo(pAuxText)
	call nuevaLinea(SEC_OBVS)
	Call GF_squareBoxTransparent(oPDF, X_COL_ITEM_NRO, current_line+15, ANCHO_TOTAL-20, 90, 0, "", "#0B3B0B", 1, PDF_SQUARE_ROUND)	
	Call GF_setFont(oPDF,"ARIAL",10,FONT_STYLE_BOLD)
	myText = "Observaciones del Trabajo" & auxText
	if SM_cdState = STATE_CANCELED then myText = " Motivo de cancelación: "
	Call GF_writeTextAlign(oPDF,MARGEN,current_line,GF_TRADUCIR(myText) & " " & pAuxText, ANCHO_COL_CAB , PDF_ALIGN_LEFT)
	call nuevaLinea(SEC_OBVS)
end function

'-----------------------------------------------------------------------------------------
function dibujarTitulosEjecutante(pDs)
	Call GF_setFont(oPDF,"ARIAL",10,FONT_STYLE_BOLD)
	Call GF_writeTextAlign(oPDF,X_COL_CAB_1,current_line,GF_TRADUCIR("Ejecutante:"), ANCHO_COL_CAB , PDF_ALIGN_LEFT)
	Call GF_setFont(oPDF,"courier",10,FONT_STYLE_NORMAL)
	Call GF_writeTextAlign(oPDF,X_COL_CAB_1 + (MARGEN*3),current_line,pDs, ANCHO_COL_CAB , PDF_ALIGN_LEFT)
	Call GF_setFont(oPDF,"ARIAL",10,FONT_STYLE_BOLD)
	Call GF_writeTextAlign(oPDF,X_COL_CAB_1+X_COL_CAB_2,current_line,GF_TRADUCIR("Firma:"), ANCHO_COL_CAB , PDF_ALIGN_LEFT)
	Call GF_setFont(oPDF,"courier",10,FONT_STYLE_NORMAL)
	Call GF_writeTextAlign(oPDF,X_COL_CAB_1+X_COL_CAB_2 + (MARGEN*2),current_line,"....................", ANCHO_COL_CAB , PDF_ALIGN_LEFT)
	Call GF_setFont(oPDF,"ARIAL",10,FONT_STYLE_BOLD)
end function
'-----------------------------------------------------------------------------------------
function dibujarPie(pY)        
    
	current_line = current_line + 8
	Call GF_horizontalLine(oPDF,MARGEN+10 ,current_line - 8,530)
	fin_observaciones_line = current_line  + 150
	'call nuevaLinea("")
	'call nuevaLinea("")	
	Call dibujarTitulosTitulo("")
	Call GF_setFont(oPDF,"courier",8,FONT_STYLE_NORMAL)	
	Call GF_writeTextPlus(oPDF, X_COL_ITEM_NRO + 4, current_line, GF_TRADUCIR(SM_OBSERVATIONS), 390, 8, PDF_ALIGN_LEFT)				
	current_line = fin_observaciones_line
	Call nuevaLinea(SEC_OBVS)
	call dibujarTitulosEjecutante(SM_dsEXECUTEDBY)
end function
'-----------------------------------------------------------------------------------------
Function nuevaPagina(pSector)
	Call GF_newPage(oPDF)
	nroPagina = nroPagina + 1
	current_line = INICIO_PAGINA + (SEPARATION) + 4	
	call dibujarEncabezado(oPDF)
	if pSector = SEC_ITEMS then
		call dibujarTitulosItems(current_line)
	elseif pSector = SEC_TASKS then
		call dibujarTitulosTasks(current_line)
	elseif pSector = SEC_OBVS then
		call dibujarTitulosTitulo("(Cont. )")
	elseif pSector = SEC_EXE then
		call dibujarTitulosEjecutante("(Cont. )")
	else
		'Nada
	end if	
End Function
'-----------------------------------------------------------------------------------------
function nuevaLinea(pSector)
if current_line > PAGE_HEIGHT_SIZE then	
	'Response.Write "<hr>A nueva pagina"
	Call nuevaPagina(pSector)
else
	current_line = current_line + (SEPARACION)
end if
end function
%>