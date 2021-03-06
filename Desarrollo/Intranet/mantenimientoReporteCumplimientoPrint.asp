<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosPDF.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosPuertos.asp"-->
<!--#include file="Includes/procedimientosMantenimiento.asp"-->
<%
'-----------------------------------------------------------------------------------------
Function dibujarEncabezado(oPDF, pDsDivi, anio, mes)	
	Call GF_squareBox(oPDF, 2, 10, 590, 828, 0, "", "#0B3B0B", 2, PDF_SQUARE_ROUND)	
	Call GF_writeImage(oPDF, Server.MapPath("Images\logo1.jpg"),10, 15, 81, 75, 0)
	Call GF_setFont(oPDF,"ARIAL", 20,FONT_STYLE_BOLD)
	Call GF_writeTextAlign(oPDF,20,65,"Reporte de Cumplimiento", 550 , PDF_ALIGN_CENTER)	
	Call GF_setFont(oPDF,"ARIAL",8,FONT_STYLE_NORMAL)	
	Call GF_writeTextAlign(oPDF,30,110, pDsDivi, 550 , PDF_ALIGN_RIGHT)		
	Call GF_setFont(oPDF,"ARIAL", 14,FONT_STYLE_BOLD)
	Call GF_writeTextAlign(oPDF,20, 120, getNameOfMonth(mes) & " " & anio, 550 , PDF_ALIGN_CENTER)	
	Call GF_setFont(oPDF,"COURIER",8,FONT_STYLE_NORMAL)
	Call GF_writeTextAlign(oPDF,5,15, GF_FN2DTE(session("MmtoDato")), 580 , PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,5,25,session("Usuario"), 580 , PDF_ALIGN_RIGHT)	
	Call GF_horizontalLine(oPDF,2,100,590)		
End Function
'--------------------------------------------------------------------------------------------------------------------------
Function cargarValores(pDivi, pFd, pFh, ByRef totalProg, ByRef totalIniMes, ByRef totalIniAnt, ByRef totalIniProx, ByRef totalFinMes, ByRef totalFinAnt, ByRef totalFinProx)
    Dim rs, strSQL
    
    'Total de tareas programadas. (Mes Actual)
    strSQL="Select count(*) CANT from TBLSMORDER OT where IDDIVISION=" & pDivi & " and SCHEDULEDDATE >= " & pFd & " and SCHEDULEDDATE <= " & pFh & " and CDSTATE <> " & STATE_CANCELED
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
    totalProg = 0
    if (not rs.eof) then totalProg = rs("CANT")
    
    'Total de tareas iniciadas del mes (Terminan el mes iniciadas y sin finalizar.)
    strSQL="Select count(*) CANT from TBLSMORDER OT where IDDIVISION=" & pDivi & " and SCHEDULEDDATE >= " & pFd & " and SCHEDULEDDATE <= " & pFh & " and STARTDATE <= " & pFh & " and (FINISHEDDATE is Null or FINISHEDDATE > " & pFh & ") and CDSTATE <> " & STATE_CANCELED
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
    totalIniMes = 0
    if (not rs.eof) then totalIniMes = rs("CANT")
    
    'Total de tareas iniciadas meses Anteriores
    strSQL="Select count(*) CANT from TBLSMORDER OT where IDDIVISION=" & pDivi & " and SCHEDULEDDATE < " & pFd & " and STARTDATE <= " & pFh & " and (FINISHEDDATE is Null or FINISHEDDATE > " & pFh & ") and CDSTATE <> " & STATE_CANCELED
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
    totalIniAnt = 0
    if (not rs.eof) then totalIniAnt = rs("CANT")
    
    'Total de tareas iniciadas de meses proximos
    strSQL="Select count(*) CANT from TBLSMORDER OT where IDDIVISION=" & pDivi & " and SCHEDULEDDATE > " & pFh & " and STARTDATE <= " & pFh & " and (FINISHEDDATE is Null or FINISHEDDATE > " & pFh & ") and CDSTATE <> " & STATE_CANCELED
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
    totalIniProx = 0
    if (not rs.eof) then totalIniProx = rs("CANT")
    
    'Total de tareas finalizadas del mes
    strSQL="Select count(*) CANT from TBLSMORDER OT where IDDIVISION=" & pDivi & " and SCHEDULEDDATE >= " & pFd & " and SCHEDULEDDATE <= " & pFh & " and FINISHEDDATE <= " & pFh & " and CDSTATE <> " & STATE_CANCELED
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
    totalFinMes = 0
    if (not rs.eof) then totalFinMes = rs("CANT")
    
    'Total de tareas finalizadas de meses anterioriores
    strSQL="Select count(*) CANT from TBLSMORDER OT where IDDIVISION=" & pDivi & " and SCHEDULEDDATE < " & pFd & " and FINISHEDDATE >= " & pFd & " and FINISHEDDATE <= " & pFh & " and CDSTATE <> " & STATE_CANCELED
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
    totalFinAnt = 0
    if (not rs.eof) then totalFinAnt = rs("CANT")
    
    'Total de tareas finalizadas de meses proximos
    strSQL="Select count(*) CANT from TBLSMORDER OT where IDDIVISION=" & pDivi & " and SCHEDULEDDATE > " & pFh & " and FINISHEDDATE >= " & pFd & " and FINISHEDDATE <= " & pFh & " and CDSTATE <> " & STATE_CANCELED
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
    totalFinProx = 0
    if (not rs.eof) then totalFinProx = rs("CANT")
    
End Function
'--------------------------------------------------------------------------------------------------------------------------
Function metricaProgramaMes(oPDF, totalProg, totalIniMes, totalFinMes)
    
    Dim Xo, Yo, Ho, datos(2, 2), totalNoIniciado, porcIni, porcFin, porcNoIni
    
    porcIni = 0
    porcFin = 0
    porcNoIni = 0
    totalNoIniciado = 0        
    if (totalProg > 0) then
        porcIni = round(CDbl(CLng(totalIniMes)*100/CLng(totalProg)), 2)
        porcFin = round(CDbl(CLng(totalFinMes)*100/CLng(totalProg)), 2)
        porcNoIni = 100-porcIni-porcFin
        totalNoIniciado = CLng(totalProg) - CLng(totalIniMes) - CLng(totalFinMes)    
    end if
    
    Xo=50
    Yo=200
    Ho = 15
    Call GF_setFont(oPDF,"ARIAL",8,FONT_STYLE_BOLD)
    'TITULO
    Call GF_writeText(oPDF, Xo, Yo, "1.- EJECUCIÓN DE LA PROGRAMACIÓN DEL MES", 0)
    'TABLA
    '   Filas
    Call GF_horizontalLine(oPDF, Xo+250,Yo+20,120)
    Call GF_squareBox(oPDF, Xo+120, Yo+35, 250, 15, 0, "", "#0B3B0B", 1, PDF_SQUARE_NORMAL)
    Call GF_squareBox(oPDF, Xo+120, Yo+50, 250, 15, 0, "", "#0B3B0B", 1, PDF_SQUARE_NORMAL)
    Call GF_squareBox(oPDF, Xo+120, Yo+65, 250, 15, 0, "", "#0B3B0B", 1, PDF_SQUARE_NORMAL)
    Call GF_squareBox(oPDF, Xo+120, Yo+80, 250, 15, 0, "", "#0B3B0B", 1, PDF_SQUARE_NORMAL)
    '   Columnas    
    Call GF_verticalLine(oPDF, Xo+250, Yo+20, 75)
    Call GF_verticalLine(oPDF, Xo+310, Yo+20, 75)
    Call GF_verticalLine(oPDF, Xo+370, Yo+20, 75)
    'TEXTO
    Call GF_writeTextAlign(oPDF,Xo+250,Yo+23, "Cantidad", 60, PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(oPDF,Xo+310,Yo+23, "%", 60, PDF_ALIGN_CENTER)    
    Call GF_setFont(oPDF,"ARIAL",8,FONT_STYLE_NORMAL)
    Call GF_writeText(oPDF,Xo+122,Yo+38,"Tareas En Curso "  , 0)
    Call GF_writeTextAlign(oPDF,Xo+250,Yo+38, CLng(totalIniMes), 60, PDF_ALIGN_CENTER)    
    Call GF_writeTextAlign(oPDF,Xo+310,Yo+38, GF_EDIT_DECIMALS(porcIni*100, 2) & " %", 60, PDF_ALIGN_CENTER)
    Call GF_writeText(oPDF,Xo+122,Yo+53,"Tareas Finalizadas"    , 0)
    Call GF_writeTextAlign(oPDF,Xo+250,Yo+53, totalFinMes, 60, PDF_ALIGN_CENTER)    
    Call GF_writeTextAlign(oPDF,Xo+310,Yo+53, GF_EDIT_DECIMALS(porcFin*100, 2) & " %", 60, PDF_ALIGN_CENTER)
    Call GF_writeText(oPDF,Xo+122,Yo+68,"Tareas No Iniciadas"  , 0)    
    Call GF_writeTextAlign(oPDF,Xo+250,Yo+68, totalNoIniciado, 60, PDF_ALIGN_CENTER)    
    Call GF_writeTextAlign(oPDF,Xo+310,Yo+68, GF_EDIT_DECIMALS(porcNoIni*100, 2) & " %", 60, PDF_ALIGN_CENTER)
    Call GF_setFont(oPDF,"ARIAL",9,FONT_STYLE_BOLD)
    Call GF_writeText(oPDF,Xo+122,Yo+83,"Total Tareas Programadas"  , 0)    
    Call GF_writeTextAlign(oPDF,Xo+250,Yo+83, totalProg, 60, PDF_ALIGN_CENTER)           
    'GRAFICO
    if (totalProg > 0) then
        datos(0, 0) = "Iniciadas"
        datos(0, 1) = totalIniMes  
        datos(0, 2) = "#3366FF"
        datos(1, 0) = "Finalizadas"
        datos(1, 1) = totalFinMes
        datos(1, 2) = "#33FF33"
        datos(2, 0) = "No Iniciadas"    
        datos(2, 1) = totalNoIniciado
        datos(2, 2) = "#AAAAAA"
        Call GF_setFont(oPDF, "ARIAL", 8, FONT_STYLE_NORMAL)
        Call drawPieChart(p_oPDF, Xo+80, Yo+100, 300, 150, "", datos)
    end if        
    
End Function
'--------------------------------------------------------------------------------------------------------------------------
Function metricaTareas(oPDF, Xo, Yo, Titulo, totalIniMes, totalIniAnt, totalIniProx)
    
    Dim Ho, datos(2, 2), totalIni, porcMesAnt, porcMes, porcMesProx
    
    totalIni = CLng(totalIniAnt) + CLng(totalIniMes) + CLng(totalIniProx)
    porcMesAnt = 0
    porcMes = 0
    porcMesProx = 0
    if (totalIni > 0) then 
        porcMesAnt = CDbl(CLng(totalIniAnt)*100/CLng(totalIni))
        porcMes = CDbl(CLng(totalIniMes)*100/CLng(totalIni))
        porcMesProx = CDbl(CLng(totalIniProx)*100/CLng(totalIni))
    end if
    
    Ho = 15
    Call GF_setFont(oPDF,"ARIAL",8,FONT_STYLE_BOLD)
    'TITULO
    Call GF_writeText(oPDF, Xo, Yo, Titulo, 0)
    'TABLA
    '   Filas    
    Call GF_horizontalLine(oPDF, Xo+250,Yo+20,120)
    Call GF_squareBox(oPDF, Xo+120, Yo+35, 250, 15, 0, "", "#0B3B0B", 1, PDF_SQUARE_NORMAL)
    Call GF_squareBox(oPDF, Xo+120, Yo+50, 250, 15, 0, "", "#0B3B0B", 1, PDF_SQUARE_NORMAL)
    Call GF_squareBox(oPDF, Xo+120, Yo+65, 250, 15, 0, "", "#0B3B0B", 1, PDF_SQUARE_NORMAL)
    Call GF_squareBox(oPDF, Xo+120, Yo+80, 250, 15, 0, "", "#0B3B0B", 1, PDF_SQUARE_NORMAL)    
    '   Columnas    
    Call GF_verticalLine(oPDF, Xo+250, Yo+20, 75)
    Call GF_verticalLine(oPDF, Xo+310, Yo+20, 75)
    Call GF_verticalLine(oPDF, Xo+370, Yo+20, 75)
    'TEXTO    
    Call GF_writeTextAlign(oPDF,Xo+250,Yo+23, "Cantidad", 60, PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(oPDF,Xo+310,Yo+23, "%", 60, PDF_ALIGN_CENTER)    
    Call GF_setFont(oPDF,"ARIAL",8,FONT_STYLE_NORMAL)
    Call GF_writeText(oPDF,Xo+122,Yo+38,"Tareas de Meses Anteriores "  , 0)    
    Call GF_writeTextAlign(oPDF,Xo+250,Yo+38, totalIniAnt, 60, PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(oPDF,Xo+310,Yo+38, GF_EDIT_DECIMALS(porcMesAnt*100, 2) & " %", 60, PDF_ALIGN_CENTER)
    Call GF_writeText(oPDF,Xo+122,Yo+53,"Tareas del Mes Actual"    , 0)    
    Call GF_writeTextAlign(oPDF,Xo+250,Yo+53, totalIniMes, 60, PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(oPDF,Xo+310,Yo+53, GF_EDIT_DECIMALS(porcMes*100, 2) & " %", 60, PDF_ALIGN_CENTER)
    Call GF_writeText(oPDF,Xo+122,Yo+68,"Tareas de Meses Próximos"  , 0)    
    Call GF_writeTextAlign(oPDF,Xo+250,Yo+68, totalIniProx, 60, PDF_ALIGN_CENTER)    
    Call GF_writeTextAlign(oPDF,Xo+310,Yo+68, GF_EDIT_DECIMALS(porcMesProx*100, 2) & " %", 60, PDF_ALIGN_CENTER)
    Call GF_setFont(oPDF,"ARIAL",9,FONT_STYLE_BOLD)
    Call GF_writeText(oPDF,Xo+122,Yo+83,"Total Tareas En Curso"  , 0)    
    Call GF_writeTextAlign(oPDF,Xo+250,Yo+83, totalIni, 60, PDF_ALIGN_CENTER)           
    'GRAFICO
    if (totalIni > 0) then 
        datos(0, 0) = "Mes Anterior"
        datos(0, 1) = totalIniAnt  
        datos(0, 2) = "#3366FF"
        datos(1, 0) = "Mes Actual"
        datos(1, 1) = totalIniMes
        datos(1, 2) = "#33FF33"
        datos(2, 0) = "Meses Proximos"    
        datos(2, 1) = totalIniProx
        datos(2, 2) = "#AAAAAA"
        Call GF_setFont(oPDF, "ARIAL", 8, FONT_STYLE_NORMAL)
        Call drawPieChart(p_oPDF, Xo+80, Yo+100, 300, 150, "", datos)
    end if        
    
End Function
'--------------------------------------------------------------------------------------------------------------------------
Function armadoPDF(pDivi, anio, mes)    
    Dim oPDF
    Dim dsDivi, rsList
    Dim totalIniMes, totalIniAnt, totalIniProx
    Dim totalFinMes, totalFinAnt, totalFinProx
    Dim totalProg, fnDesde, fnHasta
    
    Set oPDF = GF_createPDF("PDFTemp")
	Call GF_setPDFMODE(PDF_STREAM_MODE)
	
	fnDesde = anio & mes & "01"
    fnHasta = anio & mes & getLastDayOfMonth(anio, mes)

    Call executeProcedureDb(DBSITE_SQL_INTRA, rsList, "TBLDIVISIONES_GET_BY_LIST", pDivi)
	if (not rsList.eof) then dsDivi=rsList("DSDIVISION")
	
    Call cargarValores(pDivi, fnDesde, fnHasta, totalProg, totalIniMes, totalIniAnt, totalIniProx, totalFinMes, totalFinAnt, totalFinProx)
    'Armo el reporte
	Call dibujarEncabezado(oPDF, dsDivi, anio, mes)	
	'-->Call metricaProgramaMes(oPDF, totalProg, totalIniMes, totalFinMes)
	Call metricaTareas(oPDF, 50, 200, "1.- TAREAS FINALIZADAS DURANTE EL MES", totalFinMes, totalFinAnt, totalFinProx)	
	Call metricaTareas(oPDF, 50, 500, "2.- TAREAS EN CURSO A FIN DE MES",  totalIniMes, totalIniAnt, totalIniProx)
	'-->Call GF_newPage(oPDF)
	'-->Call dibujarEncabezado(oPDF, dsDivi, anio, mes)
	'-->Call metricaTareas(oPDF, 50, 200, "3.- TAREAS FINALIZADAS DURANTE EL MES", totalFinMes, totalFinAnt, totalFinProx)	
	Call GF_closePDF(oPDF)
	    
End Function
'****************************************************************************
'****************************************************************************
'*****                      COMIENZO DE LA PAGINA                       *****  
'****************************************************************************
'****************************************************************************
Dim divi, mes, anio

Call GP_CONFIGURARMOMENTOS

divi = GF_PARAMETROS7("div", 0, 6)
mes = GF_PARAMETROS7("mes", "", 6)
if (mes = "") then fd = Right(Left(session("MmtoDato"), 6), 2)
anio = GF_PARAMETROS7("anio", "", 6)
if (anio = "") then fh = Left(session("MmtoDato"), 4)

Call armadoPDF(divi, anio, mes)
%>