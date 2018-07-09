<!--#include file="../Includes/procedimientosCompras.asp"-->
<!--#include file="../Includes/procedimientosCupos.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->
<!--#include file="../Includes/procedimientosPuertos.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientosFechas.asp"-->
<!--#include file="../Includes/procedimientosPDF.asp"-->
<%

'/***************************************************************\
' *         REPORTE DE CUPOS PARA RESOLUCION0025/13
' *          AGENCIA PROVINCIAL DE SEGURIDAD VIAL
' *                 PROVINCIA DE SANTA FE
'\***************************************************************/
Const MAX_LINEAS_PAGINA = 53
'-----------------------------------------------------------------------------------------------------------------
Function armarEstructuraTabla(fecha, pPagina, pto)
    Dim i, Xtxt, Ytxt
    
    'Cuandos de linea.
    for i=0 to 8
        Call GF_squareBox(oPDF, Xo, Yo + (hfila*i), wtabla, hfila, 0, "#FFFFFF", "#000000", 1, PDF_SQUARE_NORMAL)    
    next
    Call GF_squareBox(oPDF, Xo, Yo + (hfila*i), wtabla, hfila*2, 0, "#FFFFFF", "#000000", 1, PDF_SQUARE_NORMAL)
    'Lineas verticales de division.
    Call GF_drawLine(oPDF, Xo + (wtabla/2), Yo + (hfila*3), Xo + (wtabla/2), Yo + (hfila*8))
    Call GF_drawLine(oPDF, Xo + (wtabla*3/4), Yo + (hfila*9), Xo + (wtabla*3/4), Yo + (hfila*11))
    'Textos fijos de la tabla
    Xtxt = Xo + 2
    Ytxt = Yo + 2
    
    Call GF_setFont(oPDF,"ARIAL", 12, FONT_STYLE_BOLD)
    Call GF_writeTextAlign(oPDF, Xtxt, Ytxt, "PLANILLA DE INFORMACIÓN DE ASIGNACIÓN DE ESPACIO FÍSICO", wtabla, PDF_ALIGN_CENTER)
    Call GF_setFont(oPDF,"ARIAL", 10, FONT_STYLE_BOLD)
    Call GF_writeText(oPDF, Xtxt, Ytxt + (hfila*2), "1. DATOS DE LA EMPRESA EMISORA DEL CUPO", 0)	
    Call GF_writeText(oPDF, Xtxt, Ytxt + (hfila*8), "2. DATOS DEL DESTINATARIO DEL CUPO", 0)	    
    Call GF_writeText(oPDF, Xtxt, Ytxt + (hfila*9), "Nombre o Razón Social", 0)    
    Xtxt = Xo + (wtabla*3/4) + 2
    Call GF_writeTextPlus(oPDF, Xtxt, Ytxt + (hfila*9), "Cantidad de Cupos Asignados", wtabla/4, hfila, PDF_ALIGN_LEFT)
    
    Call GF_setFont(oPDF,"ARIAL", 10, FONT_STYLE_NORMAL)
    Xtxt = Xo + 2
    Call GF_writeText(oPDF, Xtxt, Ytxt + hfila, "Fecha de Emisión: " + GF_FN2DTE(session("MmtoDato")), 0)
    Call GF_writeText(oPDF, Xtxt, Ytxt + (hfila*3), "Nombre o Razón Social:", 0)
    Call GF_writeText(oPDF, Xtxt, Ytxt + (hfila*4), "C.U.I.T.:", 0)
    Call GF_writeText(oPDF, Xtxt, Ytxt + (hfila*5), "Fecha asignada para la descarga:", 0)
    Call GF_writeText(oPDF, Xtxt, Ytxt + (hfila*6), "Cantidad Total de Cupos otorgados en el día:", 0)
    Call GF_writeText(oPDF, Xtxt, Ytxt + (hfila*7), "Domicilio de Descarga:", 0)    
    
    'Datos de la Terminal
    Xtxt = Xo + (wtabla/2) + 2
    Call GF_setFont(oPDF,"ARIAL", 8, FONT_STYLE_BOLD)        
    Call GF_writeText(oPDF, Xtxt, Ytxt + (hfila*3), getDsClienteByCUIT(CUIT_TOEPFER), 0)
    Call GF_writeText(oPDF, Xtxt, Ytxt + (hfila*4), GF_STR2CUIT(CUIT_TOEPFER), 0)
    Call GF_writeText(oPDF, Xtxt, Ytxt + (hfila*5), GF_FN2DTE(fecha), 0)
    if (Ucase(pto) = TERMINAL_ARROYO) then
        Call GF_writeText(oPDF, Xtxt, Ytxt + (hfila*7), "Ruta 21 Km. 277", 0)
    else
        Call GF_writeText(oPDF, Xtxt, Ytxt + (hfila*7), "Alem esq. América S/N", 0)
    end if        
    Call GF_writeTextAlign(oPDF, Xo, Yo + (hfila*54), "Pagina " & pPagina, wtabla, PDF_ALIGN_CENTER)

End Function
'------------------------------------------------------------------------------------------
'Funcion responsable por armar la parte variable de la tabla con la información de los cupos asignados.
Function armarDatosCupos(fecha, pCuitCliente, pCorredor, pVendedor, pto)
    Dim rs, strSQL, cuposTotales, cuposProveedor
    Dim Xtxt, Ytxt, receptorDs    
    
    'Dibujo     
    Call GF_setFont(oPDF,"ARIAL", 8, FONT_STYLE_NORMAL)    
    Ytxt = Yo + (hfila*11)    
            
    'Dibujo la cabecera de la numeración
    Xtxt = Xo + 2
    Call GF_squareBox(oPDF, Xo, Ytxt, wtabla, hfila, 0, "#FFFFFF", "#000000", 1, PDF_SQUARE_NORMAL)
    Call GF_drawLine(oPDF, Xo + (wtabla*3/4), Ytxt, Xo + (wtabla*3/4), Ytxt+hfila)        
    receptorDs = getDsClienteByCUIT(pCuitCliente)    
    if (CDbl(pCuitCliente) = CDbl(CUIT_TOEPFER)) then 
        if (CLng(pCorredor) > 0) then
            receptorDs = getDsCorredor(pCorredor)            
        else
            receptorDs = getDsVendedor(pVendedor)            
        end if
    end if        
    Call GF_writeText(oPDF, Xtxt, Ytxt+2 , receptorDs, 0)    
    Call GF_setFont(oPDF,"ARIAL", 10, FONT_STYLE_BOLD)
    Call GF_squareBox(oPDF, Xo, (Ytxt+hfila), wtabla, hfila, 0, "#FFFFFF", "#000000", 1, PDF_SQUARE_NORMAL)        
    Call GF_writeText(oPDF, Xtxt, (Ytxt+hfila)+2 , "Numeración correspondiente", 0)     
    Call GF_squareBox(oPDF, Xo, (Ytxt+(hfila*2)), wtabla, (hfila*40), 0, "#FFFFFF", "#000000", 1, PDF_SQUARE_NORMAL)       
    Call GF_setFont(oPDF,"ARIAL", 8, FONT_STYLE_NORMAL)    
    
    strSQL= "Select * from CODIGOSCUPO" &_
            " where CUITCLIENTE='" & pCuitCliente & "' and FECHACUPO=" & fecha & " and ESTADO >= " & CUPO_OTORGADO
            if (CLng(pCorredor) > 0) then strSQL = strSQL & " and CDCORREDOR=" & pCorredor
            if (CLng(pVendedor) > 0) then strSQL = strSQL & " and CDVENDEDOR=" & pVendedor
    Call executeQueryDb(pto, rs, "OPEN", strSQL)    
    cuposTotales = 0     
    while (not rs.eof)
        codigos = codigos & " - " & rs("CODIGOCUPO")     
        cuposTotales = cuposTotales + 1
        rs.MoveNext()   
    wend    
    Call GF_writeTextAlign(oPDF, Xo + (wtabla*3/4), Ytxt+2, cuposTotales, (wtabla/4), PDF_ALIGN_CENTER)      
    codigos = Right(codigos, Len(codigos)-3)
    if (cuposTotales >300) then Call GF_setFont(oPDF,"ARIAL", 6, FONT_STYLE_NORMAL)    
    Call GF_writeTextPlus(oPDF, Xtxt, Ytxt+(hfila*3), codigos, wtabla, hfila, PDF_ALIGN_LEFT)
        
    armarDatosCupos = cuposTotales
End Function
'------------------------------------------------------------------------------------------	
'Obtiene el nombre del archivo a generar.
Function getFilename(fecha, puerto)
         Randomize()
         getFilename = "cuposResolucion25-" & puerto & "-" & fecha & "-" & Int(100 * Rnd()) & ".pdf"
End Function
'-----------------------------------------------------------------------------------
Function armarPDF(pFechaCupos, pPto, pTipo)
	Dim filename, strSQL, rs, myPagina, totalCupos, pag

    pathPDF = ""		    
	strSQL= "Select CUITCLIENTE, CDCORREDOR, CDVENDEDOR from " &_
	        "(Select CUITCLIENTE, case when (CUITCLIENTE = '" & CUIT_TOEPFER & "') and (CDCORREDOR not in  (0, " & SIN_CORREDOR & ")) then CDCORREDOR else 0 end CDCORREDOR,  case when (CUITCLIENTE = '" & CUIT_TOEPFER & "') and (CDCORREDOR in (0,  " & SIN_CORREDOR & ")) then CDVENDEDOR else 0 end CDVENDEDOR  from CODIGOSCUPO where FECHACUPO=" & pFechaCupos & "and ESTADO >= " & CUPO_OTORGADO & ") T" &_
	        " GROUP BY CUITCLIENTE, CDCORREDOR, CDVENDEDOR " &_ 
	        "ORDER BY CUITCLIENTE, CDCORREDOR, CDVENDEDOR"				
	Call executeQueryDb(pPto, rs, "OPEN", strSQL)		
	if (not rs.eof) then
	    filename   = getFilename(pFechaCupos, pPto)
	    if(pTipo = PDF_FILE_MODE)then
		    pathPDF = Server.MapPath("temp/" & filename)
	    else
		    pathPDF = Server.MapPath("../temp/" & filename)
	    end if
	    Set oPDF = GF_createPDF(pathPDF)
	    Call GF_setPDFMODE(pTipo)	
	    myPagina=0
	    totalCupos = 0				
	    while (not rs.eof)        
	        myPagina = myPagina + 1
	        if (myPagina > 1) then Call GF_newPage(oPDF)
	        Call armarEstructuraTabla(pFechaCupos, myPagina, pPto)	
	        totalCupos = totalCupos + armarDatosCupos(pFechaCupos,rs("CUITCLIENTE"), rs("CDCORREDOR"), rs("CDVENDEDOR"), pPto)        
	        rs.MoveNext()
        wend          
        'Se graba la cantidad total de cupos.
        Call GF_setFont(oPDF,"ARIAL", 8, FONT_STYLE_BOLD)
        For pag = 1 to myPagina
            Call setWorkPage(oPDF, pag)
            Call GF_writeText(oPDF, Xo + (wtabla/2) + 2, Yo + (hfila*6) + 2, totalCupos, 0)
        Next            
	    Call GF_closePDF(oPDF)
    end if	    
	armarPDF =pathPDF
end function
'*************************************************************************************
'***************************** COMIENZO DE LA PAGINA ***********************************
'*************************************************************************************
Dim oPDF, filename
Dim fechaCupos, pto, g_strPuerto
Dim Yo, wtabla, hfila, Xo

'Variables generales de posicionamiento y dimension de la tabla.
Xo = 50
Yo = 20
wtabla = 500
hfila = 15

Call GP_CONFIGURARMOMENTOS()
    
pto = GF_Parametros7("pto","",6)
g_strPuerto = pto
fechaCupos = GF_Parametros7("fecha","",6)

if (LCase(request.servervariables("script_name")) = "/actisaintra/poseidon/reportecuposresolucion25print.asp") then
	Call armarPDF(fechaCupos, pto, PDF_STREAM_MODE)	
end if
%>

