<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientosFechas.asp"-->
<!--#include file="../Includes/procedimientosPDF.asp"-->
<%

'-----------------------------------------------------------------------------------
function getDsPuerto(p_Port)
    if (session("DS_PUERTO_" & p_Port) = "") then
        strSQL = "select DESCDE from MERFL.MER192F1 where CODIDE='" & p_Port & "'"
        call executeQuery(rs, "OPEN", strSQL)
        if not rs.eof then
            session("DS_PUERTO_" & p_Port) = trim(rs("DESCDE"))
            getDsPuerto = trim(rs("DESCDE"))
        else
            getDsPuerto = "#Puerto no valido#"
        end if
    else
         getDsPuerto = session("DS_PUERTO_" & p_Port)
    end if        
end function
'-----------------------------------------------------------------------------------
Function armarCabecera(pMiCto, pSuCto, pPuerto, pPlanta, pDsProveedor, pDirProveedor, pLocProveedor, pCdProducto)
    Dim Ytxt

    Ytxt = Yo      
    Call GF_writeImage(oPDF, Server.MapPath("..\images\Logo.gif"), 50, 50, 126, 30, 0)
    Call GF_setFont(oPDF, "COURIER", 10, FONT_STYLE_NORMAL)  
    Call GF_writeText(oPDF, 450, Ytxt, GF_FN2DTE(Left(session("MmtoDato"), 8)), 0)    
    Ytxt = Ytxt + 5*hfila
    Call GF_writeText(oPDF,  50, Ytxt, "Sres. : ", 0)
    Call GF_setFont(oPDF, "COURIER", 10, FONT_STYLE_BOLD)      
    Call GF_writeText(oPDF, 100, Ytxt, pDsProveedor, 0) 
    Ytxt = Ytxt + 2*hfila
    Call GF_writeText(oPDF, 100, Ytxt, pDirProveedor, 0)    
    Ytxt = Ytxt + 2*hfila
    Call GF_writeText(oPDF, 100, Ytxt, pLocProveedor, 0)    
    Call GF_setFont(oPDF, "COURIER", 14, FONT_STYLE_BOLD)       
    Ytxt = Ytxt + 3.5*hfila
    Call GF_writeText(oPDF, 170, Ytxt, "CARTA CUPOS", 0)
    Call GF_setFont(oPDF, "COURIER", 10, FONT_STYLE_NORMAL)       
    Ytxt = Ytxt + 0.5*hfila
    Call GF_writeText(oPDF, 100, Ytxt, "Referencia:", 0)    
    Call GF_setFont(oPDF, "COURIER", 10, FONT_STYLE_NORMAL)       
    Ytxt = Ytxt + 3*hfila
    Call GF_writeTextPlus(oPDF, 100, Ytxt, "A efectos de cumplir las operaciones pendientes de entrega que más abajo detallamos, otorgamos los siguientes cupos de camiones.", 320, 14, PDF_ALIGN_JUSTIFY)        
    Ytxt = Ytxt + 5*hfila
    Call GF_writeText(oPDF, 100, Ytxt, "CONTRATO:", 0)
    Call GF_setFont(oPDF, "COURIER", 10, FONT_STYLE_BOLD)
    Call GF_writeText(oPDF, 160, Ytxt, pMiCto, 0)    
    Call GF_setFont(oPDF, "COURIER", 10, FONT_STYLE_NORMAL)       
    Ytxt = Ytxt + 2*hfila
    Call GF_writeText(oPDF, 100, Ytxt, "PUERTO  :", 0)
    Call GF_setFont(oPDF, "COURIER", 10, FONT_STYLE_BOLD)
    Call GF_writeText(oPDF, 160, Ytxt, getDsPuerto(pPuerto), 0)    
    Call GF_setFont(oPDF, "COURIER", 10, FONT_STYLE_NORMAL)
    Ytxt = Ytxt + 2*hfila    
    Call GF_writeText(oPDF, 100, Ytxt, "PLANTA  :", 0)
    Call GF_setFont(oPDF, "COURIER", 10, FONT_STYLE_BOLD)
    Call GF_writeText(oPDF, 160, Ytxt, getDsPuerto(pPlanta), 0)    
    Call GF_setFont(oPDF, "COURIER", 10, FONT_STYLE_NORMAL)    
    armarCabecera = Ytxt
    
End Function
'-----------------------------------------------------------------------------------
Function armarLineaTitulo(pY)
    Dim Ytxt, Xtxt
    Call GF_setFont(oPDF, "COURIER", 10, FONT_STYLE_BOLD)
    Ytxt = pY + 5*hfila    
    Call GF_squareBox(oPDF, Xo, Ytxt, wtabla, hfila*2, 0, "#FFFFFF", "#000000", 1, PDF_SQUARE_NORMAL)
    Call GF_drawLine(oPDF, Xo + (wtabla*20/100), Ytxt, Xo + (wtabla*20/100), Ytxt+(2*hfila))
    Call GF_drawLine(oPDF, Xo + (wtabla*70/100), Ytxt, Xo + (wtabla*70/100), Ytxt+(2*hfila))
    Call GF_drawLine(oPDF, Xo + (wtabla*85/100), Ytxt, Xo + (wtabla*85/100), Ytxt+(2*hfila))
    Call GF_writeTextAlign(oPDF, Xo                , Ytxt+3,  "Su Nro.", (wtabla*20/100), PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(oPDF, Xo + (wtabla*20/100), Ytxt+3, "Vendedor", (wtabla*50/100), PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(oPDF, Xo + (wtabla*70/100), Ytxt+3,  "Entrega", (wtabla*15/100), PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(oPDF, Xo + (wtabla*85/100), Ytxt+3, "Camiones", (wtabla*15/100), PDF_ALIGN_CENTER)
    armarLineaTitulo = Ytxt
End Function
'-----------------------------------------------------------------------------------
Function armarLineaCuerpo(pY, col1, col2, col3, col4)
    Dim Ytxt, Xtxt
    Call GF_setFont(oPDF, "COURIER", 10, FONT_STYLE_NORMAL)
    Ytxt = pY + 2*hfila    
    Call GF_squareBox(oPDF, Xo, Ytxt, wtabla, hfila*2, 0, "#FFFFFF", "#000000", 1, PDF_SQUARE_NORMAL)
    Call GF_writeTextAlign(oPDF, Xo                , Ytxt+3, Trim(col1), (wtabla*20/100), PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(oPDF, Xo + (wtabla*20/100)+2, Ytxt+3, Left(col2, 37), (wtabla*50/100), PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(oPDF, Xo + (wtabla*70/100), Ytxt+3, col3, (wtabla*15/100), PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(oPDF, Xo + (wtabla*85/100), Ytxt+3, col4, (wtabla*15/100), PDF_ALIGN_CENTER)
    armarLineaCuerpo = Ytxt
End Function
'-----------------------------------------------------------------------------------
Function armarPie(pY)    
    Dim Ytxt, Xtxt
    Call GF_setFont(oPDF, "COURIER", 10, FONT_STYLE_NORMAL)
    Ytxt = pY + 4*hfila    
    Call GF_writeTextPlus(oPDF, 50, Ytxt, "En Razón de los reiterados pedidos de entrega de esta mercadería sin respuesta de vuestra parte, intimamos estricto cumplimiento de estos nuevos cupos otorgados bajo apercibimiento de recurrir ante la cámara arbitral respectiva a iniciar la correspondiente demandade acuerdo a las reglas de uso y costumbres que rigen el comercio de granos y reservándose el derecho de reclamar lo que nos corresponda de acuerdo a condiciones cámara contractuales convenidas además de los mayores gastos, daños, perjuicios y todos los gastos por demora de embarques que se produzcan a consecuencia de vuestro incumplimiento.", 350, 14, PDF_ALIGN_JUSTIFY)    
    Call GF_writeImage(oPDF, Server.MapPath("..\images\firmas\LPE.jpg"), 300, Ytxt + 20*hfila, 200, 75, 0)
    Call GF_setFont(oPDF, "COURIER", 8, FONT_STYLE_BOLD)    
    Ytxt = Ytxt + 28*hfila    
    Call GF_writeTextAlign(oPDF, 300, Ytxt, "Leonel Pecci", 200, PDF_ALIGN_CENTER)    
    Ytxt = Ytxt + hfila    
    Call GF_writeTextAlign(oPDF, 300, Ytxt, "Dto. de Logística", 200, PDF_ALIGN_CENTER)       
    Ytxt = Ytxt + hfila    
    Call GF_writeTextAlign(oPDF, 300, Ytxt, "Alfred C. Toepfer Intl. Arg. S.R.L.", 200, PDF_ALIGN_CENTER)    
    
    
End Function
'-----------------------------------------------------------------------------------
Function verificarCorte(pRs, pro, suc, ope, nro, cos, pto, pla)
    verificarCorte = false
    if (not pRs.eof) then
        if ((CInt(pRs("CUCPRO")) = CInt(pro)) and (CInt(pRs("CUCSUC")) = CInt(suc)) and (CInt(pRs("CUCOPE")) = CInt(ope)) and (CLng(pRs("CUNCTO")) = CLng(nro)) and (CInt(pRs("CUACOS")) = CInt(cos)) and (CInt(pRs("PLANTA")) = CInt(pla)) and (CInt(pRs("PUERTO")) = CInt(pto))) then
            verificarCorte = true
        end if
    end if
End Function
'------------------------------------------------------------------------------------------	
'Obtiene el nombre del archivo a generar.
Function getFilename(pCto)
         getFilename = "CARTA_CUPOS-" & pCto & ".pdf"
End Function
'-----------------------------------------------------------------------------------
Function armarPDF(pArrFecha, pArrCupos, pPro, pSuc, pOpe, pNro, pCos, pPlanta, pTipo)
	Dim filename, strSQL, rs, myPagina, dsProveedor, dirProveedor, locProveedor, rsPro, Ytxt, idProveedor
	Dim myCto, x

    pathPDF = ""		    
	strSQL="Select CONCR1 CTOCORREDOR, CCORR1, dsCorredor, CVENR1, dsVendedor, CDESR1 PUERTO " &_
            " from MERFL.MER311F1 " &_
            " inner join (Select NROPRO, CASE when NOMAMP = '' then RAZSOC else NOMAMP end dsCorredor from MERFL.TCB6A1F1) COR on COR.NROPRO= CCORR1 " &_
            " inner join (Select NROPRO, CASE when NOMAMP = '' then RAZSOC else NOMAMP end dsVendedor from MERFL.TCB6A1F1) VEN on VEN.NROPRO= CVENR1 " &_
            " where CPROR1=" & pPro & " and CSUCR1=" & pSuc & " and COPER1=" & pOpe & " and NCTOR1=" & pNro & " and ACOSR1=" & pCos
	Call executeQuery(rs, "OPEN", strSQL)	
	if (not rs.eof) then
	    myCto = GF_EDIT_CONTRATO(pPro, pSuc, pOpe, pNro, pCos)
	    idProveedor = rs("CCORR1")
	    if (CLng(idProveedor) = 0) then idProveedor = rs("CVENR1")	    
	    strSQL="Select DOMICI, LOCALI from MERFL.TCB6A1F1 where NROPRO=" & idProveedor
	    Call executeQuery(rsPro, "OPEN", strSQL)	
	    if (not rsPro.eof) then
	        dsProveedor = rs("dsCorredor")
	        dirProveedor = rsPro("DOMICI")
	        locProveedor = rsPro("LOCALI") 
        end if    
	    filename   = getFilename(pPro & "_" & pSuc & "_" & pOpe & "_" & pNro & "_" & pCos & "_" & Trim(dsProveedor))
	    pathPDF = Server.MapPath("temp/" & filename)
	    Set oPDF = GF_createPDF(pathPDF)
	    Call GF_setPDFMODE(pTipo)	
	    myPagina=0	    
	    'while (not rs.eof)        
	    '    myPagina = myPagina + 1
	    '    if (myPagina > 1) then Call GF_newPage(oPDF)
	    '    proOld = rs("CUCPRO")
	    '    sucOld = rs("CUCSUC")
	    '    opeOld = rs("CUCOPE")
	    '    nroOld = rs("CUNCTO")
	    '    cosOld = rs("CUACOS")
	    '    ptoOld = rs("PUERTO")
	    '    plaOld = rs("PLANTA")	        
	        Ytxt = armarCabecera(myCto, rs("CTOCORREDOR"), rs("PUERTO"), g_strPuerto, dsProveedor, dirProveedor, locProveedor, pPro)
	        'Se dibuja el cuerpo.
	        Ytxt = armarLineaTitulo(Ytxt)
	    '    while (verificarCorte(rs, proOld, sucOld, opeOld, nroOld, cosOld, ptoOld, plaOld))       
	        For x = LBound(pArrFecha) to UBound(pArrFecha)	        
	            if (CLng(pArrCupos(x)) > 0) then 
	                Ytxt = armarLineaCuerpo(Ytxt, rs("CTOCORREDOR"), rs("dsVendedor"), GF_FN2DTE(pArrFecha(x)), pArrCupos(x))
                end if	                
	        Next
	    '        rs.MoveNext()
	    '    wend	        
	        Call armarPie(Ytxt)
        'wend         
	    Call GF_closePDF(oPDF)
    end if	    
	armarPDF =pathPDF
end function
'*************************************************************************************
'***************************** COMIENZO DE LA PAGINA ***********************************
'*************************************************************************************
Dim oPDF, filename
Dim fechaDesde, fechaHasta, pto, g_strPuerto, idProveedor
Dim Yo, wtabla, hfila, Xo

'Variables generales de posicionamiento y dimension de la tabla.
Xo = 50
Yo = 75
wtabla = 475
hfila = 8

Call GP_CONFIGURARMOMENTOS()
g_strPuerto = GF_PARAMETROS7("pto","",6)
    
if (LCase(request.servervariables("script_name")) = "/actisaintra/poseidon/cartacuposprint.asp") then
    idProveedor = GF_Parametros7("pro","",6)
    fechaDesde = GF_Parametros7("fd","",6)
    fechaHasta = GF_Parametros7("fh","",6)

    'Call armarPDF(fechaDesde, fechaHasta, idProveedor,  PDF_STREAM_MODE)
end if
%>

