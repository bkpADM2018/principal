<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosPDF.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->

<%
Const P_Y_HEADER = 78
Const P_Y_BODY = 110
Const P_Y_FINALLY = 550
Const SEPARATION = 10
Const TEXT_ANLGE_HORIZONTAL = 90

'*****************************************************************************************************************************
Function armadoPDF(p_NroLote, p_FechaLote)
    Dim totalProvsionPesos, totalGastoPesos, totalCancelacionPesos, porcentaje, fechaSaldo,estado,totalInclusionPesos
    Dim totalProvsionDolares, totalGastoDolares, totalCancelacionDolares, totalInclusionDolares
    totalInclusionPesos = 0
    totalProvsionPesos = 0
    totalGastoPesos = 0
    totalCancelacionPesos = 0
    totalInclusionDolares = 0
    totalProvsionDolares = 0
    totalGastoDolares = 0
    totalCancelacionDolares = 0
    Call dibujarEncabezado()
    Set sp_ret = executeSP(rsPvs, "EJIFL.TBLPROVISIONESCANE_GET_BY_PARAMETERS", p_NroLote &"||"& p_FechaLote &"||||1||0$$totalRegistros")
    if (not rsPvs.Eof) then
        estado = rsPvs("ESTADO")
        Call dibujarCabeceraProvisiones(p_NroLote, p_FechaLote, rsPvs("FECHASALDO"),rsPvs("PORCENTAJE"),estado)
        pY = P_Y_BODY
        Call dibujarTitulosDetalleProvisiones()
        while (not rsPvs.Eof)
            if (Cint(pY) > P_Y_FINALLY) then
                Call nuevaHoja()
                Call dibujarTitulosDetalleProvisiones()        
            end if        
            Call dibujarDetalleProvisiones(rsPvs)
            totalProvsionPesos = Cdbl(totalProvsionPesos) + Cdbl(rsPvs("PROVISIONPESOS"))
            totalGastoPesos = Cdbl(totalGastoPesos) + Cdbl(rsPvs("GASTOPESOS"))
            totalCancelacionPesos = Cdbl(totalCancelacionPesos) + Cdbl(rsPvs("IMPORTEPESOS"))
            totalProvsionDolares = Cdbl(totalProvsionDolares) + Cdbl(rsPvs("PROVISIONDOLARES"))
            totalGastoDolares = Cdbl(totalGastoDolares) + Cdbl(rsPvs("GASTODOLARES"))
            totalCancelacionDolares = Cdbl(totalCancelacionDolares) + Cdbl(rsPvs("IMPORTEDOLAR"))
            if (Cstr(rsPvs("MARCAINCLUSION")) = "S") then 
                totalInclusionPesos = Cdbl(totalInclusionPesos) + Cdbl(rsPvs("IMPORTEPESOS"))
                totalInclusionDolares = Cdbl(totalInclusionDolares) + Cdbl(rsPvs("IMPORTEDOLAR"))
            end if
            rsPvs.MoveNext()
        wend
        Call dibujarTotalesProvisiones(totalProvsionPesos,totalCancelacionPesos,totalGastoPesos,totalInclusionPesos,totalProvsionDolares,totalCancelacionDolares,totalGastoDolares,totalInclusionDolares)
        if (CStr(estado) <> PROVISCIONES_ESTADO_GENERADO) then Call dibujarFirmasRegsitradas(p_NroLote, p_FechaLote)
        Call dibujarTotalNroPagina()
    end if
End Function
'-----------------------------------------------------------------------------------------------------------------------------
Function dibujarFirmasRegsitradas(p_NroLote, p_FechaLote)
    Dim rsFir
    Call executeSP(rsFir, "EJIFL.TBLPROVISIONESFIRMAS_GET_BY_PARAMETERS", p_NroLote &"||"& p_FechaLote)
    if (not rsFir.Eof) then 
        pY = pY + 10
        if (Cint(pY)+20 > P_Y_FINALLY) then Call nuevaHoja()
        Call GF_squareBox(oPDF, py, 830, 810, 15, TEXT_ANLGE_HORIZONTAL, "#2e6b4d", "#000000", 1, PDF_SQUARE_NORMAL)
        call GF_setFont(oPDF,"ARIAL",8,8)
        Call GF_setFontColor("#FFFFFF")
        pY = pY + 4
        Call GF_writeVerticalText(oPDF,pY, 830, GF_TRADUCIR("AUTORIZACIONES") , 810, PDF_ALIGN_CENTER)
        Call GF_setFontColor("#000000")
        Call GF_setFont(oPDF,"COURIER",8,0)
        pY = pY + 11
        Call GF_squareBox(oPDF, pY, 830, 810, 33, TEXT_ANLGE_HORIZONTAL, "#FFFFFF", "#000000", 1, PDF_SQUARE_NORMAL)
        pY = pY + 3
        while (not rsFir.Eof)
            Call GF_writeVerticalText(oPDF,pY, 825, rsFir("CDUSUARIO") & " - " & getUserDescription(rsFir("CDUSUARIO")) & " - " & armarTextoPlanoFirma(rsFir("HKEY"), rsFir("FECHAFIRMA")), 810, PDF_ALIGN_LEFT)
            pY = pY + SEPARATION
            rsFir.MoveNext()
        wend 
    end if
End Function 
'-----------------------------------------------------------------------------------------------------------------------------
Function dibujarTotalesProvisiones(p_totalProvsionPesos,p_totalCancelacionPesos,p_totalGastoPesos,p_totalInclusionPesos,p_totalProvsionDolares,p_totalCancelacionDolares,p_totalGastoDolares,p_totalInclusionDolares)
    if (Cint(pY) + 10 > P_Y_FINALLY) then Call nuevaHoja()
    'Call GF_squareBox(oPDF, pY, 830, 810, 15, TEXT_ANLGE_HORIZONTAL, "#2e6b4d", "#000000", 1, PDF_SQUARE_NORMAL)
    Call GF_squareBox(oPDF, pY, 830, 210, 15, TEXT_ANLGE_HORIZONTAL, "#2e6b4d", "#000000", 1, PDF_SQUARE_NORMAL)
    Call GF_squareBox(oPDF, pY, 620, 75, 15, TEXT_ANLGE_HORIZONTAL, "#2e6b4d", "#000000", 1, PDF_SQUARE_NORMAL)
    Call GF_squareBox(oPDF, pY, 545, 75, 15, TEXT_ANLGE_HORIZONTAL, "#2e6b4d", "#000000", 1, PDF_SQUARE_NORMAL)
    Call GF_squareBox(oPDF, pY, 470, 75, 15, TEXT_ANLGE_HORIZONTAL, "#2e6b4d", "#000000", 1, PDF_SQUARE_NORMAL)
    Call GF_squareBox(oPDF, pY, 395, 75, 15, TEXT_ANLGE_HORIZONTAL, "#2e6b4d", "#000000", 1, PDF_SQUARE_NORMAL)
    Call GF_squareBox(oPDF, pY, 320, 75, 15, TEXT_ANLGE_HORIZONTAL, "#2e6b4d", "#000000", 1, PDF_SQUARE_NORMAL)
    Call GF_squareBox(oPDF, pY, 245, 75, 15, TEXT_ANLGE_HORIZONTAL, "#2e6b4d", "#000000", 1, PDF_SQUARE_NORMAL)
    Call GF_squareBox(oPDF, pY, 170, 150, 15, TEXT_ANLGE_HORIZONTAL, "#2e6b4d", "#000000", 1, PDF_SQUARE_NORMAL)
    pY = pY + 3
    Call GF_setFont(oPDF,"ARIAL",8,8)
    Call GF_setFontColor("#FFFFFF")
    Call GF_writeVerticalText(oPDF, pY, 621, TIPO_MONEDA_DOLAR &" "& GF_EDIT_DECIMALS(Cdbl(p_totalProvsionDolares)*100,2), 75, PDF_ALIGN_RIGHT)
    Call GF_writeVerticalText(oPDF, pY, 546, TIPO_MONEDA_DOLAR &" "& GF_EDIT_DECIMALS(Cdbl(p_totalGastoDolares)*100,2), 75, PDF_ALIGN_RIGHT)
    Call GF_writeVerticalText(oPDF, pY, 471, TIPO_MONEDA_DOLAR &" "& GF_EDIT_DECIMALS(Cdbl(p_totalCancelacionDolares)*100,2), 75, PDF_ALIGN_RIGHT)
    Call GF_writeVerticalText(oPDF, pY, 396, TIPO_MONEDA_PESO &" "& GF_EDIT_DECIMALS(Cdbl(p_totalProvsionPesos)*100,2), 75, PDF_ALIGN_RIGHT)
    Call GF_writeVerticalText(oPDF, pY, 321, TIPO_MONEDA_PESO &" "& GF_EDIT_DECIMALS(Cdbl(p_totalGastoPesos)*100,2), 75, PDF_ALIGN_RIGHT)
    Call GF_writeVerticalText(oPDF, pY, 246, TIPO_MONEDA_PESO &" "& GF_EDIT_DECIMALS(Cdbl(p_totalCancelacionPesos)*100,2), 75, PDF_ALIGN_RIGHT)
    Call GF_setFontColor("#000000")
    pY = pY + 12
    'Call GF_squareBox(oPDF, pY, 830, 810, 15, TEXT_ANLGE_HORIZONTAL, "#2e6b4d", "#000000", 1, PDF_SQUARE_NORMAL)
    Call GF_squareBox(oPDF, pY, 830, 210, 15, TEXT_ANLGE_HORIZONTAL, "#2e6b4d", "#000000", 1, PDF_SQUARE_NORMAL)
    Call GF_squareBox(oPDF, pY, 620, 225, 15, TEXT_ANLGE_HORIZONTAL, "#2e6b4d", "#000000", 1, PDF_SQUARE_NORMAL)
    Call GF_squareBox(oPDF, pY, 395, 225, 15, TEXT_ANLGE_HORIZONTAL, "#2e6b4d", "#000000", 1, PDF_SQUARE_NORMAL)
    Call GF_squareBox(oPDF, pY, 170, 150, 15, TEXT_ANLGE_HORIZONTAL, "#2e6b4d", "#000000", 1, PDF_SQUARE_NORMAL)
    pY = pY + 3
    Call GF_setFont(oPDF,"ARIAL",8,8)
    Call GF_setFontColor("#FFFFFF")
    Call GF_writeVerticalText(oPDF,pY, 620,  GF_TRADUCIR("TOTAL INCLUSIÓN: ") & getSimboloMoneda(MONEDA_DOLAR) & GF_EDIT_DECIMALS(Cdbl(p_totalInclusionDolares)*100,2) , 225, PDF_ALIGN_CENTER)
    Call GF_writeVerticalText(oPDF,pY, 395,  GF_TRADUCIR("TOTAL INCLUSIÓN: ") & getSimboloMoneda(MONEDA_PESO) & GF_EDIT_DECIMALS(Cdbl(p_totalInclusionPesos)*100,2) , 225, PDF_ALIGN_CENTER)
    Call GF_setFontColor("#000000")
    pY = pY + 12
End Function
'-----------------------------------------------------------------------------------------------------------------------------
Function dibujarTitulosDetalleProvisiones()
    Call GF_squareBox(oPDF, py, 830, 40, 30, TEXT_ANLGE_HORIZONTAL, "#2e6b4d", "#000000", 1, PDF_SQUARE_NORMAL)
    Call GF_squareBox(oPDF, py, 790, 60, 30, TEXT_ANLGE_HORIZONTAL, "#2e6b4d", "#000000", 1, PDF_SQUARE_NORMAL)
    Call GF_squareBox(oPDF, py, 730, 110, 30, TEXT_ANLGE_HORIZONTAL, "#2e6b4d", "#000000", 1, PDF_SQUARE_NORMAL)
    Call GF_squareBox(oPDF, py , 620, 225, 15, TEXT_ANLGE_HORIZONTAL, "#2e6b4d", "#000000", 1, PDF_SQUARE_NORMAL)
    Call GF_squareBox(oPDF, py + 15, 620, 75, 15, TEXT_ANLGE_HORIZONTAL, "#2e6b4d", "#000000", 1, PDF_SQUARE_NORMAL)
    Call GF_squareBox(oPDF, py + 15, 545, 75, 15, TEXT_ANLGE_HORIZONTAL, "#2e6b4d", "#000000", 1, PDF_SQUARE_NORMAL)
    Call GF_squareBox(oPDF, py + 15, 470, 75, 15, TEXT_ANLGE_HORIZONTAL, "#2e6b4d", "#000000", 1, PDF_SQUARE_NORMAL)
    Call GF_squareBox(oPDF, py , 395, 225, 15, TEXT_ANLGE_HORIZONTAL, "#2e6b4d", "#000000", 1, PDF_SQUARE_NORMAL)
    Call GF_squareBox(oPDF, py + 15, 395, 75, 15, TEXT_ANLGE_HORIZONTAL, "#2e6b4d", "#000000", 1, PDF_SQUARE_NORMAL)
    Call GF_squareBox(oPDF, py + 15, 320, 75, 15, TEXT_ANLGE_HORIZONTAL, "#2e6b4d", "#000000", 1, PDF_SQUARE_NORMAL)
    Call GF_squareBox(oPDF, py + 15, 245, 75, 15, TEXT_ANLGE_HORIZONTAL, "#2e6b4d", "#000000", 1, PDF_SQUARE_NORMAL)
    Call GF_squareBox(oPDF, py , 170, 50, 30, TEXT_ANLGE_HORIZONTAL, "#2e6b4d", "#000000", 1, PDF_SQUARE_NORMAL)
    Call GF_squareBox(oPDF, py , 120, 50, 30, TEXT_ANLGE_HORIZONTAL, "#2e6b4d", "#000000", 1, PDF_SQUARE_NORMAL)
    Call GF_squareBox(oPDF, py , 70, 50, 30, TEXT_ANLGE_HORIZONTAL, "#2e6b4d", "#000000", 1, PDF_SQUARE_NORMAL)
    pY = pY + 10
    call GF_setFont(oPDF,"ARIAL",8,8)
    Call GF_setFontColor("#FFFFFF")
    Call GF_writeText(oPDF, pY, 825,  GF_TRADUCIR("BUQUE") , TEXT_ANLGE_HORIZONTAL)
    Call GF_writeText(oPDF, pY, 786,  GF_TRADUCIR("NOMINACIÓN") , TEXT_ANLGE_HORIZONTAL)
    Call GF_writeText(oPDF, pY, 698,  GF_TRADUCIR("CONCEPTO") , TEXT_ANLGE_HORIZONTAL)
    Call GF_writeText(oPDF, pY - 7, 525,  GF_TRADUCIR("DOLARES") , TEXT_ANLGE_HORIZONTAL)
    Call GF_writeText(oPDF, pY + 8, 603,  GF_TRADUCIR("PROVISIÓN") , TEXT_ANLGE_HORIZONTAL)
    Call GF_writeText(oPDF, pY + 8, 520,  GF_TRADUCIR("GASTO") , TEXT_ANLGE_HORIZONTAL)
    Call GF_writeText(oPDF, pY + 8, 461,  GF_TRADUCIR("CANCELACIÓN") , TEXT_ANLGE_HORIZONTAL)
    Call GF_writeText(oPDF, pY - 7, 300,  GF_TRADUCIR("PESOS") , TEXT_ANLGE_HORIZONTAL)
    Call GF_writeText(oPDF, pY + 8, 378,  GF_TRADUCIR("PROVISIÓN") , TEXT_ANLGE_HORIZONTAL)
    Call GF_writeText(oPDF, pY + 8, 299,  GF_TRADUCIR("GASTO") , TEXT_ANLGE_HORIZONTAL)
    Call GF_writeText(oPDF, pY + 8, 236,  GF_TRADUCIR("CANCELACIÓN") , TEXT_ANLGE_HORIZONTAL)
    Call GF_writeText(oPDF, pY, 162,  GF_TRADUCIR("MONEDA") , TEXT_ANLGE_HORIZONTAL)
    Call GF_writeText(oPDF, pY - 7, 105,  GF_TRADUCIR("TIPO") , TEXT_ANLGE_HORIZONTAL)
    Call GF_writeText(oPDF, pY + 8, 110,  GF_TRADUCIR("CAMBIO") , TEXT_ANLGE_HORIZONTAL)
    Call GF_writeText(oPDF, pY, 66,  GF_TRADUCIR("INCLUSIÓN") , TEXT_ANLGE_HORIZONTAL)
    Call GF_setFontColor("#000000")
    pY = pY + 22
End Function
'-----------------------------------------------------------------------------------------------------------------------------
Function dibujarDetalleProvisiones(p_RsPvs)
    Dim dsConcepto
    call GF_setFont(oPDF,"COURIER",8,0)
    Call GF_writeVerticalText(oPDF, pY, 830,p_RsPvs("BUQUE"),40,PDF_ALIGN_CENTER)
    Call GF_writeVerticalText(oPDF, pY, 790,p_RsPvs("NOMINACION"),60,PDF_ALIGN_CENTER)
    dsConcepto = p_RsPvs("CONCEPTO") &"-"& Trim(p_RsPvs("MGDES"))
    if (Len(dsConcepto) > 23) then dsConcepto = Left(dsConcepto,21) & ".."
    Call GF_writeVerticalText(oPDF, pY, 730, dsConcepto,110,PDF_ALIGN_LEFT)
    Call GF_writeVerticalText(oPDF, pY, 620, GF_EDIT_DECIMALS(Cdbl(p_RsPvs("PROVISIONDOLARES"))*100,2),75,PDF_ALIGN_RIGHT)
    Call GF_writeVerticalText(oPDF, pY, 545, GF_EDIT_DECIMALS(Cdbl(p_RsPvs("GASTODOLARES"))*100,2) , 75, PDF_ALIGN_RIGHT)
    Call GF_writeVerticalText(oPDF, pY, 470, GF_EDIT_DECIMALS(Cdbl(p_RsPvs("IMPORTEDOLAR"))*100,2), 75, PDF_ALIGN_RIGHT)
    Call GF_writeVerticalText(oPDF, pY, 395, GF_EDIT_DECIMALS(Cdbl(p_RsPvs("PROVISIONPESOS"))*100,2),75,PDF_ALIGN_RIGHT)
    Call GF_writeVerticalText(oPDF, pY, 320, GF_EDIT_DECIMALS(Cdbl(p_RsPvs("GASTOPESOS"))*100,2) , 75, PDF_ALIGN_RIGHT)
    Call GF_writeVerticalText(oPDF, pY, 245, GF_EDIT_DECIMALS(Cdbl(p_RsPvs("IMPORTEPESOS"))*100,2), 75, PDF_ALIGN_RIGHT)
    Call GF_writeVerticalText(oPDF, pY, 170, getSimboloMonedaLetras(Cdbl(p_RsPvs("MONEDA"))) , 50, PDF_ALIGN_CENTER)
    Call GF_writeVerticalText(oPDF, pY, 120, GF_EDIT_DECIMALS(Cdbl(p_RsPvs("TIPOCAMBIO"))*10000,4) , 50, PDF_ALIGN_RIGHT)
    Call GF_writeVerticalText(oPDF, pY, 70, p_RsPvs("MARCAINCLUSION") , 50, PDF_ALIGN_CENTER)
    pY = pY + SEPARATION
End Function
'-----------------------------------------------------------------------------------------------------------------------------
Function dibujarCabeceraProvisiones(p_NroLote, p_FechaLote, p_FechaSaldo, p_Porcentaje, p_Estado)
    call GF_setFont(oPDF,"COURIER",8,0)
    Call GF_writeText(oPDF,pY, 830, GF_TRADUCIR("Nro.Lote: ") & p_NroLote, TEXT_ANLGE_HORIZONTAL)
    Call GF_writeText(oPDF,pY, 510, GF_TRADUCIR("Fecha Lote: ") & GF_FN2DTE(p_FechaLote), TEXT_ANLGE_HORIZONTAL)
    Call GF_writeText(oPDF,pY, 130, GF_TRADUCIR("Porcentaje: ") & p_Porcentaje & "%", TEXT_ANLGE_HORIZONTAL)
    pY = pY + 12
    Call GF_writeText(oPDF,pY, 830, GF_TRADUCIR("Fecha Saldo: ") & GF_FN2DTE(p_FechaSaldo), TEXT_ANLGE_HORIZONTAL)
    Call GF_writeText(oPDF,pY, 510, GF_TRADUCIR("Estado: ") & getEstadoProvisionesCancelacion(p_Estado), TEXT_ANLGE_HORIZONTAL)
End function
'-----------------------------------------------------------------------------------------------------------------------------
Function dibujarTotalNroPagina()
    for i = 1 to nroPagina
        Call setWorkPage(oPDF, i)
        call GF_setFont(oPDF,"ARIAL",8,0)
	    Call GF_writeVerticalText(oPDF, 582 , 30, " de " & nroPagina , 50 , PDF_ALIGN_LEFT)
    next
End Function 
'-----------------------------------------------------------------------------------------------------------------------------
Function dibujarEncabezado()
	Call GF_squareBox(oPDF, 8, 8, 570, 833, 0, "", "#0B3B0B", 2, PDF_SQUARE_ROUND)
	Call GF_writeImage(oPDF, Server.MapPath("images\kogge64.gif"), 15, 834, 48, 48, TEXT_ANLGE_HORIZONTAL)
	Call GF_setFont(oPDF,"ARIAL",16,8)
	Call GF_writeText(oPDF,28, 500, GF_TRADUCIR("PROVISIONES"), TEXT_ANLGE_HORIZONTAL)
	Call GF_verticalLine(oPDF,70,10,830)
	Call GF_setFont(oPDF,"ARIAL",8,0)
	Call GF_writeText(oPDF, 582 , 66, "Página  " & nroPagina, TEXT_ANLGE_HORIZONTAL)
	Call GF_setFont(oPDF,"COURIER",8,0)
	GP_CONFIGURARMOMENTOS
	Call GF_writeText(oPDF,12,110,GF_FN2DTE(session("MmtoSistema")), TEXT_ANLGE_HORIZONTAL)
	Call GF_writeText(oPDF,22,30,session("Usuario"), TEXT_ANLGE_HORIZONTAL)
    pY = P_Y_HEADER
end Function
'-----------------------------------------------------------------------------------------------------------------------------
Function nuevaHoja()
    Call GF_newPage(oPDF)
    Call PDFGirarHoja(TEXT_ANLGE_HORIZONTAL)
	nroPagina = nroPagina + 1
	Call dibujarEncabezado()
End function
'****************************************************************************************************************************
'********************************	             COMIENZO DE LA PAGINA              ********************************
'***********************************************************************************************************************************
Dim nroLote,fechaLote,oPDF,rsPvs,nroPagina,pY

nroLote = GF_Parametros7("nroLote",0,6)
fechLote = GF_Parametros7("fechaLote",0,6)    

nroPagina = 1
Set oPDF = GF_createPDF("PDFTemp")
Call PDFGirarHoja(TEXT_ANLGE_HORIZONTAL)
Call GF_setPDFMODE(PDF_STREAM_MODE)
call armadoPDF(nroLote, fechLote)
Call GF_closePDF(oPDF)

'***************************************************************************************************************************************


%>
