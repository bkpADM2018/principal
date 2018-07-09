<%
'---------------------------------------------------------------------------------------------------------------------
' Mod: 2017-09-25 - JAS
Function crearPDF_Local(pNroReg, ByRef pOpdf)
	Dim nroFac,rsDet,totalPages,currPage, letraFAC,rsFac,lineasTotales,puntoVenta,auxIdioma, paramsLeyenda()
	Dim flagProforma, myConcepto, myTasaIVA, bodyText, idx, totalLineas
	
	Set rsFac = getRSFactura(pNroReg)
	puntoVenta = GF_nDigits(rsFac("succbt"),4)
	nroFac = puntoVenta &"-"&	GF_nDigits(rsFac("nrocbt"),8)
	'Tomo la letra del comprobante
	letraFAC = rsFAC("letra")		
	flagProforma = false	
	Set rsDet = getRSFacturaDetalle(pNroReg)
	'Limpio las observaciones
	g_Observaciones = ""	
	if (not rsDet.eof) then myConcepto = rsDet("codconce")
	'Arma la lsita con las lineas de la factura. Seteo la fuente a la misma que se va a utilizar para la impresión.
	'Esto es importante ya que se utiliza para determinar el tamaño de las letras y su distribucion en el espacio disponible
	Call GF_setFont(pOpdf,"COURIER", 9 , FONT_STYLE_NORMAL)
	bodyText = generateBodyText(pOpdf, rsDet, totalLineas, 440)
	totalPages = Ceil(totalLineas/MAX_LINEAS_PAGINA)
	currPage= 1		
	idx = 0	
	while (idx <= totalLineas)
	    if (currPage <> 1 ) then Call GF_newPage(pOpdf)		
	    Call dibujarCabecera(rsFac("codcia"), nroFac, letraFAC, rsFac("feccbt_fn"), rsFac("tipcbt"),rsFac("succbt"),pOpdf)						   
	    Call dibujarOrigenDestinoLocal(rsFac("cliente"),pOpdf)	    
	    Call dibujarDatosCompraLocal(rsFac("guid"),pOpdf)	    
	    Call dibujarDetalleTitulosLocal(pOpdf, flagProforma)	    	    
		Call dibujarDetalleContenidoSTD(bodyText, idx, MAX_LINEAS_PAGINA, totalLineas, pOpdf)				
		Call dibujarPieSTD(rsFac("cai"), rsFac("vencai_fn"),pOpdf,currPage,totalPages)
	    currPage= currPage + 1		
    wend
    'Se imprimen los totales.    
	myTasaIVA = getTasaIVA(rsFac("codcia"), myConcepto)
    Call dibujarTotalesLocal(rsFac, myTasaIVA, pOpdf)
    Call dibujarPieUltimaPaginaSTD(rsFac("cai"), rsFac("vencai_fn"),pOpdf,currPage, totalPages,letraFAC,auxIdioma, puntoVenta, rsFac, myConcepto)    
End Function

'--------------------------------------------------------------------------------------------------------------------
' Mod: 2017-09-25 - JAS
Function getTasaIVA(pCia, pConcepto)
	Dim strSQL, rs, ret
	
	strSQL="Select porivains, porivanoi from FAT002A where codconce=" & pConcepto & " and cia='" & pCia & "'"
	Call executeQueryDb(DBSITE_SQL_MAGIC, rs, "OPEN", strSQL)
	ret = 0
	if (not rs.eof) then
		ret = rs("porivains")
	end if
	getTasaIVA = ret
End Function
'--------------------------------------------------------------------------------------------------------------------
' Mod: 2017-09-25 - JAS
Function dibujarPieSTD(pCAE, pFecVto, ByRef pOpdf,pCurrPage, pTotalPages)
    
    Call GF_setFont(pOpdf,"ARIAL", 6 , FONT_STYLE_NORMAL)	
	'Recueadro Observaciones
	Call GF_squareBox(pOpdf, 3, 695, 587 , 50, 0, "#FFFFFF", NEGRO, 1, PDF_SQUARE_ROUND)	
	Call GF_verticalLine(pOPDF, 497, 695, 50)
	Call GF_writeText(pOpdf, 5, 700, "OBSERVACIONES:" , 0)
    'Recuadro payment
	Call GF_setFont(pOpdf,"ARIAL", 8 , FONT_STYLE_NORMAL)	
	Call GF_squareBox(pOpdf, 3, 745, 587 , 75, 0, "#FFFFFF", NEGRO, 1, PDF_SQUARE_ROUND)    	
	'Nro de pagina.
	Call GF_writeTextAlign(pOpdf,3,825, "Pagina " & pCurrPage &" de "&pTotalPages , 200,PDF_ALIGN_LEFT)
	'CAE y su vto.
	Call GF_setFont(pOpdf,"ARIAL", 10 , FONT_STYLE_BOLD)
	Call GF_writeTextAlign(pOpdf,465,825, "C.A.E. Nº " & pCAE , 200,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(pOpdf,465,835, "Fecha Vto: " & GF_FN2DTE(pFecVto), 200,PDF_ALIGN_LEFT)
	
End Function
'--------------------------------------------------------------------------------------------------------------------
' Mod: 2017-09-25 - JAS
Function dibujarPieUltimaPaginaSTD(pCAE, pFecVto, ByRef pOpdf,pCurrPage, pTotalPages, pLetraFAC, pIdioma, pPuntoVenta, rsFAC, pConcepto)
    Dim strLeyenda, arr, ol, params

    'Codigo de Barras
    Call drawCodeBar(obtenerCodigoBarras(pCAE,pFecVto,pPuntoVenta, rsFac("tipcbt"), pLetraFAC),15,650,40)

    'Observaciones    
    if (g_Observaciones <> "") then 
        Call GF_setFont(pOpdf,"ARIAL", 7 , FONT_STYLE_NORMAL)
        'Se imprimen todas las lineas.        
        Call GF_writeTextPlus(pOpdf, 5, 750 + ol, g_Observaciones, 260, 8, PDF_ALIGN_LEFT)        
    end if    
    'Si es en dolares se imrpime el tipo de cambio
    if (CInt(rsFac("codmone")) = MONEDA_DOLAR_NUMERICO) then
        Call GF_setFont(pOpdf,"ARIAL", 8 , FONT_STYLE_BOLD)
        Call GF_writeTextAlign(pOpdf, 275, 810, GF_TRADUCIR("Tipo de Cambio:") & " $" & rsFAC("tcfin") , 300,PDF_ALIGN_LEFT)
        Call GF_setFont(pOpdf,"ARIAL", 8 , FONT_STYLE_NORMAL)
    end if		            
    
    'Leyenda    
    params = getParamsLeyenda(rsFAC)    
    strLeyenda = getDocLeyendaFactura(rsFac("recno"), rsFac("tipcbt"), pLetraFAC, pConcepto, rsFAC("codcia"),rsFac("codmone"), rsFac("feccbt_fn"), params)    
    if (strLeyenda <> "") then 					  
        if (Len(strLeyenda) > 950) then strLeyenda = Left(strLeyenda, 650) & "..."
	    Call GF_setFont(pOpdf,"ARIAL", 6 , FONT_STYLE_NORMAL)
	    Call GF_writeTextPlus(pOpdf, 275, 750, GF_TRADUCIR(strLeyenda) , 300, 8,PDF_ALIGN_JUSTIFY)			    
    end if	    
    	
End Function

%>