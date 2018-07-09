<%
Function validarFormatoDetalleGP(pLinea)
    'Formato esperado para la linea: 
    ' Ctro: 190091584616 C.P. 556962067
    ' 123456789012345678901234567890123
    '          1         2         3
    
    dim aux
    
    validarFormatoDetalleGP = false
    aux = split(Trim(pLinea), " ")        
    if (UBound(aux) = 3) then
        if (aux(0) = "Ctro:") then
            if (aux(2) = "C.P.") then
                if (isNumeric(aux(1))) then
                    if (isNumeric(aux(3))) then
                        validarFormatoDetalleGP = true
                    end if
                end if
            end if
        end if 
    end if
    
End Function
'-------------------------------------------------------------------------------------------------------
Function crearPDF_GP(pNroReg, ByRef pOpdf)
	Dim nroFac,rsDet,totalPages,currPage, letraFAC,rsFac,lineasTotales,puntoVenta,auxIdioma, paramsLeyenda()
	Dim flagProforma, flagFormato
	vDatosCAE = getDatosCAE(pNroReg)
	Set rsFac = getRSfactura(pNroReg)
	puntoVenta = GF_nDigits(rsFac("FCCMDV"),4)
	nroFac = puntoVenta &"-"&	GF_nDigits(rsFac("FCCMNR"),8)
	'Tomo la letra del comprobante
	letraFAC = rsFAC("FCCMTF")		
	flagProforma = false
	if (CInt(rsFac("FCCMST")) < FAC_AUTORIZADA) then flagProforma = true
	'Valido si el detalle cumple con el formato esperado  
	flagFormato = false
	Set sp_ret = executeSP(rsDet, "TFFL.TF101F1_GET_BY_PARAMETERS", rsFac("FCRGNR") & "||||1||0" & "$$totalRegistros")
	if (not rsDet.eof) then
	    flagFormato = validarFormatoDetalleGP(rsDet("COL2"))
	    if (flagFormato) then
	        'El formato de regsitro es correcto, leo el regsitro de acuerdo al concepto del cbte para que se edite correctamente.
	        Set sp_ret = executeSP(rsDet, "TFFL.TF101F1_GET_BY_PARAMETERS", rsFac("FCRGNR") & "||"&g_CodConcepto&"||1||0" & "$$totalRegistros")      
        end if	    
    end if        
	'Limpio las observaciones
	g_Observaciones = ""
	'SI ES GP TENGO Q MULTIPLICAR LA LINEA POR 4 POR Q EL DETALLE DEL GP OCUPA 3 RENGLONES SIEMPRE MAS UNO DE ESPACIO
	lineasTotales = Cdbl(sp_ret("totalRegistros"))
	lineasTotales = lineasTotales * 3
	if (CLng(rsFac("FCCLNR")) = PROVEEDOR_ESPECIAL_MAT) then lineasTotales = 1
	'Verifico si debo agregar una linea mas para los totales. Esto se da cuando la cantidad de lineas libres en la ultima pagina no alcanza para las lineas de totales.	
	totalPages = Ceil((lineasTotales)/MAX_LINEAS_PAGINA)	
	currPage= 1		
	if (not (rsDet.eof)) then	    
	    while ((not (rsDet.eof)) and (currPage <= totalPages))	    	    
	        if (currPage <> 1 ) then Call GF_newPage(pOpdf)	 	  	    
	        Call dibujarCabecera(nroFac, letraFAC, rsFac("FCCMFC"), rsFac("FCCMTP"),rsFac("FCCMDV"),pOpdf)
	        Call dibujarOrigenDestinoLocal(rsFac("FCCLNR"),pOpdf)	    
	        Call dibujarDatosCompraLocal(rsFac("FCRGNR"),pOpdf)	    
	        Call dibujarDetalleTitulosLocal(pOpdf, flagProforma)
	        if (flagFormato) then
		        Call dibujarDetalleContenidoGP(pNroReg, rsDet, rsFac("FCCLNR"),currPage,totalPages,pOpdf)
            else
                Call dibujarDetalleContenidoSTD(rsDet, pOpdf)
            end if		        
		    Call dibujarPieGP(vDatosCAE(0),vDatosCAE(1),pOpdf,currPage,totalPages)
	        currPage= currPage + 1	   
        wend
    else
        Call dibujarCabecera(nroFac, letraFAC, rsFac("FCCMFC"), rsFac("FCCMTP"),rsFac("FCCMDV"),pOpdf)
        Call dibujarOrigenDestinoLocal(rsFac("FCCLNR"),pOpdf)	    
        Call dibujarDatosCompraLocal(rsFac("FCRGNR"),pOpdf)	    
        Call dibujarDetalleTitulosLocal(pOpdf, flagProforma)        
        Call dibujarPieGP(vDatosCAE(0),vDatosCAE(1),pOpdf,1,1)
    end if        
    'Se imprimen los totales.    
    Call dibujarTotalesLocal(rsFac("FCRGNR"),pOpdf)    
    Call dibujarPieUltimaPaginaGP(vDatosCAE(0),vDatosCAE(1),pOpdf,currPage, rsFac("FCCLNR"), rsFac("FACTO3"), totalPages,letraFAC,auxIdioma, puntoVenta, rsFac)    
    
    'Si es una factura del MAT se imprime la página con el detalle.
    if ((CLng(rsFac("FCCLNR")) = PROVEEDOR_ESPECIAL_MAT) and (flagFormato)) then Call imprimirDetalleAuxiliar(rsDet, rsFAC, nroFac, pOpdf)
        
End Function
'--------------------------------------------------------------------------------------------------------------------
Function imprimirDetalleAuxiliar(pDatos, rsFAC, pNroFac, ByRef pOpdf)
    Dim indexIni, pagina, linea, auxLeyenda

    pagina = 0    
    'Se reinicia el recordset del detalle.
    pDatos.MoveFirst()    
    while (not pDatos.eof)        
        Call GF_newPage(pOpdf)                
        Call GF_squareBox(pOpdf, 3, 5, 587, 140, 0, "#FFFFFF", NEGRO, 1, PDF_SQUARE_ROUND) 
        'logo
	    Call GF_writeImage(pOpdf, Server.MapPath("..\Images\logo1.gif"),10, 6, 200, 50, 0)        
	    'Titulo del reporte
	    Call GF_setFont(pOpdf,"ARIAL", 12 , FONT_STYLE_BOLD)
	    Call GF_writeTextAlign(pOpdf,3, 85, "DETALLE DEL COMPROBANTE Nº " & pNroFac , 587,PDF_ALIGN_CENTER)	                         
        Call dibujarDetalleTitulosGPAuxiliar(pOpdf)
        linea = 0         
        indexIni = 163
        indexIni = dibujarDetalleDescargasGP(indexIni, pDatos, rsFac("FCCLNR"), pOpdf, MAX_LINEAS_PAGINA_AUXILIAR)                        	                         
        'Nro de pagina.
        pagina = pagina + 1
	    Call GF_writeTextAlign(pOpdf,3,825, "Pagina " & pagina, 200,PDF_ALIGN_LEFT)
    wend
    Call GF_writeTextAlign(pOpdf,3, indexIni, "--- Fin del Reporte ---" , 587,PDF_ALIGN_CENTER)
End Function
'--------------------------------------------------------------------------------------------------------------------
Function dibujarDetalleTitulosGPAuxiliar(ByRef pOPDF)
	Call GF_squareBox(pOPDF,  3, 145,  62, 15,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_ROUND)
	Call GF_squareBox(pOPDF, 65, 145, 350,15,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_ROUND) 
	Call GF_squareBox(pOPDF,415, 145,  78,15,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_ROUND) 
	Call GF_squareBox(pOPDF,493, 145,  97,15,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_ROUND)	
	Call GF_setFont(pOPDF,"ARIAL", 10 , FONT_STYLE_NORMAL)
	Call GF_writeTextAlign(pOPDF,  3,148, "CANTIDAD"  ,  62,PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(pOPDF, 65,148, "DETALLE"   , 350,PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(pOPDF,415,148, "P.UNITARIO",  78,PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(pOPDF,493,148, "TOTAL"     ,  97,PDF_ALIGN_CENTER)
	Call GF_squareBox(pOpdf,3,160,587 ,660,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_ROUND)
End Function
'--------------------------------------------------------------------------------------------------------------------
Function dibujarDetalleContenidoGP(pNroReg, ByRef pDatos, pCliente, pCurrPage,pTotalPages,ByRef pOpdf)
	Dim i,strDetalle,auxVCto, cdMoneda, myGravado, myNoGravado, myIVA, myPercepcionIVA, myPercepcionIIBB, myTasaIVA, myTotal, indexIni
	i = 0	
	indexIni = 268
	if (CLng(pCliente) = PROVEEDOR_ESPECIAL_MAT) then
	    Call calcularTotalesLocal(pNroReg, cdMoneda, myGravado, myNoGravado, myIVA, myPercepcionIVA, myPercepcionIIBB, myTasaIVA, myTotal)
	    Call GF_setFont(pOpdf,"COURIER", 9 , FONT_STYLE_NORMAL)	    
	    Call GF_writeTextAlign(pOpdf,   65, indexIni , "Recupero de gastos Caratula Nº" & Trim(pDatos("CTOTERCERO")) , 62, PDF_ALIGN_LEFT)				
	    Call GF_writeTextAlign(pOpdf, 493, indexIni  + (i*10), getSimboloMoneda(cdMoneda) &" "& GF_EDIT_DECIMALS(myGravado, 2) , 97, PDF_ALIGN_RIGHT)	    	    
	    pDatos.MoveLast()
	    pDatos.MoveNext()
	else				
	    Call GF_setFont(pOpdf,"COURIER", 10 , FONT_STYLE_BOLD)			
		Call GF_writeTextAlign(pOpdf,   65, indexIni , "GASTOS DE ACONDICIONAMIENTO" , 62, PDF_ALIGN_LEFT)				
        indexIni = indexIni + 12
		Call dibujarDetalleDescargasGP(indexIni, pDatos, pCliente,pOpdf, MAX_LINEAS_PAGINA)
	end if	    	    	
End Function
'--------------------------------------------------------------------------------------------------------------------
Function dibujarDetalleDescargasGP(pIndexIni, ByRef pDatos, pCliente,ByRef pOpdf, maxLineas)
	Dim i, strDetalle,auxVCto
	i = 0	
	While (getControlLineByPage(pDatos, i, maxLineas))			
		Call GF_setFont(pOpdf,"COURIER", 9 , FONT_STYLE_NORMAL)
		if (CDbl(pDatos("COL1")) <> 0) then Call GF_writeTextAlign(pOpdf,   3, pIndexIni  + (i*10), GF_EDIT_DECIMALS(pDatos("COL1"), 3) , 62, PDF_ALIGN_CENTER)
		strLine = "Ctro:" & Trim(GF_EDIT_CONTRATO(pDatos("CDPRODUCTO"),pDatos("CDSUCURSAL"),GF_nDigits(pDatos("CDOPERACION"),2),pDatos("NROCTO"),pDatos("COSECHA"))) & " C.P.:"& Trim(GF_nDigits(pDatos("CARTAPORTE1"),4)) &"-"&Trim(pDatos("CARTAPORTE2"))				
		if (Trim(pDatos("CTOTERCERO")) <> "") then 
		    auxVCto = TEXTO_CTO_GRAL
		    if (CLng(pCliente) = PROVEEDOR_ESPECIAL_MAT) then auxVCto = TEXTO_CTO_MAT
			if (Len(Trim(pDatos("CTOTERCERO"))) < 17) then
				strLine = strLine & " " & auxVCto & ":"&Trim(pDatos("CTOTERCERO"))
			else
				Call GF_writeTextAlign(pOpdf, 65, pIndexIni+20  + (i*10), auxVCto & ":"&Trim(pDatos("CTOTERCERO")) ,350, PDF_ALIGN_LEFT)
			end if
		end if
		Call GF_writeTextAlign(pOpdf,  65, pIndexIni  + (i*10), strLine ,350, PDF_ALIGN_LEFT)
		Call GF_writeTextAlign(pOpdf,  65, (pIndexIni+10)  + (i*10), Trim(pDatos("DSCONCEPTO")) &" "& pDatos("VALOR") & " Prod.:"&Trim(pDatos("DSPRODUCTO")),350, PDF_ALIGN_LEFT)
		if (Trim(pDatos("MONEDA")) <> "") then auxMoneda = getSimboloMoneda(pDatos("MONEDA"))
		if (CDbl(pDatos("COL3")) <> 0) then Call GF_writeTextAlign(pOpdf, 415, pIndexIni  + (i*10), auxMoneda &" "& GF_EDIT_DECIMALS(pDatos("COL3"), 3) , 78, PDF_ALIGN_RIGHT)
		if (CDbl(pDatos("COL4")) <> 0) then Call GF_writeTextAlign(pOpdf, 493, pIndexIni  + (i*10), auxMoneda &" "& GF_EDIT_DECIMALS(pDatos("COL4"), 2) , 97, PDF_ALIGN_RIGHT)
		i = i + 3
		pDatos.MoveNext()		    
	wend
	dibujarDetalleDescargasGP = pIndexIni  + (i*10)
End Function
'--------------------------------------------------------------------------------------------------------------------
Function dibujarPieGP(pCAE, pFecVto, ByRef pOpdf,pCurrPage,pTotalPages)
	Dim strSQL
			
    'CAE y su vto.
	Call GF_setFont(pOpdf,"ARIAL", 10 , FONT_STYLE_BOLD)
	Call GF_writeTextAlign(pOpdf,465,825, "C.A.E. Nº " & pCAE , 200,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(pOpdf,465,835, "Fecha Vto: " & GF_FN2DTE(pFecVto), 200,PDF_ALIGN_LEFT)
	'Recuadro payment
	Call GF_setFont(pOpdf,"ARIAL", 8 , FONT_STYLE_NORMAL)
    Call GF_squareBox(pOpdf,3,745,587 ,75,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_ROUND)    
	'Nro de pagina.
	Call GF_writeTextAlign(pOpdf,3,825, "Pagina " & pCurrPage &" de "&pTotalPages , 200,PDF_ALIGN_LEFT)
	
End Function
'--------------------------------------------------------------------------------------------------------------------
Function dibujarPieUltimaPaginaGP(pCAE, pFecVto, ByRef pOpdf,pCurrPage, pCliente, pCdProveedor, pTotalPages, pLetraFAC, pIdioma, pPuntoVenta, rsFAC)
    Dim strLeyenda, arr, ol, params
    
		
    if (CLng(pCliente) <> PROVEEDOR_ESPECIAL_MAT) then  
        strSQL = "select RAZSOC DESCR1 from MERFL.TCB6A1F1 where NROPRO = " & pCdProveedor
	    Call executeQuery(rs, "OPEN", strSQL)
	    if not rs.EoF then
		    if (rs("DESCR1") <> "") then 
		        g_Observaciones = g_Observaciones & " Corredor: (" & pCdProveedor & ") " & rs("DESCR1") & OBS_EOL_TOKEN				
		    end if	
	    end if		    
	    'Leyenda
        params = getParamsLeyenda(rsFAC)    
        strLeyenda = getDocLeyendaFactura(rsFac("FCRGNR"), rsFac("FCCMTP"), pLetraFAC, rsFAC("LDLYCD"), rsFAC("FCSCNR"), rsFac("FCMNCD"), pIdioma, 1, rsFac("FCCMFC"), params)    
        if (strLeyenda <> "") then 					  
            if (Len(strLeyenda) > 950) then strLeyenda = Left(strLeyenda, 650) & "..."
	        Call GF_setFont(pOpdf,"ARIAL", 6 , FONT_STYLE_NORMAL)
	        Call GF_writeTextPlus(pOpdf, 275, 750, GF_TRADUCIR(strLeyenda) , 300, 8,PDF_ALIGN_JUSTIFY)			    
        end if	
    end if            
    	
    Call dibujarPieUltimaPaginaLocal(pCAE, pFecVto, pOpdf,pCurrPage, pTotalPages, pLetraFAC, pIdioma, pPuntoVenta, rsFAC)        
	
End Function

%>