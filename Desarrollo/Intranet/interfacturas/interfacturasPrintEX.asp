<%
'---------------------------------------------------------------------------------------------------------------------
Function obtenerDatosCompradorEx(pNroPro)
	Dim rs,conn,strSQL,rtrn()
	redim rtrn(5)	
	strSQL = "select * from MERFL.TCB6A1F1 where NROPRO = " & pNroPro	
	Call executeQuery(rs, "OPEN", strSQL)	
	if (not rs.EoF) then
	    rtrn(0) = rs("NOMAMP")
		rtrn(1) = GF_STR2CUIT(rs("NRODOC"))		
		rtrn(2) = rs("DOMICI") & " - " & rs("CODPOS") & " " & rs("LOCALI")
		'JAS - PARCHE POR FALTA DE ESPACIO EN CAMPO DE LOCALIDAD - HORRIBLE!!!
		if (rs("LOCALI") = "HAMBURG GERMAN") THEN rtrn(2) = rtrn(2) & "Y"
		rtrn(3) = "EXENTO OPERACIÓN DE EXPORTACÓN"
		rtrn(4) = pNroPro
	end if
	obtenerDatosCompradorEx = rtrn
End Function
'---------------------------------------------------------------------------------------------------------------------
Function obtenerDatosShippingForm(pNroIng)
	Dim strSQL,conn,rs,rtrn()
	redim rtrn(5)	
	strSQL = "select * from tffl.tf112 where INNRRG = " & pNroIng
	Call GF_BD_COMPRAS(rs, conn, "OPEN", strSQL)	
	rtrn(0) = "0"
	rtrn(1) = "0"
	rtrn(2) = "0"
	rtrn(3) = getDescripcionProveedor(CD_TOEPFER)
	rtrn(4) = GF_FN2DTE("19000101")
	if (not rs.EoF) then
		rtrn(0) = rs("INCTOI")
		rtrn(1) = rs("INDE")
		rtrn(2) = rs("INA")
		rtrn(3) = getDescripcionProveedor(CD_TOEPFER)
		rtrn(4) = GF_FN2DTE(rs("INBL"))
	end if	
	obtenerDatosShippingForm = rtrn
End Function
'----------------------------------------------------------------------------------------------------------------------
Function dibujarOrigenDestinoEx(pClNbr,ByRef pOPdf)
    Dim datosComprador    
    datosComprador = obtenerDatosCompradorEx(pClNbr)	
    
    'Estructura
    Call GF_squareBox(pOPdf,3,150,587 ,50,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_ROUND) 
	Call GF_writeImage(pOPdf, Server.MapPath("..\Images\facturas\MarcaAguaToepfer.gif"),5, 152, 570, 45, 0)	
    
    'Datos
	Call GF_setFont(pOPdf,"ARIAL", 8 , FONT_STYLE_NORMAL)
	Call GF_writeTextAlign(pOPdf,15,155, "SEÑORES:" , 200,PDF_ALIGN_LEFT)	
	Call GF_writeTextAlign(pOPdf,400,155, "C.U.I.T.:" , 200,PDF_ALIGN_LEFT)	
	
	Call GF_writeTextAlign(pOPdf,15,185, "DOMICILIO:", 200,PDF_ALIGN_LEFT)	
	Call GF_writeTextAlign(pOPdf,400,185, "I.V.A.:" , 200,PDF_ALIGN_LEFT)	
	
	'muestra de datos
	Call GF_setFont(pOPdf,"ARIAL", 8 , FONT_STYLE_BOLD)
	Call GF_writeTextAlign(pOPdf, 65,155, Trim(datosComprador(0)) & " ("&datosComprador(4)&")", 200,PDF_ALIGN_LEFT)	
	Call GF_writeTextAlign(pOPdf, 65,185, datosComprador(2), 200,PDF_ALIGN_LEFT)	
	Call GF_writeTextAlign(pOPdf,440,155, datosComprador(1), 200,PDF_ALIGN_LEFT)	
	Call GF_writeTextAlign(pOPdf,425,185, datosComprador(3), 200,PDF_ALIGN_LEFT)	
	Call GF_squareBox(pOPdf,3,200,587 ,50,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_ROUND) 
End Function
'----------------------------------------------------------------------------------------------------------------------
Function dibujarDatosCompraEx(pNroRegFac, ByRef pOPdf)
    datosCompra = obtenerDatosShippingForm(pNroRegFac)
    'Estructura
	Call GF_setFont(pOPdf,"ARIAL", 8 , FONT_STYLE_NORMAL)
	Call GF_writeTextAlign(pOPdf,15,205, "CONTRACT:" , 200,PDF_ALIGN_LEFT)	
	Call GF_writeTextAlign(pOPdf,15,220, "SHIPPED FROM" , 200,PDF_ALIGN_LEFT)	
	
	Call GF_writeTextAlign(pOPdf,240,220, "TO" , 200,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(pOPdf,15,234, "BY" , 200,PDF_ALIGN_LEFT)	
	Call GF_writeTextAlign(pOPdf,450,234, "B/L DATE" , 200,PDF_ALIGN_LEFT)
	
	'muestra de datos
	Call GF_setFont(pOPdf,"ARIAL", 8 , FONT_STYLE_BOLD)
	Call GF_writeTextAlign(pOPdf,85,205, datosCompra(0) , 200,PDF_ALIGN_LEFT)	
	Call GF_writeTextAlign(pOPdf,85,220, datosCompra(1) , 200,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(pOPdf,260,220, datosCompra(2) , 200,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(pOPdf,40,234, datosCompra(3) , 200,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(pOPdf,495,234, datosCompra(4) , 200,PDF_ALIGN_LEFT)
End Function
'----------------------------------------------------------------------------------------------------------------------
Function dibujarDetalleEx(ByRef pDatos, ByRef pOPdf)
    Call dibujarDetalleTitulosEx(pOPdf)
	Call GF_squareBox(pOPdf,3,265,587 ,480,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_ROUND)
	'Call GF_writeImage(pOPdf, Server.MapPath("..\Images\facturas\MarcaAguaLogoToepfer.gif"),100, 325, 374, 373, 0)
	Call dibujarDetalleContenidoEx(pDatos,pOPdf)
End Function
'----------------------------------------------------------------------------------------------------------------------
Function dibujarDetalleTitulosEx(ByRef pOPdf)
    'recuadro titulo detalle	
    Call GF_squareBox(pOPdf,  3, 250, 72, 15,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_ROUND) 	
	Call GF_squareBox(pOPdf, 75, 250, 75,15,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_ROUND) 
	Call GF_squareBox(pOPdf,150, 250,300,15,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_ROUND) 
	Call GF_squareBox(pOPdf,450, 250, 140,15,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_ROUND) 
	
	Call GF_setFont(pOPdf,"ARIAL", 10 , FONT_STYLE_NORMAL)
	Call GF_writeTextAlign(pOPdf,  3,253, "MARKS"  , 72,PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(pOPdf, 75,253, "NUMBER" , 75,PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(pOPdf,150,253, "PACKAGES, DESCRIPTION OF GOODS, PRICE" , 300,PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(pOPdf,450,253, "AMOUNT" , 140,PDF_ALIGN_CENTER)	
End Function
'----------------------------------------------------------------------------------------------------------------------
Function dibujarDetalleContenidoEx(ByRef pDatos,ByRef pOPdf)
	Dim i,auxTotal
	i = 0
	auxTotal = 0
	indexIni = 268
	While (getControlLineByPage(pDatos, i, MAX_LINEAS_PAGINA))		
		Call GF_setFont(pOPdf,"COURIER", 9 , FONT_STYLE_NORMAL)	
	    Call GF_writeTextAlign(pOPdf,10,indexIni  + (i*12), pDatos("COL1") , 60,PDF_ALIGN_CENTER)
	    if (Cdbl(pDatos("COL2")) > 0) then Call GF_writeTextAlign(pOPdf,80,indexIni + (i*12) , GF_EDIT_DECIMALS(cdbl(pDatos("COL2"))*100,2) , 65,PDF_ALIGN_CENTER)
	    Call GF_writeTextAlign(pOPdf,155,indexIni + (i*12), pDatos("COL3") , 300,PDF_ALIGN_LEFT)
	    if (Cdbl(pDatos("COL4")) > 0) then Call GF_writeTextAlign(pOPdf,455,indexIni + (i*12),"USD " & GF_EDIT_DECIMALS(cdbl(pDatos("COL4")),2) , 130,PDF_ALIGN_RIGHT)
	    auxTotal = auxTotal + cdbl(pDatos("COL4"))
		i = i + 1
		pDatos.MoveNext()
	wend
	'Se imprime la leyenda del importe
	i = i + 1
	Call GF_writeTextAlign(pOPdf,155,270 + (i*12), "Say U$S: " , 250,PDF_ALIGN_LEFT)
	i = i + 1
    Call GF_writeTextPlus(pOPdf,155, 270 + (i*12), num2words(cdbl(auxTotal)/100), 230, 12, PDF_ALIGN_LEFT)
End Function
'----------------------------------------------------------------------------------------------------------------------
Function dibujarPieEx(pCAE, pFecVto, ByRef pOPdf,pCurrPage,pTotalPages,pLetra,pTipo,pMoneda,pIdioma,pSecuencia, pPuntoVenta)
	
	'Codigo de Barras	
	Call drawCodeBar(obtenerCodigoBarras(pCAE,pFecVto,pPuntoVenta, pTipo, pLetra),15,700,40)
	
	'CAE y su vto.
	Call GF_setFont(pOPdf,"ARIAL", 10 , FONT_STYLE_BOLD)
	Call GF_writeTextAlign(pOPdf,465,825, "C.A.E. Nº " & pCAE , 200,PDF_ALIGN_LEFT)	
	Call GF_writeTextAlign(pOPdf,465,835, "Fecha Vto: " & GF_FN2DTE(pFecVto), 200,PDF_ALIGN_LEFT)	
	
	'Nro de pagina.	
	Call GF_writeTextAlign(pOpdf,3,825, "Pagina " & pCurrPage &" de "&pTotalPages , 200,PDF_ALIGN_LEFT)
	
	'Recuadro payment
	Call GF_setFont(pOPdf,"ARIAL", 6 , FONT_STYLE_NORMAL)
    Call GF_squareBox(pOPdf,3,745,587 ,75,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_ROUND)     
    Call GF_writeTextAlign(pOPdf,15,755, "PAYMENT" , 135,PDF_ALIGN_LEFT)
    
End Function
'----------------------------------------------------------------------------------------------------------------------
Function dibujarTotalesEx(pNroReg,ByRef pOpdf)
    Dim txtImporte, strSQL, rs, simboloMoneda        
    strSQL="Select FCTTGR*100 TOTAL, FCMNCD from TFFL.TF100F1 C left join TFFL.TF111 D on C.FCRGNR=D.FANRRG where FCRGNR=" & pNroReg
    Call executeQuery(rs, "OPEN", strSQL)
    if (not rs.Eof) Then
        simboloMoneda = getSimboloMonedaLetras(rs("FCMNCD"))
        'Se imrpime el total general
        Call GF_setFont(pOpdf,"ARIAL", 10 , FONT_STYLE_BOLD)
        Call GF_writeTextAlign(pOpdf, 400, 726, "TOTAL", 250,PDF_ALIGN_LEFT)
        Call GF_writeTextAlign(pOpdf, 473, 726, simboloMoneda, 250,PDF_ALIGN_LEFT)
        Call GF_writeTextAlign(pOpdf, 473, 726, GF_EDIT_DECIMALS(rs("TOTAL"), 2), 114, PDF_ALIGN_RIGHT)        
    end if
End Function
'---------------------------------------------------------------------------------------------------------------------
Function crearPDF_Ex(pNroReg,ByRef pOPdf)
	Dim nroFac,rsDet,totalPages, currPage, letraFAC,rsFac,puntoVenta,auxIdioma,auxSecuencia
	vDatosCAE = getDatosCAE(pNroReg)
	Set rsFac = getRSfactura(pNroReg)		
	puntoVenta = GF_nDigits(rsFac("FCCMDV"),4)
	nroFac = puntoVenta &"-"&	GF_nDigits(rsFac("FCCMNR"),8)	
	'Tomo la letra del comprobante
	letraFAC = rsFAC("FCCMTF")
	if (CInt(rsFac("FCCMST")) < FAC_AUTORIZADA) then letraFAC = "P"
	Set sp_ret = executeSP(rsDet, "TFFL.TF101F1_GET_BY_PARAMETERS", rsFac("FCRGNR") & "||"&g_CodConcepto&"||1||0" & "$$totalRegistros")
	lineasTotales = Cdbl(sp_ret("totalRegistros"))
	'Verifico si debo agregar una linea mas para los totales. Esto se da cuando la cantidad de lineas libres en la ultima pagina no alcanza para las lineas de totales.	
    totalPages = Ceil((lineasTotales)/MAX_LINEAS_PAGINA)
	'Response.Write lineasTotales &" | "& totalPages
	'Response.End
	currPage= 1
	while ((not (rsDet.eof)) or (currPage <= totalPages))	    
	    if (currPage <> 1 ) then Call GF_newPage(oPDF)	    
	    Call dibujarCabecera(nroFac, letraFAC, rsFac("FCCMFC"), rsFac("FCCMTP"),rsFac("FCCMDV"),pOPdf)
	    Call dibujarOrigenDestinoEx(rsFac("FCCLNR"),pOPdf)
	    Call dibujarDatosCompraEx(pNroReg,pOPdf)	    
	    Call dibujarDetalleEx(rsDet,pOPdf)	    	    
	    Call dibujarPieEx(vDatosCAE(0),vDatosCAE(1),pOPdf,currPage,totalPages,letraFAC,rsFac("FCCMTP"),rsFac("FCMNCD"),auxIdioma,auxSecuencia, puntoVenta)
	    currPage= currPage + 1
    wend
    'Se imprimen los totales.
    Call dibujarTotalesEx(rsFac("FCRGNR"), pOpdf)        
End Function
%>