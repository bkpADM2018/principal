<%
'TAREA 1827
'-----------------------------------------------------------------------------------------------------------------------------
Function crearPDF_LS(pNroReg, ByRef pOpdf)
	Dim rsDet, rsFac, auxY, flagProforma, auxMoneda, precioTN, negocioCliente
	Dim ctoProducto, ctoSucursal, ctoOperacion, ctoNumero, ctoCosecha, precioOperacion, cantidadTN, alicuota, fechaOperacion
	Dim dsDestino
	vDatosCAE = getDatosCAE(pNroReg)
	Set rsFac = getRSfactura(pNroReg)
    flagProforma = false
	if (CInt(rsFac("FCCMST")) < FAC_AUTORIZADA) then flagProforma = true
	if (not rsFac.eof) then
	    Call calcularValoresGlobales(rsFAC, ctoProducto, ctoSucursal, ctoOperacion, ctoNumero, ctoCosecha, cantidadTN, precioTN, precioOperacion, auxMoneda, dsDestino, alicuota, fechaOperacion, negocioCliente)
        Call dibujarCabeceraLS(pOpdf, rsFAC("FCCMFC"), rsFac("FCCMTP"),rsFac("FCCMDV"),flagProforma)
        Call GF_writeTextAlign(pOpdf,300,60, "ID REGISTRO: " & rsFAC("FCRGNR") , 260,PDF_ALIGN_RIGHT)
	    Call dibujarIntervinientesLS(pOpdf,auxY,vDatosCAE(0),rsFac("FCCLNR"),rsFac("FCCMTP"),rsFac("FCCMDV"),rsFac("FACTO3"))	    
        Call dibujarCondicionesOperacionLS(pOpdf,auxY,auxMoneda,precioTN, fechaOperacion, ctoProducto, dsDestino)
	    if ((Cdbl(pTipoFac) = FAC_LIQUIDACION_SECUNDARIA_CREDITO)or(Cdbl(pTipoFac) = FAC_LIQUIDACION_SECUNDARIA_DEBITO))then
            'NOTA: Si es una nota de credito o debito debo fijarme si trae varios registro detalles o no. 
            '      En ese caso hago un while por cada corte de control dependieno de la operacion o ajuste
        else            
            Call dibujarDeduccionesLS(pOpdf, auxY, auxMoneda, abs(CDbl(rsFAC("FAIMPN"))), 0, rsFAC("FAPORC"))
            Call dibujarPercepcionesLS(pOpdf, auxY, rsFac("FCRGNR"), auxMoneda, precioOperacion, rsFAC("FAIMP"))
            Call dibujarOperacionLS(pOpdf, auxY, rsFac("FCCMTP"), cantidadTN, precioTN, precioOperacion, alicuota, rsFAC("FAIVAI"), abs(CDbl(rsFAC("FAIMPN"))), rsFAC("FAIMP"),rsFAC("FAIMP2"), rsFAC("FCTTGR"),auxMoneda)
            Call dibujarPiePaginaLS(pOpdf, auxY, rsFac("FCCLNR"), GF_EDIT_CONTRATO(ctoProducto, ctoSucursal, ctoOperacion, ctoNumero, ctoCosecha), rsFac("FAVTO"), negocioCliente)
        end if
    end if
End Function
'-----------------------------------------------------------------------------------------------------------------------------
Function dibujarDeduccionesLS(pOpdf, ByRef pY, pMoneda, pImporteBase, pImporteIVA, pPorcPago)
    
    if (CDbl(pImporteBase) > 0) then
        'Dibujo la seccion    
        Call GF_squareBox(pOpdf,20,pY,550,18,0,"#E6E6E6",NEGRO,1,PDF_SQUARE_NORMAL)
        Call GF_setFont(oPDF,"ARIAL", 10 , FONT_STYLE_BOLD)
        Call GF_writeTextAlign(pOpdf,22,pY + 5, "DEDUCCIONES" , 400,PDF_ALIGN_LEFT)
        pY = pY + 18
        Call GF_setFont(oPDF,"ARIAL", 8 , FONT_STYLE_BOLD)
        Call GF_squareBox(pOpdf,20,pY,220,15,0,"#E6E6E6",NEGRO,1,PDF_SQUARE_NORMAL)
        Call GF_squareBox(pOpdf,240,pY,110,15,0,"#E6E6E6",NEGRO,1,PDF_SQUARE_NORMAL)
        Call GF_squareBox(pOpdf,350,pY,110,15,0,"#E6E6E6",NEGRO,1,PDF_SQUARE_NORMAL)
        Call GF_squareBox(pOpdf,460,pY,110,15,0,"#E6E6E6",NEGRO,1,PDF_SQUARE_NORMAL)
        Call GF_writeTextAlign(pOpdf,20,pY + 4, "Concepto" , 220,PDF_ALIGN_CENTER)
        Call GF_writeTextAlign(pOpdf,240,pY + 4, "Base de Cálculo ($)" , 110,PDF_ALIGN_CENTER)
        Call GF_writeTextAlign(pOpdf,350,pY + 4, "Importe IVA ($)" , 110,PDF_ALIGN_CENTER)
        Call GF_writeTextAlign(pOpdf,460,pY + 4, "Deducciones c/IVA ($)" , 110,PDF_ALIGN_CENTER)
        pY = pY + 15
        Call GF_setFont(oPDF,"ARIAL", 8 , FONT_STYLE_NORMAL)
        Call GF_squareBox(pOpdf,20,pY,220,15,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_NORMAL)
        Call GF_squareBox(pOpdf,240,pY,110,15,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_NORMAL)
        Call GF_squareBox(pOpdf,350,pY,110,15,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_NORMAL)
        Call GF_squareBox(pOpdf,460,pY,110,15,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_NORMAL)
        Call GF_writeTextAlign(pOpdf,20,pY + 4, "Porcentaje Pendiente de Cobro (" & (100 - cDbl(pPorcPago)) & "%)" , 220,PDF_ALIGN_CENTER)
        Call GF_writeTextAlign(pOpdf,240,pY + 4, pMoneda & " " & GF_EDIT_DECIMALS(Cdbl(pImporteBase)*100, 2) , 110,PDF_ALIGN_CENTER)
        Call GF_writeTextAlign(pOpdf,350,pY + 4, pMoneda & " " & GF_EDIT_DECIMALS(Cdbl(pImporteIVA)*100, 2) , 110,PDF_ALIGN_CENTER)
        Call GF_writeTextAlign(pOpdf,460,pY + 4, pMoneda & " " & GF_EDIT_DECIMALS((Cdbl(pImporteBase) + CDbl(pImporteIVA))*100, 2) , 110,PDF_ALIGN_CENTER)
        pY = pY + 30
    end if        
End Function
'-----------------------------------------------------------------------------------------------------------------------------
Function dibujarPercepcionesLS(pOpdf, ByRef pY, pNroReg, pMoneda, pImporteBase, pImporteIVA)
    
    DIm rs, importeIIBB, dsProvincia
    'Obtengo las percepciones de IIBB del registro    
    Call executeSP(rs, "TFFL.TF114GET_BY_FANRR4", pNroReg)
    
    if ((CDbl(pImporteIVA) > 0) or (not rs.eof)) then
        'Dibujo la seccion    
        Call GF_squareBox(pOpdf,20,pY,550,18,0,"#E6E6E6",NEGRO,1,PDF_SQUARE_NORMAL)
        Call GF_setFont(oPDF,"ARIAL", 10 , FONT_STYLE_BOLD)
        Call GF_writeTextAlign(pOpdf,22,pY + 5, "PERCEPCIONES" , 400,PDF_ALIGN_LEFT)
        pY = pY + 18        
        Call GF_setFont(oPDF,"ARIAL", 8 , FONT_STYLE_BOLD)
        Call GF_squareBox(pOpdf,20,pY,220,15,0,"#E6E6E6",NEGRO,1,PDF_SQUARE_NORMAL)
        Call GF_squareBox(pOpdf,240,pY,110,15,0,"#E6E6E6",NEGRO,1,PDF_SQUARE_NORMAL)
        Call GF_squareBox(pOpdf,350,pY,110,15,0,"#E6E6E6",NEGRO,1,PDF_SQUARE_NORMAL)
        Call GF_squareBox(pOpdf,460,pY,110,15,0,"#E6E6E6",NEGRO,1,PDF_SQUARE_NORMAL)
        Call GF_writeTextAlign(pOpdf,20,pY + 4, "Concepto" , 220,PDF_ALIGN_CENTER)
        Call GF_writeTextAlign(pOpdf,240,pY + 4, "Base de Cálculo ($)" , 110,PDF_ALIGN_CENTER)
        Call GF_writeTextAlign(pOpdf,350,pY + 4, "Alicuota (%)" , 110,PDF_ALIGN_CENTER)
        Call GF_writeTextAlign(pOpdf,460,pY + 4, "Importe Percepción ($)" , 110,PDF_ALIGN_CENTER)        
        Call GF_setFont(oPDF,"ARIAL", 8 , FONT_STYLE_NORMAL)
        'Se dibujan las percepciones de IVA.
        if (CDbl(pImporteIVA) > 0) then 
            pY = pY + 15
            Call dibujarPercepcionesLineaLS(pOpdf, pY, pMoneda, "Percepción de IVA", pImporteBase, GF_EDIT_DECIMALS((CDbl(pImporteIVA)/pImporteBase)*10000, 2), pImporteIVA)           
        end if
        'Se dibujan las percepciones de IIBB.
        while (not rs.eof)
            pY = pY + 15
            importeIIBB = CDbl(rs("FAIMP4"))
            Call executeSP(rs2, "MERFL.MER1K2F1_GET_BY_CODIPO", rs("FAPRO4"))
            dsProvincia = "# ERROR #"
            if (not rs2.eof) then dsProvincia=Trim(rs2("DESCPO"))
            Call dibujarPercepcionesLineaLS(pOpdf, pY, pMoneda, "Percepción IIBB " & dsProvincia, pImporteBase, GF_EDIT_DECIMALS((importeIIBB/pImporteBase)*10000, 2), importeIIBB)
            rs.MoveNext()
        wend                
        pY = pY + 30
    end if        
End Function
'-----------------------------------------------------------------------------------------------------------------------------
Function dibujarPercepcionesLineaLS(pOpdf, ByRef pY, pMoneda, pConcepto, pImporteBase, pAlicuota, pImportePercepcion)
    Call GF_squareBox(pOpdf,20,pY,220,15,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_NORMAL)
    Call GF_squareBox(pOpdf,240,pY,110,15,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_NORMAL)
    Call GF_squareBox(pOpdf,350,pY,110,15,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_NORMAL)
    Call GF_squareBox(pOpdf,460,pY,110,15,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_NORMAL)
    Call GF_writeTextAlign(pOpdf,20,pY + 4, pConcepto , 220,PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(pOpdf,240,pY + 4, pMoneda & " " & GF_EDIT_DECIMALS(Cdbl(pImporteBase)*100, 2) , 110,PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(pOpdf,350,pY + 4, pAlicuota , 110,PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(pOpdf,460,pY + 4, pMoneda & " " & GF_EDIT_DECIMALS(Cdbl(pImportePercepcion)*100, 2) , 110,PDF_ALIGN_CENTER)
End Function
'-----------------------------------------------------------------------------------------------------------------------------    
Function calcularValoresGlobales(rsFAC, ByRef ctoProducto, ByRef ctoSucursal, ByRef ctoOperacion, ByRef ctoNumero, ByRef ctoCosecha, ByRef cantidad, ByRef precioTN, ByRef precioOperacion, ByRef auxMoneda, ByRef dsDestino, ByRef alicuota, ByRef fechaOperacion, ByRef negocioCliente)
    Dim rs, rs1
    
    if (Trim(rsFAC("FCMNCD")) <> "") then auxMoneda = getSimboloMoneda(rsFAC("FCMNCD"))
    
    'Se obtienen los campos del contrato.
    ctoProducto = left(rsFAC("FANCTO"),2)
    ctoSucursal = mid(rsFAC("FANCTO"),3,1)
    ctoOperacion = mid(rsFAC("FANCTO"),4,2)
    ctoNumero = mid(rsFAC("FANCTO"),6,5)
    ctoCosecha = right(rsFAC("FANCTO"),2)	    
    
    'Obtengo el destino, fecha del contrato y Nro negocio cliente.
    Call executeSP(rs, "MERFL.MER311F1_GET_BY_CONTRATO", ctoProducto & "||" & ctoSucursal & "||" & ctoOperacion & "||" & ctoNumero & "||" & ctoCosecha)
    dsDestino = "# ERROR #"
    fechaOperacion = "# ERROR #"
    if (not rs.eof) then 
        negocioCliente = rs("CONCR1")
        fechaOperacion = rs("FCCTR1")
        Call executeSP(rs1, "TOEPFERDB.TBLAFIPPUERTOS_GET_BY_IDPUERTOTOEPFER", rs("CDESR1"))
        if (not rs1.eof) then  dsDestino = rs1("DSPUERTO")
    end if        
            
    'Se totalizan los campos del detalle.    
    Call executeSP(rs, "TFFL.TF101F1GET_BY_FDRGNR", rsFac("FCRGNR"))
    precioTN = 0
    cantidad = 0        
    factor = 1
    if (rsFac("FAUQUI") = UNIDAD_QUINTALES) then factor=10            
    while (not rs.eof)
        if (CDbl(rs("FDPREC")) > 0) then precioTN = precioTN + (CDbl(rs("FDPREC"))* factor)        
        if (CDbl(rs("FDCANT")) <> 0) then cantidad = cantidad + CDbl(rs("FDCANT"))
        rs.MoveNext()
    wend
    precioOperacion = CDbl(rsFAC("FASTOT")) - CDbl(rsFAC("FAIMPN"))
    
    'Alicuota de IVA
    Call executeSP(rs, "TFFL.TF102F1_GET_BY_F2RGNR", rsFAC("FCRGNR"))
    if (not rs.eof) then
        alicuota = CDbl(rs("F2IIPR"))
        if (alicuota = 0) then alicuota = CDbl(rs("F2INPR"))
        'Si sigue en cero, la calculo con los importes.
        if (alicuota = 0) then alicuota = round((CDbl(rsFAC("FAIVAI")) / CDbl(precioTotal)) * 100, 2)        
    end if
End Function
'-----------------------------------------------------------------------------------------------------------------------------
Function dibujarCabeceraLS(ByRef pOpdf, pFechaFac, pTipoFac,pPntVenta, pIsProforma)
    Call GF_writeImage(pOpdf, Server.MapPath("..\Images\afip_C1116B.png"),20, 15, 200, 75, 0)
    Call GF_setFont(pOpdf,"ARIAL", 8 , FONT_STYLE_NORMAL)
    Call GF_writeTextAlign(pOpdf,300,25, GF_FN2DTE(pFechaFac) &", VICENTE LOPEZ" , 260,PDF_ALIGN_RIGHT)
    Call GF_setFont(pOpdf,"ARIAL", 12 , FONT_STYLE_BOLD)
    if ((Cdbl(pTipoFac) = FAC_LIQUIDACION_SECUNDARIA_CREDITO)or(Cdbl(pTipoFac) = FAC_LIQUIDACION_SECUNDARIA_DEBITO))then
        Call GF_writeTextAlign(pOpdf,300,40, "LIQUIDACIÓN SECUNDARIA" , 260,PDF_ALIGN_RIGHT)
        Call GF_writeTextAlign(pOpdf,300,55, "AJUSTE POR COE" , 260,PDF_ALIGN_RIGHT)
    else
        Call GF_writeTextAlign(pOpdf,300,40, "LIQUIDACIÓN SECUNDARIA DE GRANOS" , 260,PDF_ALIGN_RIGHT)
    end if
    if (pIsProforma) then
	    Call GF_writeImage(pOPDF, Server.MapPath("..\Images\facturas\MarcaAguaProforma.gif"),100, 355, 374, 373, 0)
	end if
End Function
'-----------------------------------------------------------------------------------------------------------------------------
Function dibujarIntervinientesLS(ByRef pOpdf, ByRef p_Y, pCOE, pNroCliente, pTipoFac, pPntVenta, pCdCorredor)
    'Dibujo la Actividad del vendedor y el COE
    Dim rs, vDatosProv,auxRazonSocial,auxDomicilio,auxLocalidad
    Call GF_setFont(pOpdf,"ARIAL", 10 , FONT_STYLE_NORMAL)
    Call GF_squareBox(pOpdf,20,85,550 ,30,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_NORMAL)
    Call GF_writeTextAlign(pOpdf,25,90, "Actividad Vendedor: EXPORTADOR" , 260,PDF_ALIGN_LEFT)
    pxComp = 315    
    Call GF_writeTextAlign(pOpdf,25,102, "C.O.E.: " & pCOE , 260,PDF_ALIGN_LEFT)
    'Dibujo los recuadros del comprador y vendedor
    Call GF_setFont(pOpdf,"ARIAL", 10 , FONT_STYLE_BOLD)
    Call GF_squareBox(pOpdf,20,125,260,18,0,"#E6E6E6",NEGRO,1,PDF_SQUARE_NORMAL)
    Call GF_writeTextAlign(pOpdf,20,130, "COMPRADOR" , 260,PDF_ALIGN_CENTER)
    Call GF_squareBox(pOpdf,20,143,260,70,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_NORMAL)

    Call GF_squareBox(pOpdf,310,125,260,18,0,"#E6E6E6",NEGRO,1,PDF_SQUARE_NORMAL)
    Call GF_writeTextAlign(pOpdf,310,130, "VENDEDOR" , 260,PDF_ALIGN_CENTER)
    Call GF_squareBox(pOpdf,310,143,260,70,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_NORMAL)
    'Dibujo los datos del comprador y vendedor    
	vDatosProv = obtenerDatosCompradorLocal(CD_TOEPFER)
	Call dibujarProveedor(pOpdf, vDatosProv, 315)
    vDatosProv = obtenerDatosCompradorLocal(pNroCliente)
    Call dibujarProveedor(pOpdf, vDatosProv, 25)
    'Se detemina si participa el corredor.    
    p_Y = 220
    if (Cdbl(pCdCorredor) <> 0) then
        vDatosCorr = obtenerDatosProveedorLS(pCdCorredor)
        Call GF_writeTextAlign(pOpdf,25,p_Y, "Actuó Corredor: Si", 255,PDF_ALIGN_LEFT)
        p_Y = p_Y + 10
	    Call GF_writeTextAlign(pOpdf,25,p_Y, "C.U.I.T.: " & vDatosCorr(3) , 255,PDF_ALIGN_LEFT)
        Call GF_writeTextAlign(pOpdf,315,p_Y, "Razon Social: " & vDatosCorr(0) , 255,PDF_ALIGN_LEFT)
        p_Y = p_Y + 10
        Call GF_writeTextAlign(pOpdf,25,p_Y, "Ingresos Brutos: " & vDatosCorr(5) , 255,PDF_ALIGN_LEFT)
        Call GF_writeTextAlign(pOpdf,315,p_Y, "COE Original: " , 255,PDF_ALIGN_LEFT)
    else
        Call GF_writeTextAlign(pOpdf,25,p_Y, "Actuó Corredor: No", 255,PDF_ALIGN_LEFT)
    end if
    p_Y = p_Y + 15
End Function
'-----------------------------------------------------------------------------------------------------------------------------
Function dibujarProveedor(ByRef pOpdf, vDatosProv, pX)
    Call GF_setFont(oPDF,"ARIAL", 8 , FONT_STYLE_NORMAL)
    auxRazonSocial = Trim(vDatosProv(0))
    'if (Len(auxRazonSocial) > 46) then auxRazonSocial = Left(auxRazonSocial,45) & ".."
    if (Len(auxRazonSocial) > 38) then Call GF_setFont(oPDF,"ARIAL", 7 , FONT_STYLE_NORMAL)    
    Call GF_writeTextAlign(pOpdf,pX,150, "Razon Social: " & auxRazonSocial , 255,PDF_ALIGN_LEFT)
    if (Len(auxRazonSocial) > 38) then Call GF_setFont(oPDF,"ARIAL", 8 , FONT_STYLE_NORMAL)    
    auxDomicilio = Trim(vDatosProv(2))
    if (Len(auxDomicilio) > 50) then auxDomicilio =  Left(auxDomicilio,48) & ".."
    Call GF_writeTextAlign(pOpdf,pX,160, "Domicilio: "& auxDomicilio , 255,PDF_ALIGN_LEFT)
    auxLocalidad = Trim(vDatosProv(6))
    if (Len(auxLocalidad) > 50) then auxLocalidad =  Left(auxLocalidad,47) & ".."    
    Call GF_writeTextAlign(pOpdf,pX,170, "Localidad: "& auxLocalidad , 255,PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(pOpdf,pX,180, "C.U.I.T.: "& vDatosProv(1) , 255,PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(pOpdf,pX,190, "I.V.A.: "& vDatosProv(3) , 255,PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(pOpdf,pX,200, "Ingresos Brutos N° "&vDatosProv(5) , 255,PDF_ALIGN_LEFT)
	'Call GF_writeTextAlign(pOpdf,pxComp,150, "Razon Social: Alfred C. Toepfer International Argentina S.R.L." , 255,PDF_ALIGN_LEFT)
    'Call GF_writeTextAlign(pOpdf,pxComp,160, "Domicilio: Av. del Libertador 350 10º piso" , 255,PDF_ALIGN_LEFT)
    'Call GF_writeTextAlign(pOpdf,pxComp,170, "Localidad: Vicente Lopez" , 255,PDF_ALIGN_LEFT)
    'Call GF_writeTextAlign(pOpdf,pxComp,180, "C.U.I.T.: 30-62197317-3" , 255,PDF_ALIGN_LEFT)
    'Call GF_writeTextAlign(pOpdf,pxComp,190, "I.V.A.: RI" , 255,PDF_ALIGN_LEFT)
    'Call GF_writeTextAlign(pOpdf,pxComp,200, "Ingresos Brutos N° 30-62197317-3/901" , 255,PDF_ALIGN_LEFT)
End Function
'-----------------------------------------------------------------------------------------------------------------------------
Function dibujarCondicionesOperacionLS(ByRef pOpdf, ByRef pY, pMoneda, pPrecioTn, pFecha, pCdProducto, pDsPuerto)
    Dim auxPuerto, auxDsProducto, rs    
    
    'Se obtienen las descripciones de producto
    Call executeSP(rs, "MERFL.MER112F1_GET_BY_CODIPR", pCdProducto)
    auxDsProducto = "# ERROR #"
    if (not rs.eof) then auxDsProducto = rs("DESCPR")
    
    'Dibujo la seccion
    Call GF_squareBox(pOpdf,20,pY,400,18,0,"#E6E6E6",NEGRO,1,PDF_SQUARE_NORMAL)
    Call GF_setFont(oPDF,"ARIAL", 10 , FONT_STYLE_BOLD)
    Call GF_writeTextAlign(pOpdf,22,pY + 5, "CONDICIONES DE LA OPERACIÓN" , 400,PDF_ALIGN_LEFT)
    Call GF_setFont(oPDF,"ARIAL", 8 , FONT_STYLE_BOLD)
    Call GF_squareBox(pOpdf,420,pY,150,18,0,"#E6E6E6",NEGRO,1,PDF_SQUARE_NORMAL)
    Call GF_writeTextAlign(pOpdf,422,pY + 6, "Fecha: " & GF_FN2DTE(pFecha) , 150,PDF_ALIGN_LEFT)
    pY = pY + 18    
    Call GF_squareBox(pOpdf,20,pY,120,15,0,"#E6E6E6",NEGRO,1,PDF_SQUARE_NORMAL)
    Call GF_squareBox(pOpdf,140,pY,240,15,0,"#E6E6E6",NEGRO,1,PDF_SQUARE_NORMAL)
    Call GF_squareBox(pOpdf,380,pY,190,15,0,"#E6E6E6",NEGRO,1,PDF_SQUARE_NORMAL)
    Call GF_writeTextAlign(pOpdf,20,pY + 4, "Precio/Tn" , 120,PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(pOpdf,140,pY + 4, "Grano" , 240,PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(pOpdf,380,pY + 4, "Puerto" , 190,PDF_ALIGN_CENTER)
    pY = pY + 15
    Call GF_setFont(oPDF,"ARIAL", 8 , FONT_STYLE_NORMAL)
    Call GF_squareBox(pOpdf,20,pY,120,15,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_NORMAL)
    Call GF_squareBox(pOpdf,140,pY,240,15,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_NORMAL)
    Call GF_squareBox(pOpdf,380,pY,190,15,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_NORMAL)
    Call GF_writeTextAlign(pOpdf,20,pY + 4, pMoneda & " " & GF_EDIT_DECIMALS(Cdbl(pPrecioTn)*100, 2), 120,PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(pOpdf,140,pY + 4, pCdProducto &" - "& auxDsProducto , 240,PDF_ALIGN_CENTER)
    auxPuerto = pDsPuerto
    if (Len(auxPuerto) > 43) then auxPuerto = Left(auxPuerto,42) & ".."
    Call GF_writeTextAlign(pOpdf,380,pY + 4, auxPuerto , 190,PDF_ALIGN_CENTER)
    pY = pY + 30
End Function
'-----------------------------------------------------------------------------------------------------------------------------
Function dibujarOperacionLS(ByRef pOpdf, ByRef pY, pTipoFac, pCantidad, pPrecioUnidad, pSubtotal, pAlicuota, pImporteIVA, pDeducciones, pPercepcionIVA, pPercepcionIBB, pImporteNeto, pMoneda)
    Dim auxSubTotal
    Call GF_squareBox(pOpdf,20,pY,550,18,0,"#E6E6E6",NEGRO,1,PDF_SQUARE_NORMAL)
    Call GF_setFont(oPDF,"ARIAL", 10 , FONT_STYLE_BOLD)
    Call GF_writeTextAlign(pOpdf,22,pY + 5, "OPERACIÓN" , 400,PDF_ALIGN_LEFT)
    
    pY = pY + 18
    Call GF_setFont(oPDF,"ARIAL", 8 , FONT_STYLE_BOLD)
    Call GF_squareBox(pOpdf,20,pY,70,15,0,"#E6E6E6",NEGRO,1,PDF_SQUARE_NORMAL)
    Call GF_squareBox(pOpdf,90,pY,70,15,0,"#E6E6E6",NEGRO,1,PDF_SQUARE_NORMAL)
    Call GF_squareBox(pOpdf,160,pY,70,15,0,"#E6E6E6",NEGRO,1,PDF_SQUARE_NORMAL)
    Call GF_squareBox(pOpdf,220,pY,60,15,0,"#E6E6E6",NEGRO,1,PDF_SQUARE_NORMAL)
    Call GF_squareBox(pOpdf,280,pY,70,15,0,"#E6E6E6",NEGRO,1,PDF_SQUARE_NORMAL)
    Call GF_squareBox(pOpdf,350,pY,80,15,0,"#E6E6E6",NEGRO,1,PDF_SQUARE_NORMAL)
    Call GF_squareBox(pOpdf,430,pY,70,15,0,"#E6E6E6",NEGRO,1,PDF_SQUARE_NORMAL)
    Call GF_squareBox(pOpdf,500,pY,70,15,0,"#E6E6E6",NEGRO,1,PDF_SQUARE_NORMAL)
    Call GF_writeTextAlign(pOpdf,20 ,pY + 4, "Cantidad" , 70,PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(pOpdf,90 ,pY + 4, "Precio/Tn" , 70,PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(pOpdf,160,pY + 4, "Subtotal" , 70,PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(pOpdf,220,pY + 4, "% Alicuota" , 60,PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(pOpdf,280,pY + 4, "Importe IVA" , 70,PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(pOpdf,350,pY + 4, "Operacion c/IVA" , 80,PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(pOpdf,430,pY + 4, "Deducciones" , 70,PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(pOpdf,500,pY + 4, "Percepciones" , 70,PDF_ALIGN_CENTER)
    pY = pY + 15
    Call GF_setFont(oPDF,"ARIAL", 8 , FONT_STYLE_NORMAL)
    Call GF_squareBox(pOpdf,20,pY,70,15,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_NORMAL)
    Call GF_squareBox(pOpdf,90,pY,70,15,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_NORMAL)
    Call GF_squareBox(pOpdf,160,pY,70,15,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_NORMAL)
    Call GF_squareBox(pOpdf,220,pY,60,15,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_NORMAL)
    Call GF_squareBox(pOpdf,280,pY,70,15,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_NORMAL)
    Call GF_squareBox(pOpdf,350,pY,80,15,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_NORMAL)
    Call GF_squareBox(pOpdf,430,pY,70,15,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_NORMAL)
    Call GF_squareBox(pOpdf,500,pY,70,15,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_NORMAL)
    Call GF_writeTextAlign(pOpdf,20 ,pY + 4, GF_EDIT_DECIMALS(pCantidad, 3) &" Tn", 70,PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(pOpdf,90 ,pY + 4, pMoneda &" "& GF_EDIT_DECIMALS(Cdbl(pPrecioUnidad)*100, 2) , 70,PDF_ALIGN_CENTER)    
    Call GF_writeTextAlign(pOpdf,160,pY + 4, pMoneda &" "& GF_EDIT_DECIMALS(Cdbl(pSubtotal)*100, 2) , 60,PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(pOpdf,220,pY + 4, GF_EDIT_DECIMALS(Cdbl(pAlicuota)*100, 2), 60,PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(pOpdf,280,pY + 4, pMoneda &" "& GF_EDIT_DECIMALS(Cdbl(pImporteIVA)*100 , 2), 70,PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(pOpdf,350,pY + 4, pMoneda &" "& GF_EDIT_DECIMALS((Cdbl(pImporteIVA)+Cdbl(pSubtotal))*100,2) , 80,PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(pOpdf,430,pY + 4, pMoneda &" "& GF_EDIT_DECIMALS(Cdbl(pDeducciones)*100 , 2) , 70,PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(pOpdf,500,pY + 4, pMoneda &" "& GF_EDIT_DECIMALS((Cdbl(pPercepcionIVA)+Cdbl(pPercepcionIBB))*100,2) , 70,PDF_ALIGN_CENTER)
    pY = pY + 30
    Call GF_squareBox(pOpdf,147, pY, 300,20,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_NORMAL)
    Call GF_setFont(pOpdf,"ARIAL", 10 , FONT_STYLE_BOLD)
    Call GF_writeTextAlign(pOpdf, 150, py + 5, "IMPORTE NETO LIQUIDACIÓN: " , 300, PDF_ALIGN_LEFT)
    Call GF_setFont(pOpdf,"ARIAL", 8 , FONT_STYLE_NORMAL)
    Call GF_writeTextAlign(pOpdf, 300, py + 6,  pMoneda &" "& GF_EDIT_DECIMALS(Cdbl(pImporteNeto)*100,2), 145, PDF_ALIGN_CENTER)
    pY = pY + 30
End Function
'-----------------------------------------------------------------------------------------------------------------------------
Function dibujarPiePaginaLS(ByRef pOpdf, ByRef pY, pCliente, pNegocioToepfer, pFechaVto, pNegocioCliente)
    Dim rsDet, auxVCto, strLine
    
    Call GF_setFont(pOpdf,"ARIAL", 9 , FONT_STYLE_NORMAL)
    Call GF_writeTextAlign(pOpdf, 20, py, "Datos Adicionales: " , 80, PDF_ALIGN_LEFT)
    if (CLng(pCliente) = PROVEEDOR_ESPECIAL_MAT) then
        py = GF_writeTextAlign(pOpdf, 100, py, TEXTO_CTO_MAT & pNegocioCliente , 400, PDF_ALIGN_LEFT)    
    else
        'Nuestro Contrato
        Call GF_writeTextAlign(pOpdf, 100, py, "Nuestro Cto:" & pNegocioToepfer , 400, PDF_ALIGN_LEFT)            
        'Dibujo el contrato de terceros.
        py=py+12
        Call GF_writeTextAlign(pOpdf, 100, py, TEXTO_CTO_GRAL & pNegocioCliente , 400, PDF_ALIGN_LEFT)    
        'Fecha Vto Factura
        py=py+12
        Call GF_writeTextAlign(pOpdf, 100, py, "Fecha Vto: " & GF_FN2DTE(pFechaVto), 400, PDF_ALIGN_LEFT)        
	end if	    
    
End Function
'-----------------------------------------------------------------------------------------------------------------------------
'Esta funcion obtiene los datos de un comprador en caso que sea Liquidacion secundaria de granos o del vendedor en caso
' que sea Liquidacion secundaria por ajuste de COE
'Function obtenerDatosProveedorLS(pNroPro)
'	Dim rs,conn,strSQL,rtrn()
'	redim rtrn(6)	
'	strSQL = "select NOMAMP,NRODOC,DOMICI,CODPOS,LOCALI,CODIVA,NROIBR from MERFL.TCB6A1F1 where NROPRO = " & pNroPro
'	Call executeQuery(rs, "OPEN", strSQL)	
'	if (not rs.EoF) then
'	    rtrn(0) = Trim(rs("NOMAMP"))
 '       strSQL="Select DESCR1,NDOCR1 from DGI.DGI600F1 where NDOCR1='" & rs("NRODOC") & "'"
  '      Call executeQuery(rs1, "OPEN", strSQL)		
'	    if (not rs1.eof) then rtrn(0) = Trim(rs1("DESCR1"))
 '       rtrn(1) = Trim(rs("DOMICI")) 
  '      rtrn(2) = Trim(rs("LOCALI"))
'		rtrn(3) = GF_STR2CUIT(rs("NRODOC"))
'		'Condicion frente al IVA.
'	    strSQL = "Select DESCR1 from DGI.DGI601F1 where CDIMR1='" & rs("CODIVA") & "'"
'	    Call executeQuery(rs1, "OPEN", strSQL)
'	    rtrn(4) = "ERROR - IVA"
'	    if (not rs1.eof) then rtrn(4) = rs1("DESCR1")
'       rtrn(5) = rs("NROIBR")
'	end if
'	obtenerDatosProveedorLS = rtrn
'End Function


%>