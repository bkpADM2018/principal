<!--#include file="../Includes/procedimientosCompras.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientosFechas.asp"-->
<!--#include file="../Includes/procedimientosTraducir.asp"-->
<!--#include file="../Includes/procedimientosMath.asp"-->
<!--#include file="../Includes/procedimientosPDF.asp"-->
<!--#include file="interfacturas.asp"-->
<!--#include file="interfacturasPrintSTD.asp"-->
<!--#include file="interfacturasPrintGP.asp"-->
<!--#include file="interfacturasPrintEX.asp"-->
<%
Const MAX_LINEAS_PAGINA          = 29		'Cant total de lineas que entran por pagina.
Const MAX_LINEAS_PAGINA_AUXILIAR = 63		'Cant total de lineas que entran por pagina.
Const LINEAS_DE_TOTALES = 7			'Cant de lineas para imprimir totales
Const OBS_EOL_TOKEN 	= "|"
Const TEXTO_CTO_GRAL 	= "V/Ctro."
Const TEXTO_CTO_MAT 	= "Caratula"
Const PROVEEDOR_ESPECIAL_MAT = 5454
Const UNIDAD_QUINTALES 		 = "Q"
Const PROVINCIA_IIBB_CABA 	 = 24
Const PROVINCIA_IIBB_BA 	 = 1
Const PROVINCIA_IIBB_SF 	 = 20
'---------------------------------------------------------------------------------------------------------------------
'--------------------FUNCIONES GENERICAS UTILIZADAS PARA FACTURAS LOCALES COMO EXPORTACION----------------------------
'---------------------------------------------------------------------------------------------------------------------
Function obtenerCodigoBarras(pNroCAE,pFecVenCAE,pPuntoVenta, pTipoCbte, pLetra)
	Dim rtrn, codigoFAC
	codigoFAC = getCodigoCbteAFIP(pTipoCbte, pLetra)
	if (not isNumeric(codigoFAC)) then codigoFAC = "00"		
	codigoFAC = GF_nDigits(codigoFAC,2)
	rtrn = CUIT_TOEPFER & codigoFac & pPuntoVenta & trim(pNroCAE) & trim(pFecVenCAE)
	obtenerCodigoBarras = obtenerDigitoVerificador(rtrn)
End Function
'------------------------------------------------------------------------------------------------------------------------
Function obtenerDatosCompradorLocal(pNroPro)
	Dim rs,conn,strSQL,rtrn()
	redim rtrn(7)	
	strSQL = "select alaord NOMAMP, nomemp RAZSOC, NRODOC, domemp DOMICI, numero, piso, oficina, CODPOS, localidad LOCALI,CODIVA from MET001A where nroemp = " & pNroPro
	Call executeQueryDb(DBSITE_SQL_MAGIC, rs, "OPEN", strSQL)	
	if (not rs.EoF) then	    	    	    
        rtrn(0) = Trim(rs("NOMAMP"))
        if (Len(rtrn(0)) < 3) then rtrn(0) = Trim(rs("RAZSOC"))        
        rtrn(4) = pNroPro	                                
		rtrn(1) = GF_STR2CUIT(rs("NRODOC"))
		rtrn(2) = Trim(rs("DOMICI")) & " " & trim(rs("numero"))
		if (rs("piso") <> "") and (rs("piso") <> "0") then rtrn(2) = rtrn(2) & " " & Trim(rs("piso")) & "°"
		if (rs("oficina") <> "") then rtrn(2) = rtrn(2) & " " & Trim(rs("oficina"))
		'Condicion frente al IVA.
	    strSQL = "Select desiva from MET039A where codiva=" & rs("CODIVA")
	    Call executeQueryDb(DBSITE_SQL_MAGIC, rs1, "OPEN", strSQL)
	    rtrn(3) = "ERROR - IVA"
	    if (not rs1.eof) then rtrn(3) = rs1("desiva")	
	    'Nro IIBB
        rtrn(5) = "0"        
	    if (CLng(pNroPro) = CLng(CD_TOEPFER)) then rtrn(5) = "30-62197317-3/901"
	    rtrn(6) = rs("LOCALI")
		rtrn(7) = rs("CODPOS")
	end if
	obtenerDatosCompradorLocal = rtrn
End Function
'----------------------------------------------------------
' Mod: 2017-09-26 - JAS
Function dibujarOrigenDestinoLocal(pNroCliente,ByRef pOpdf)
    Dim datosComprador    
    datosComprador = obtenerDatosCompradorLocal(pNroCliente)    
    'Estructura
    Call GF_squareBox(pOpdf,3,150,587 ,50,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_ROUND) 
	'Call GF_writeImage(pOpdf, Server.MapPath("..\Images\facturas\MarcaAguaToepfer.gif"),5, 152, 570, 45, 0)	    
    'Datos
	Call GF_setFont(pOpdf,"ARIAL", 8 , FONT_STYLE_NORMAL)
	Call GF_writeTextAlign(pOpdf,15,155, "SEÑORES:" , 200,PDF_ALIGN_LEFT)	
	Call GF_writeTextAlign(pOpdf,400,155, "C.U.I.T.:" , 200,PDF_ALIGN_LEFT)	
	Call GF_writeTextAlign(pOpdf,15,185, "DOMICILIO:", 200,PDF_ALIGN_LEFT)	
	Call GF_writeTextAlign(pOpdf,400,185, "I.V.A.:" , 200,PDF_ALIGN_LEFT)	
	'muestra de datos
	Call GF_setFont(pOpdf,"ARIAL", 8 , FONT_STYLE_BOLD)
	Call GF_writeTextAlign(pOpdf, 65,155, Trim(datosComprador(0)) & " ("& datosComprador(4) &")", 200,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(pOpdf, 65,185, datosComprador(2) & " - C.P.:" & datosComprador(7) & " " & datosComprador(6)		, 200,PDF_ALIGN_LEFT)	
	Call GF_writeTextAlign(pOpdf,440,155, datosComprador(1), 200,PDF_ALIGN_LEFT)	
	Call GF_writeTextAlign(pOpdf,425,185, datosComprador(3), 200,PDF_ALIGN_LEFT)	
	'DATO EXCLUSIVO PARA NAVIERA CHACO
	if (CLng(datosComprador(4)) = 18412) then
	    Call GF_writeTextAlign(pOpdf,65,170, "RUC 80001089-2" , 200,PDF_ALIGN_LEFT)	
	end if
End Function
'----------------------------------------------------------
' Mod: 2017-09-29 - JAS
Function dibujarDatosCompraLocal(pNroRegFac,ByRef pOpdf)
	Dim strSQL, rs
	strSQL = "Select Format(fecvto, 'dd/MM/yyyy') fvto from FAC001A where guid='" & pNroRegFac & "'"
	Call executeQueryDb(DBSITE_SQL_MAGIC, rs, "OPEN", strSQL)	
	
    Call GF_squareBox(pOpdf, 3, 200, 587, 50,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_ROUND) 
    'Call GF_squareBox(pOpdf,296, 200, 177 ,50,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_ROUND)
    'Call GF_squareBox(pOpdf,473, 200, 117 ,50,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_ROUND)
    'Datos fijos
	Call GF_setFont(pOpdf,"ARIAL", 6 , FONT_STYLE_NORMAL)
	Call GF_writeTextAlign(pOpdf, 10, 205, "OPERACIÓN:" , 150, PDF_ALIGN_LEFT)	
	'Call GF_writeTextAlign(pOpdf,305, 205, "O/COMPRA:"       , 200, PDF_ALIGN_LEFT)	
	Call GF_writeTextAlign(pOpdf,480, 205, "FECHA VTO.:"      , 200, PDF_ALIGN_LEFT)		
	'if (datosCompra(3) <> "")then Call GF_writeTextAlign(pOpdf, 160, 205, "FECHA VTO.:" , 100, PDF_ALIGN_LEFT)
	'Datos variables
	if (not rs.eof) then
		Call GF_setFont(pOpdf,"ARIAL", 10 , FONT_STYLE_BOLD)
		'Call GF_writeTextAlign(pOpdf, 70, 210, datosCompra(0), 117, PDF_ALIGN_LEFT)	
		'Call GF_writeTextAlign(pOpdf,170, 210, datosCompra(3), 117, PDF_ALIGN_CENTER)		
		'Call GF_writeTextAlign(pOpdf,296, 210, datosCompra(1), 177, PDF_ALIGN_CENTER)		
		Call GF_writeTextAlign(pOpdf,480, 210, rs("fvto"), 117, PDF_ALIGN_CENTER)
	end if
End Function
'--------------------------------------------------------------------------------------------------------------------
' Mod: 2017-10-03 - JAS
Function dibujarDetalleTitulosLocal(ByRef pOPDF, pIsProforma)
	Call GF_squareBox(pOPDF,  3, 250,  587, 15,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_ROUND) 	
	Call GF_verticalLine(pOPDF, 497, 250, 15)
	Call GF_setFont(pOPDF,"ARIAL", 10 , FONT_STYLE_NORMAL)
	Call GF_writeTextAlign(pOPDF,   3, 253, "DETALLE", 451, PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(pOPDF, 500, 253, "IMPORTE",  90, PDF_ALIGN_CENTER)	
End Function
'--------------------------------------------------------------------------------------------------------------------
Function dibujarPieUltimaPaginaLocal(pCAE, pFecVto, ByRef pOpdf,pCurrPage, pTotalPages, pLetraFAC, pIdioma, pPuntoVenta, rsFAC)
    Dim strLeyenda, arr, ol, params

    'Codigo de Barras
    Call drawCodeBar(obtenerCodigoBarras(pCAE,pFecVto,pPuntoVenta, rsFac("FCCMTP"), pLetraFAC),15,700,40)

    'Observaciones    
    if (g_Observaciones <> "") then 
        Call GF_setFont(pOpdf,"ARIAL", 7 , FONT_STYLE_NORMAL)
        'Se imprimen todas las lineas.        
        Call GF_writeTextPlus(pOpdf, 5, 750 + ol, g_Observaciones, 260, 8, PDF_ALIGN_LEFT)        
    end if    
    'Si es en dolares se imrpime el tipo de cambio
    if (rsFac("FCMNCD") = MONEDA_DOLAR_FACTURACION) then
        Call GF_setFont(pOpdf,"ARIAL", 8 , FONT_STYLE_BOLD)
        Call GF_writeTextAlign(pOpdf, 275, 810, GF_TRADUCIR("Tipo de Cambio:") & " $" & rsFAC("FCCBTP") , 300,PDF_ALIGN_LEFT)
        Call GF_setFont(pOpdf,"ARIAL", 8 , FONT_STYLE_NORMAL)
    end if		            
	
End Function
'--------------------------------------------------------------------------------------------------------------------
' Mod: 2017-10-04 - JAS
Function dibujarTotalesLocal(rsFac, tasaIVA, ByRef pOpdf)
    Dim strSQL, rs, Xo, Yo, aux
    Dim cdMoneda, myGravado, myNoGravado, myIVA, myPercepcionIVA, myPercepIIBB_BA, myPercepIIBB_SF, myPercepIIBB_CABA, myTasaIVA, myTotal
        
	strSQL="Select * from FAC001C where guid='" & rsFac("guid") & "'"	
	Call executeQueryDb(DBSITE_SQL_MAGIC, rs, "OPEN", strSQL)
	myPercepIIBB_BA = 0
	myPercepIIBB_SF = 0
	myPercepIIBB_CABA = 0
	while (not rs.eof)
		Select case CInt(rs("codprv"))
			case PROVINCIA_IIBB_CABA
				myPercepIIBB_CABA = myPercepIIBB_CABA + CDbl(rs("importe"))
			case PROVINCIA_IIBB_BA
				myPercepIIBB_BA = myPercepIIBB_BA + CDbl(rs("importe"))
			case PROVINCIA_IIBB_SF
				myPercepIIBB_SF = myPercepIIBB_SF + CDbl(rs("importe"))
		end select
		rs.MoveNext()
	wend
	if (CDbl(rsFAC("no_gravado")) = 0) 	then myNoGravado = "" 		else myNoGravado 		= FormatNumber(rsFAC("no_gravado")	, 2, 0, 0, -1)
	if (CDbl(rsFAC("gravado")) = 0) 	then myGravado = "" 		else myGravado 			= FormatNumber(rsFAC("gravado")		, 2, 0, 0, -1)
	if (CDbl(rsFAC("impIVAcbt")) = 0) 	then myIVA = "" 			else myIVA 				= FormatNumber(rsFAC("impIVAcbt")	, 2, 0, 0, -1)
	if (CDbl(rsFAC("percepIVA")) = 0) 	then myPercepcionIVA = "" 	else myPercepcionIVA 	= FormatNumber(rsFAC("percepIVA")	, 2, 0, 0, -1)
	if (myPercepIIBB_BA = 0) 			then myPercepIIBB_BA   = "" else myPercepIIBB_BA   	= FormatNumber(myPercepIIBB_BA  	, 2, 0, 0, -1)
	if (myPercepIIBB_SF = 0) 			then myPercepIIBB_SF   = "" else myPercepIIBB_SF   	= FormatNumber(myPercepIIBB_SF  	, 2, 0, 0, -1)
	if (myPercepIIBB_CABA = 0) 			then myPercepIIBB_CABA = "" else myPercepIIBB_CABA 	= FormatNumber(myPercepIIBB_CABA	, 2, 0, 0, -1)	
	Call GF_setFont(pOpdf,"ARIAL", 10 , FONT_STYLE_NORMAL)    
    'Textos Fijos
	Yo = 610
	Xo = 390
    Call GF_writeTextAlign(pOpdf, Xo,     Yo, "No Gravado"				, 250,PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(pOpdf, Xo,  Yo+12, "Gravado"					, 250,PDF_ALIGN_LEFT)
	aux = "IVA Inscr."
	if (CDbl(tasaIVA) > 0) then aux = aux & " (" & tasaIVA & "%)"
    Call GF_writeTextAlign(pOpdf, Xo,  Yo+24, aux 						, 250,PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(pOpdf, Xo,  Yo+36, "Percepcion IVA "			, 250,PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(pOpdf, Xo,  Yo+48, "Percepcion IIBB Bs. As."	, 250,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(pOpdf, Xo,  Yo+60, "Percepcion IIBB Sta Fe"	, 250,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(pOpdf, Xo,  Yo+72, "Percepcion IIBB CABA"	, 250,PDF_ALIGN_LEFT)
    'Valores
    Call GF_writeTextAlign(pOpdf, 454,    Yo, myNoGravado		, 130, PDF_ALIGN_RIGHT)
    Call GF_writeTextAlign(pOpdf, 454, Yo+12, myGravado			, 130, PDF_ALIGN_RIGHT)
    Call GF_writeTextAlign(pOpdf, 454, Yo+24, myIVA				, 130, PDF_ALIGN_RIGHT)
    Call GF_writeTextAlign(pOpdf, 454, Yo+36, myPercepcionIVA	, 130, PDF_ALIGN_RIGHT)
    Call GF_writeTextAlign(pOpdf, 454, Yo+48, myPercepIIBB_BA	, 130, PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(pOpdf, 454, Yo+60, myPercepIIBB_SF	, 130, PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(pOpdf, 454, Yo+72, myPercepIIBB_CABA	, 130, PDF_ALIGN_RIGHT)
	Call GF_setFont(pOpdf,"ARIAL", 8, FONT_STYLE_NORMAL)    
    if (Cdbl(rsFAC("imptotcbt")) > 0) then Call GF_writeTextPlus(pOpdf, 10, 750, "Son " & getNombreMoneda(rsFAC("codmone")) & " " & numeroALetras(rsFAC("imptotcbt")), 250, 12, PDF_ALIGN_LEFT)
	
    Call GF_setFont(pOpdf,"ARIAL", 12 , FONT_STYLE_BOLD)
    Call GF_writeTextAlign(pOpdf, 450, Yo+103, "TOTAL", 250,PDF_ALIGN_LEFT)    
	Call GF_setFont(pOpdf,"ARIAL", 10 , FONT_STYLE_BOLD)
	Call GF_writeText(pOpdf, 500, Yo+103, getSimboloMoneda(rsFac("codmone")), 0)
    Call GF_writeTextAlign(pOpdf, 500, Yo+103, FormatNumber(rsFAC("imptotcbt"), 2, 0, 0, -1), 84, PDF_ALIGN_RIGHT)
End Function
'---------------------------------------------------------------------------------------------------------------------
Function obtenerDigitoVerificador(pCodigo)
	'Paso 1: sumar todos los numeros de las posiciones impares.
	'Paso 2: multiplicar el nro del paso 1 por 3.
	'Paso 3: sumar todos los numeros de las posiciones pares.
	'Paso 4: sumar los numeros del paso 2 y paso 3
	'Paso 5: obtener el numero necesario para el el nro del paso 4 sea multiplo de 10	
	Dim paso1,paso2,paso3,paso4,rtrn		
	for i = 1 to len(pCodigo) step 2
		paso1 = paso1 + cint(mid(pCodigo,i,1))
	next	
	paso2 = paso1 * 3 	
	for i = 2 to len(pCodigo) step 2
		paso3 = paso3 + cint(mid(pCodigo,i,1))
	next	
	paso4 = paso2 + paso3
	rtrn = pCodigo & right(cstr(10-right(paso4,1)),1)	
	obtenerDigitoVerificador = rtrn	
End Function
'---------------------------------------------------------------------------------------------------------------------
' Mod: 2017-09-25 - JAS
Function dibujarCabecera(pCia, pNroFac,pLetraFac,pFecha, pTipoFac,pPntVenta,ByRef pOpdf)
    Dim auxCbte, myLetra
	'Estructura
    Call GF_squareBox(pOpdf,3  ,5,291,140,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_ROUND) 
	Call GF_squareBox(pOpdf,294,5,296,140,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_ROUND) 	
	Call GF_squareBox(pOpdf,265,5,60,60,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_NORMAL) 
	Call GF_squareBox(pOpdf,265,75,60,15,0,"#FFFFFF","#FFFFFF",1,PDF_SQUARE_NORMAL) 	
	'logo
	Call GF_writeImage(pOpdf, Server.MapPath("..\Images\ADMlogo2.jpg"),10, 25, 81, 81, 0)
	'Letra Principal
	Call GF_setFont(pOpdf,"ARIAL", 46 , FONT_STYLE_BOLD)
	myLetra = Trim(pLetraFac)
	if (myLetra = "") then myLetra = "E"   'Como expo no tiene letra en la tabla, se completa en este punto.
	Call GF_writeTextAlign(pOpdf,278, 10, myLetra , 200,PDF_ALIGN_LEFT)	
	'Codigo AFIP
	Call GF_setFont(pOpdf,"ARIAL", 8 , FONT_STYLE_NORMAL)
	auxCbte = getCodigoCbteAFIP(pTipoFac, pLetraFac)	
	Call GF_writeTextAlign(pOpdf,265,78, "Código Nº "& auxCbte, 60,PDF_ALIGN_CENTER)	
	'Datos
	Call dibujarCabeceraContable(pNroFac,pFecha, pTipoFac,pOpdf)
	Call dibujatCabeceraEmpresa(pCia, pPntVenta, pFecha, pOpdf)
End Function
'---------------------------------------------------------------------------------------------------------------------
'Dibuja los datos de direccion de la empresa, dependiendo del punto de venta
' Mod: 2017-09-25 - JAS
Function dibujatCabeceraEmpresa(pCia, pPntVenta, pFechaFactura,ByRef pOpdf)
	Call GF_setFont(oPDF,"ARIAL", 8 , FONT_STYLE_NORMAL)	
	strSQL="Select * from FAT001A2 where codcia = '" & pCia & "'"	
	Call executeQueryDb(DBSITE_SQL_MAGIC, rs1, "OPEN", strSQL)
	if (not rs1.EoF) then	
		Call GF_setFont(pOpdf,"ARIAL", 12 , FONT_STYLE_BOLD)
		Call GF_writeTextAlign(pOpdf,90, 30, "ADM AGRO S.R.L." , 200,PDF_ALIGN_LEFT)
		Call GF_setFont(oPDF,"ARIAL", 8 , FONT_STYLE_NORMAL)		
		Call GF_writeTextAlign(pOpdf,90, 60, Trim(rs1("domemp")) , 200,PDF_ALIGN_LEFT)
		Call GF_writeTextAlign(pOpdf,90, 75, Trim(rs1("locacion")), 200,PDF_ALIGN_LEFT)
		Call GF_writeTextAlign(pOpdf,90, 90, Trim(rs1("contacto")), 200,PDF_ALIGN_LEFT)
		Call GF_writeTextAlign(pOpdf,90,130, "I.V.A. RESPONSABLE INSCRIPTO" , 200,PDF_ALIGN_LEFT)
	else
		Call GF_writeTextAlign(pOpdf,90, 80, "ERROR-FALTA DIRECCION PTO VTA" , 200,PDF_ALIGN_LEFT)
	end if
End Function
'---------------------------------------------------------------------------------------------------------------------
' Mod: 2017-09-26 - JAS
Function dibujarCabeceraContable(pNroFac,pFecha, pTipoFac,ByRef pOpdf)
	Dim auxLeyenda, strSQL
				
	Call GF_setFont(pOpdf,"ARIAL", 8 , FONT_STYLE_NORMAL)
	Call GF_writeTextAlign(pOpdf,340, 75, "FECHA: " & GF_FN2DTE(pFecha) , 200,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(pOpdf,340, 90, "C.U.I.T.: 30-62197317-3" , 200,PDF_ALIGN_LEFT)
	if (CLng(pFecha) >= 20151101) then
	    Call GF_writeTextAlign(pOpdf,340, 100, "INGRESOS BRUTOS C.M.: 30-62197317-3/901" , 200,PDF_ALIGN_LEFT)
    else	    
	    Call GF_writeTextAlign(pOpdf,340, 100, "INGRESOS BRUTOS C.M.: 901-937800-6" , 200,PDF_ALIGN_LEFT)
	end if
	Call GF_writeTextAlign(pOpdf,340,110, "CAJA PREV. COMERCIO Nº: 1459316" , 200,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(pOpdf,340,120, "IMPUESTOS INTERNOS: NO RESPONSABLE" , 200,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(pOpdf,340,130, "INICIO DE ACTIVIDADES: 30-12-1987" , 200,PDF_ALIGN_LEFT)	
			
	auxLeyenda = getTipoFactura(pTipoFac)	
	Call GF_setFont(pOpdf,"ARIAL", 20 , FONT_STYLE_NORMAL)
	Call GF_writeTextAlign(pOpdf,340, 12, Ucase(auxLeyenda), 250,PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(pOpdf,340, 50, "Nº " & pNroFac , 200,PDF_ALIGN_LEFT)
		
End Function
'---------------------------------------------------------------------------------------------------------------------
'Obtiene los datos de la cabecera de la Factura de un determinado Numero de Registro
' Mod: 2017-09-25
Function getRSfactura(p_nroReg)
	Dim rsTemp, strSQL	
	strSQL = "select *, format(feccbt, 'yyyyMMdd') feccbt_fn, format(vencai, 'yyyyMMdd') vencai_fn from FAC001A where GUID = '" & p_nroReg & "'"	
	Call executeQueryDb(DBSITE_SQL_MAGIC, rsTemp, "OPEN", strSQL)
	set getRSfactura = rsTemp
End Function
'---------------------------------------------------------------------------------------------------------------------
'Obtiene los datos de la cabecera de la Factura de un determinado Numero de Registro
' Mod: 2017-09-25
Function getRSfacturaDetalle(p_nroReg)
	Dim rsTemp, strSQL	
	strSQL = "select * from FAC001B where GUID = '" & p_nroReg & "' and codconce < 900"	
	Call executeQueryDb(DBSITE_SQL_MAGIC, rsTemp, "OPEN", strSQL)
	set getRSfacturaDetalle = rsTemp
End Function
'----------------------------------------------------------------------------------------------------------------------
' Mod: 2017-09-25 - JAS
Function getParamsLeyenda(rs)
    Dim arr() 
    Redim arr(1)
    'if (rs("LDLYCD") = FAC_CODIGO_CONCEPTO_GP) then
        if (CInt(rs("codMone")) = MONEDA_PESO_NUMERICO) then
            Redim arr(2)
            arr(0) = GF_EDIT_DECIMALS((CDbl(rs("imptotcbt")) * 100) / CDbl(rs("tcfin")), 2)
            arr(1) = CDbl(rs("tcfin"))
        end if
    'end if
    getParamsLeyenda = arr
End Function
'----------------------------------------------------------
' Mod: 2017-09-25 - JAS
'Obtiene la leyenda del detalle de un determinado registro, para saber cual será la leyenda se le pasará algunos 
'datos de la cabecera de la factura:
'	* Letra y tipo comprobante, moneda, idioma (default español),secuencia por dafault la menor(1) 
'Para los parametros se espera un array con tantas posiciones como parametros tengan que reemplzarse en la leyenda.
' En los textos los parametros se buscan siguiendo el patron <?>. Los parametros deben ser cargados en el orden en que se reemplzaran en el texto.
Function getDocLeyendaFactura(pIdDocumento, pTipo, pLetra, pConcepto, pCia, pMoneda, pFecha, pParams)
	Dim strLeyenda,auxIdioma,rsLeyenda,strSQL,auxSecuencia, i	
		
	strSQL = " SELECT CASE WHEN LEYENDA IS NULL THEN '' ELSE RTRIM(LEYENDA) END AS LEYENDA FROM FAT001A3 "&_
			 " WHERE CODMONE = " & pMoneda &_
			 "      AND RECNO in (0, " & pIdDocumento & " )" &_			 
             "      AND TIPCBT in (0, " & pTipo & " )" &_
             "      AND LETRA in ('', '" & pLetra & "')" &_
             "      AND CODCONCE in (0, " & pConcepto & ")" &_             
			 "      AND CODCIA in ('', '" & pCia & "')" &_             
             "      AND FECVIGDESDE <= '" & pFecha & "'" &_
			 " ORDER BY RECNO DESC, CODCONCE DESC, TIPCBT DESC, LETRA DESC, CODCIA DESC, FECVIGDESDE desc"
	Call executeQueryDb(DBSITE_SQL_MAGIC, rsLeyenda, "OPEN", strSQL)	
	strLeyenda = ""
	if (not rsLeyenda.EoF) then 
	    strLeyenda = rsLeyenda("LEYENDA")
	    'Reemplazo los parametros
	    if (isArray(pParams)) then
	        For i= LBound(pParams) to UBound(pParams)
    	        strLeyenda = Replace(strLeyenda, "<" & i & ">", pParams(i), 1, 1)
	        Next 
	    end if
	end if
	getDocLeyendaFactura = strLeyenda
End Function
'--------------------------------------------------------------------------------------------------------------------
'MOD: 2017-10-04 - JAS
'Función que dibuja un detalle tal cual esta cargado, se coloca en commons ya que lo puede usar cualquier tipo de cbte si su detalle no se ajusta al formato esperado.
Function dibujarDetalleContenidoSTD(bodyText, ByRef idx, paginaLineas, maxLineas, ByRef pOpdf)
	Dim i,strDetalle,indexIni
	
	Call GF_squareBox(pOpdf, 3, 265, 587, 430, 0,"#FFFFFF",NEGRO,1,PDF_SQUARE_ROUND)
	Call GF_setFont(pOpdf,"COURIER", 9 , FONT_STYLE_NORMAL)
	Call GF_verticalLine(pOPDF, 497, 265, 435)
	if (pIsProforma) then
	    Call GF_writeImage(pOPDF, Server.MapPath("..\Images\facturas\MarcaAguaProforma.gif"),100, 325, 374, 373, 0)
	end if	
	indexIni = 268		
	i = 0
	While ((i < paginaLineas) and (idx <= maxLineas))
		Call GF_writeTextAlign(pOpdf, 10, indexIni, bodyText(idx, 0) , 440, PDF_ALIGN_LEFT)		
		if (CDbl(bodyText(idx, 1)) > 0) then
			Call GF_writeTextAlign(pOpdf, 500, indexIni, FormatNumber(bodyText(idx, 1), 2, 0, 0, -1) , 84, PDF_ALIGN_RIGHT)		
		end if
		indexIni = indexIni + 12
		idx = idx + 1
		i = i + 1		
	wend	
End Function
'----------------------------------------------------------------------------------------------------------------------
'MOD: 2017-10-04 - JAS
Function generarPDF(pNroReg,pMode)
	Dim strSQLComm,fileName,cantReg,rsCommon,oConn
	'Obtengo los conceptos de todos los numero de registros pasados, para poder saber que formulario debo utilizar
	strSQLComm = "Select * from FAC001A where guid in ('"& pNroReg &"')"	
	Call executeQueryDb(DBSITE_SQL_MAGIC, rsCommon, "OPEN", strSQLComm)
	if (not rsCommon.Eof)	then 
		fileName = "FACTURA_" & rsCommon("tipocpte") & "_" & pNroReg & ".pdf"		
		cantReg = rsCommon.RecordCount
		Randomize
		if (cantReg > 1) then  fileName = "FACTURAS_VARIAS_" & session("MmtoDato") & Int(100*Rnd) & ".pdf"
		rtrn = Server.MapPath("..\temp\" & fileName)
		Set oPDF = GF_createPDF(rtrn)		
		Call GF_setPDFMode(pMode)
		while (not rsCommon.Eof)
			'Dependiendo del numero de Formulario que tenga la factura imprimirá Exportacion o Local(Default)			
		    'Select Case (CLng(rsCommon("tipcbt")))
			    'Case FAC_FORMULARIO_IMPRESION_EX
				'    Call crearPDF_Ex(rsCommon("guid"),oPDF)
                'Case else
					Call crearPDF_Local(rsCommon("guid"),oPDF)	
		    'End Select
            rsCommon.MoveNext()
			if not rsCommon.Eof then call GF_newPage(oPDF)			
		wend
		Call GF_closePDF(oPDF)		
	end if	
	generarPDF = fileName
End Function
'************************************************************************************************************************
'************************************************************************************************************************
'********************************              COMIENZO DE LA PAGINA                   **********************************
'************************************************************************************************************************
'************************************************************************************************************************
Dim nroReg,oPDF,g_CodConcepto,rtrnNameFile,g_Observaciones

nroReg = GF_Parametros7("lote","",6)
tipoPDF = GF_Parametros7("tipoPDF",0,6)


if (nroReg <> "") then
	rtrnNameFile = generarPDF(nroReg,tipoPDF)
	if (tipoPDF = PDF_FILE_MODE) then Response.Write rtrnNameFile
end if

%>
