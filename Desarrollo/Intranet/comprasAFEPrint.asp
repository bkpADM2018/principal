<!-- #include file="Includes/procedimientosMG.asp"-->
<!-- #include file="Includes/procedimientosPDF.asp"-->
<!-- #include file="Includes/procedimientosAFE.asp"-->
<!-- #include file="Includes/procedimientosCompras.asp"-->
<!-- #include file="Includes/procedimientosPCT.asp"-->
<!-- #include file="Includes/procedimientosObras.asp"-->
<!-- #include file="Includes/procedimientosFechas.asp"-->
<!-- #include file="Includes/procedimientosTraducir.asp"-->
<!-- #include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!-- #include file="Includes/procedimientosMath.asp"-->
<!-- #include file="Includes/procedimientosUser.asp"-->
<%
Const MAX_RENGLONES_PAGINA1 = 22
Const MAX_RENGLONES = 94
'----------------------------------------------------------------------------------------------------------
Function DibujarTitular()
	'logo
	Call GF_writeImage(Gbl_oPDF, Server.MapPath("Images\LogoACTI.gif"),25, 15, 125, 30, 0)
	pdf_currentFontColor = "#0B3B0B"
	'Titulo
	Call GF_setFont(Gbl_oPDF,"ARIAL", 16,8)
	Call GF_writeTextAlign(Gbl_oPDF,25, 25, "Authorization for Expenditure (AFE)", 590,PDF_ALIGN_CENTER)
	pdf_currentFontColor = "000000"
End Function 
'----------------------------------------------------------------------------------------------------------
Function DibujarContenedor1()
	Dim PxlFila,Col(4),seleccion, cdAFE, myCompania,myAfe, myType, myCategory,myDsDivision
	  
	Call GF_horizontalLine(Gbl_oPDF,25  ,48  , 545)	 
	Call GF_horizontalLine(Gbl_oPDF,25  ,168 , 545)
	Call GF_verticalLine  (Gbl_oPDF,25  ,48  , 74 )
	Call GF_verticalLine  (Gbl_oPDF,77  ,48  , 74 )
	Call GF_verticalLine  (Gbl_oPDF,570 ,48  , 74 )
	Call GF_verticalLine  (Gbl_oPDF,25  ,128 , 14 )
	Call GF_verticalLine  (Gbl_oPDF,77  ,128 , 14 )
	Call GF_verticalLine  (Gbl_oPDF,570 ,128 , 14 )
	Call GF_verticalLine  (Gbl_oPDF,25  ,154 , 14 )
	Call GF_verticalLine  (Gbl_oPDF,77  ,154 , 14 )
	Call GF_verticalLine  (Gbl_oPDF,570 ,154 , 14 )
	 
	col(0)=100
	col(1)=220
	col(2)=340
	col(3)=480
	seleccion = 0	
	'							DIBUJO LAS COLUMNAS DE LA CABECERA	
	Call GF_setFont(Gbl_oPDF,"ARIAL", 8,8)
	Call GF_writeTextAlign(Gbl_oPDF,28, 50, "Company"	 , 45,PDF_ALIGN_LEFT)
	Call GF_horizontalLine(Gbl_oPDF,25  ,62,545)
	Call GF_writeTextAlign(Gbl_oPDF,28, 65, "Division"	 , 45,PDF_ALIGN_LEFT)
	Call GF_horizontalLine(Gbl_oPDF,25  ,77,545)
	Call GF_writeTextAlign(Gbl_oPDF,28, 80, "Location"	 , 45,PDF_ALIGN_LEFT)
	Call GF_horizontalLine(Gbl_oPDF,25  ,92,545)
	Call GF_writeTextAlign(Gbl_oPDF,28, 95, "AFE No."	 , 45,PDF_ALIGN_LEFT)
	Call GF_horizontalLine(Gbl_oPDF,25 ,107,545)
	Call GF_writeTextAlign(Gbl_oPDF,28, 110, "AFE Title" , 45,PDF_ALIGN_LEFT)
	Call GF_horizontalLine(Gbl_oPDF,25  ,122,545)	
	Call GF_horizontalLine(Gbl_oPDF,25  ,128,545)
	Call GF_writeTextAlign(Gbl_oPDF,28, 131, "Category"	 , 45,PDF_ALIGN_LEFT)	
	Call GF_horizontalLine(Gbl_oPDF,25  ,142,545)
	Call GF_horizontalLine(Gbl_oPDF,25  ,154,545)
	Call GF_writeTextAlign(Gbl_oPDF,28, 156, "Type"		 , 45,PDF_ALIGN_LEFT)	
	
	'							DIBUJO LOS DATOS DE LA CABECERA	
	myCompania	= getDescripcionProveedor(CD_TOEPFER)
	if afe_Titulo = "" then
		myAfe= "-"
	else
		myAfe= ucase(afe_Titulo)
	end if	
	if afe_IdDivision = "0" then
		myDsDivision = "-"		
	else
		myDsDivision= getDescripcionDivision(afe_IdDivision)
		if myDsDivision = "" then myDsDivision = "-"		
	end if			
	
	'	BUSCO EL TIPO CORRESPONDIENTE	
	' Tipos puede haber más de uno, divido y busco todos los que sean
	arrTipo = Split(afe_Tipo, ",")
	for k=LBound(arrTipo) to UBound(arrTipo)
		auxTipo = Trim(arrTipo(k))
		'Si se elige un tipo de cumplimiento, se muestra el tipo directamente.		
		if (auxTipo = AFE_TIPO_CUMPIMIENTO) then auxTipo = afe_TipoCC
		myType = myType & getDescripcionTipoAFE(auxTipo) 
		'Si eligió otros, entonces se muestra la descripción.	
		if (auxTipo = AFE_TIPO_OTROS)		then myType = myType & afe_TipoOtros
		myType = myType & ", "		
	Next		
	myType = left(myType, Len(myType)-2)
	
	'	BUSCO LA CATEGORIA CORRESPONDIENTE
	myCategory = getDescripcionCategoriaAFE(afe_Categoria)
	if (afe_Categoria = AFE_CATEGORIA_OTROS) then myCategory = myCategory & afe_CatOtros
	if (afe_NroAFEComplID <> 0)				then myCategory = myCategory & getCdAFE(afe_NroAFEComplID)
	
	Call GF_setFont(Gbl_oPDF,"ARIAL", 8,0)
	Call GF_writeTextAlign(Gbl_oPDF,82, 50 , myCompania	  	, 480,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,82, 65 , Ucase(myDsDivision)	, 480,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,82, 80 , "-"					, 480,PDF_ALIGN_LEFT)
	Call GF_setFont(Gbl_oPDF,"ARIAL", 8,8)
	Call GF_writeTextAlign(Gbl_oPDF,82, 95 , afe_CdAFE		, 480,PDF_ALIGN_LEFT)		
	Call GF_writeTextAlign(Gbl_oPDF,82, 110, myAfe			, 480,PDF_ALIGN_LEFT)		
	Call GF_setFont(Gbl_oPDF,"ARIAL", 8,0)
	Call GF_writeTextAlign(Gbl_oPDF,82, 131, myCategory	, 480,PDF_ALIGN_LEFT)	
	Call GF_setFont(Gbl_oPDF,"ARIAL", 6, 0)
	Call GF_writeTextAlign(Gbl_oPDF,82, 156, myType		, 480,PDF_ALIGN_LEFT)	
	Call GF_writeTextAlign(Gbl_oPDF,82, 144, "(Enter:'Capital','Expense','Supplement to AFE No.','Lease' or 'Other' + description)" , 590,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,80, 169, "(Enter:'Improved Efficiency','Increased Capacity','Maintenance','Vehicle','IT/Telecommunication','Change of Scope'," , 590,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,80, 178, "'Spare Parts','Overspend','Compliance - Environmental / Health & Safety / Quality Assurance')" , 590,PDF_ALIGN_LEFT)			
	Call GF_setFont(Gbl_oPDF,"ARIAL", 8, 0)
	
end Function 
'----------------------------------------------------------------------------------------------------------
Function DibujarContenedor2()
	Dim vecRenglones, idxRenglones, idxPagina, idxRenglonesPagina, yInicial
	
	'Se dibuja el cuadro de descripción
	Call GF_horizontalLine(Gbl_oPDF,25  ,195  , 545)
	Call GF_horizontalLine(Gbl_oPDF,25  ,420 , 545)
	Call GF_verticalLine  (Gbl_oPDF,25  ,195  , 225 )
	Call GF_verticalLine  (Gbl_oPDF,570 ,195  , 225 )
	'Se imprime el dato de la partida presupuestaria
	if (afe_IdObra > 0) then	
		Call GF_setFont(Gbl_oPDF,"ARIAL", 6,0)	
		Call GF_writeTextAlign(Gbl_oPDF,25 , 423 , "<i>Job: " & afe_ObraCD & "-" & afe_ObraDS & " (" & afe_IDArea & "-" & afe_IDDetalle & ")</i>", 570,PDF_ALIGN_LEFT)
	end if
	Call GF_setFont(Gbl_oPDF,"ARIAL", 8,8)	
	Call GF_writeTextAlign(Gbl_oPDF,28 , 200, "DESCRIPTION", 560,PDF_ALIGN_LEFT)
	Call GF_horizontalLine(Gbl_oPDF,28  ,209 , 57)
	Call GF_setFont(Gbl_oPDF,"ARIAL", 8,0)
	vecRenglones = splitRecordLines(Gbl_oPDF, afe_Descripcion, 532)	
	'Se arman las paginas	
	idxPagina = 0
	idxRenglones = 0
	yInicial = 215 
	while (idxRenglones <= UBound(vecRenglones))
		'Determino el maximo nro de renglones para la pagina.
		maxRenglones = MAX_RENGLONES
		if (idxPagina = 0) then maxRenglones = MAX_RENGLONES_PAGINA1
		'Armo el texto de la pagina.
		idxRenglonesPagina=0		
		while ((idxRenglones <= UBound(vecRenglones)) and (idxRenglonesPagina <= maxRenglones))
			'response.write vecRenglones(idxRenglones) & "<br>"
			Call GF_writeTextAlign(Gbl_oPDF, 28, yInicial, vecRenglones(idxRenglones), 532, PDF_ALIGN_LEFT)
			idxRenglones = idxRenglones + 1
			idxRenglonesPagina = idxRenglonesPagina + 1
			yInicial = yInicial + 8
		wend				
		'response.end
		idxPagina = idxPagina + 1
		yInicial = NuevaHoja()
	wend	
end Function
'----------------------------------------------------------------------------------------------------------
Function DibujarContenedor3()
	Dim total, local, code, rate, payback, irr, ROIC, NPV
	
	Call setWorkPage(Gbl_oPDF, 1)
	
	Call GF_squareBox(Gbl_oPDF,25 ,435,136,30,0,"#FFFFFF","#000000",1,0)
	Call GF_squareBox(Gbl_oPDF,161,435,136,30,0,"#FFFFFF","#000000",1,0) 
	Call GF_squareBox(Gbl_oPDF,297,435,136,30,0,"#FFFFFF","#000000",1,0) 
	Call GF_squareBox(Gbl_oPDF,433,435,136,30,0,"#FFFFFF","#000000",1,0) 
	
	Call GF_squareBox(Gbl_oPDF,25 ,465,136,30,0,"#FFFFFF","#000000",1,0)
	Call GF_squareBox(Gbl_oPDF,161,465,136,30,0,"#FFFFFF","#000000",1,0)
	Call GF_squareBox(Gbl_oPDF,297,465,136,30,0,"#FFFFFF","#000000",1,0)
	Call GF_squareBox(Gbl_oPDF,433,465,136,30,0,"#FFFFFF","#000000",1,0)
	
	Call GF_setFont(Gbl_oPDF,"ARIAL", 8,8)
	Call GF_writeTextAlign(Gbl_oPDF,28 , 445, "Total Expenditure (USD)"	 , 570,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,300, 445, "Local Currency Amount"	 , 570,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,28, 475 , "Exchange Rate"		 	 , 570,PDF_ALIGN_LEFT)	
	Call GF_writeTextAlign(Gbl_oPDF,300, 475, "Currency Code"		 	 , 570,PDF_ALIGN_LEFT)
	
	Call GF_squareBox(Gbl_oPDF,25 ,500,68,15,0 ,"#FFFFFF","#000000",1,0)
	Call GF_squareBox(Gbl_oPDF,93 ,500,68,15,0  ,"#FFFFFF","#000000",1,0)
	Call GF_squareBox(Gbl_oPDF,161,500,68,15,0 ,"#FFFFFF","#000000",1,0)
	Call GF_squareBox(Gbl_oPDF,229,500,68,15,0 ,"#FFFFFF","#000000",1,0)
	Call GF_squareBox(Gbl_oPDF,25 ,515,68,15,0 ,"#FFFFFF","#000000",1,0)
	Call GF_squareBox(Gbl_oPDF,93 ,515,68,15,0  ,"#FFFFFF","#000000",1,0)
	Call GF_squareBox(Gbl_oPDF,161,515,68,15,0 ,"#FFFFFF","#000000",1,0)
	Call GF_squareBox(Gbl_oPDF,229,515,68,15,0 ,"#FFFFFF","#000000",1,0)
	
	Call GF_writeTextAlign(Gbl_oPDF,28 , 503 , "NPV"		 , 570,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,28, 518	 , "ROIC"	 	 , 570,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,164 , 503, "IRR"		 , 570,PDF_ALIGN_LEFT)	
	Call GF_writeTextAlign(Gbl_oPDF,164, 518 , "Payback"	 , 570,PDF_ALIGN_LEFT)
	
	total   = GF_EDIT_DECIMALS(cdbl(afe_ImporteDolares),2)
	local   = GF_EDIT_DECIMALS(cdbl(afe_ImportePesos)  ,2)
	code    = getSimboloMonedaLetras(MONEDA_PESO)
	rate    = afe_TipoCambio
	if afe_PAYBACK = "" then
		payback = "NA"
	else
		payback = GF_EDIT_DECIMALS(cdbl(afe_PAYBACK),2)
	end if
	if afe_NPV = "0" then
		NPV = "NA"
	else
		NPV = GF_EDIT_DECIMALS(cdbl(afe_NPV),2)
	end if
	if afe_ROIC = "0" then
		ROIC = "NA"
	else
		ROIC = GF_EDIT_DECIMALS(cdbl(afe_ROIC),2)
	end if
	if afe_Irr = "0" then
		irr     = "NA"
	else
		irr     = GF_EDIT_DECIMALS(cdbl(afe_Irr),2)
	end if
	Call GF_setFont(Gbl_oPDF,"ARIAL", 8,0)
	Call GF_writeTextAlign(Gbl_oPDF,170 , 445, total , 118,PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(Gbl_oPDF,170 , 475, rate	 , 118,PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(Gbl_oPDF,442 , 445, local , 118,PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(Gbl_oPDF,442 , 475, code ,  118,PDF_ALIGN_LEFT)
	
	Call GF_writeTextAlign(Gbl_oPDF,235 , 503, irr ,  58,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,98 , 503, NPV   ,  58,PDF_ALIGN_LEFT)	
	Call GF_writeTextAlign(Gbl_oPDF,235 , 518, payback ,  58,PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(Gbl_oPDF,98 , 518, ROIC ,  58,PDF_ALIGN_LEFT)	
end Function
'-----------------------------------------------------------------------------------------------------------
Function drawBoxSignature(p_CdUsuario,p_FechaFirma,p_HKey,p_X,p_Y) 
    Call GF_setFont(Gbl_oPDF,"ARIAL", 6,0) 
    Call GF_writeTextAlign(Gbl_oPDF,p_X - 7, p_Y + 33, getUserDescription(p_CdUsuario), 90,PDF_ALIGN_CENTER)
    if (Trim(p_HKey) <> "") then
        Call GF_writeTextAlign(Gbl_oPDF,p_X - 50, p_Y + 25, left(GF_FN2DTE(left(p_FechaFirma,8)),5), 45,PDF_ALIGN_CENTER)
        Call GF_setFont(Gbl_oPDF,"ARIAL", 3,0)
        Call GF_writeImage(Gbl_oPDF, server.MapPath(".") & "\images\firmas\" & obtenerFirma(p_CdUsuario), p_X, p_Y, 80, 30, 0)
	    Call GF_writeTextAlign(Gbl_oPDF,p_X + 5, p_Y + 41,armarTextoPlanoFirma(p_HKey, p_FechaFirma), 41,PDF_ALIGN_CENTER)        
    end if
End Function
'----------------------------------------------------------------------------------------------------------
Function DibujarContenedor4()
	Call GF_squareBox(Gbl_oPDF,25 ,535 ,272,15,0,"#FFFFFF","#000000",1,0)
	Call GF_squareBox(Gbl_oPDF,297 ,535,272,15,0,"#FFFFFF","#000000",1,0)
	Call GF_setFont(Gbl_oPDF,"ARIAL", 8,8)
	Call GF_writeTextAlign(Gbl_oPDF,28 , 538 , "Review"	 , 570,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,164, 538 , "Date"	 , 570,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,199, 538 , "Signature"	 , 570,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,303, 538 , "Approval"	 , 570,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,436, 538 , "Date"	 , 570,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,471, 538 , "Signature"	 , 570,PDF_ALIGN_LEFT)
	Call GF_setFont(Gbl_oPDF,"ARIAL", 8,0)
	Call GF_squareBox(Gbl_oPDF,25 ,550 ,272,250,0,"#FFFFFF","#000000",1,0)
	Call GF_squareBox(Gbl_oPDF,297,550 ,272,250,0,"#FFFFFF","#000000",1,0)
	Call GF_horizontalLine(Gbl_oPDF,25   ,600 , 544)
	Call GF_horizontalLine(Gbl_oPDF,25   ,650 , 544)
	Call GF_horizontalLine(Gbl_oPDF,25   ,700 , 544)
	Call GF_horizontalLine(Gbl_oPDF,25   ,750 , 544)
	Call GF_verticalLine  (Gbl_oPDF,161  ,550 , 250)
	Call GF_verticalLine  (Gbl_oPDF,195  ,550 , 250)
	Call GF_verticalLine  (Gbl_oPDF,433  ,550 , 250)
	Call GF_verticalLine  (Gbl_oPDF,467  ,550 , 250)
	
	Call GF_writeTextAlign(Gbl_oPDF,25   , 570 , "Estimate prepared by"	 	  , 140,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(Gbl_oPDF,25   , 620 , "Expenditure requested by"	  , 140,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(Gbl_oPDF,25   , 670 , "Engineering review by"	  , 140,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(Gbl_oPDF,25   , 720 , "Local controller review by" , 140,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(Gbl_oPDF,25   , 770 , "CFO / Treasurer review by"  , 140,PDF_ALIGN_CENTER)
	
	Call GF_writeTextAlign(Gbl_oPDF,300  , 570 , "Local Management"	 , 140,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(Gbl_oPDF,300  , 620 , "Local Management"	 , 140,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(Gbl_oPDF,300  , 670 , "Toepfer Director" , 140,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(Gbl_oPDF,300  , 720 , "Country Director"	 , 140,PDF_ALIGN_CENTER)	
	
    
    'PREPARA AFE
    Call drawBoxSignature(afe_PreparedByCD,afe_PreparedByHkeyDate,afe_PreparedByHkey,205,555)
    'REQUIERE AFE
    Call drawBoxSignature(afe_RequestedByCD,afe_RequestedByHkeyDate,afe_RequestedByHkey,205,605)
    'REVISION TECNICA
    Call drawBoxSignature(afe_EngReviewCD,afe_EngReviewHkeyDate,afe_EngReviewHkey,205,655)
    'GERENTE DE PUERTOS
    Call drawBoxSignature(afe_OfficerCD,afe_OfficerHkeyDate,afe_OfficerHkey,477,555)    
    'COORDINADOR DE PUERTOS
    Call drawBoxSignature(afe_VicePresidentCD,afe_VicePresidentHkeyDate,afe_VicePresidentHkey,477,605)
    'CONTROLLER
    Call drawBoxSignature(afe_ControllerCD,afe_ControllerHkeyDate,afe_ControllerHkey,205,705)
    'FINANZAS 
    Call drawBoxSignature(afe_cfoCD,afe_cfoHkeyDate,afe_cfoHkey,205,755)
    'Toepfer Director
    Call drawBoxSignature(afe_PresidentCD,afe_PresidentHkeyDate,afe_PresidentHkey,477,655)
	'Country Director
	Call GF_setFont(Gbl_oPDF,"ARIAL", 6,0) 
	afe_BDT = "Alejandro Ingham"
	Call GF_writeTextAlign(Gbl_oPDF,470, 738, afe_BDT, 90,PDF_ALIGN_CENTER)	
    Call GF_setFont(Gbl_oPDF,"ARIAL", 3,0) 
end Function
'-----------------------------------------------------------------------------------------------------------
Function DibujarRadio(texto,x,y,chequeado)
    Call GF_writeImage(Gbl_oPDF, Server.MapPath("Images\Radio_Chk"&chequeado&".gif"),x,y,8,8, 0)
	
	if texto <> "" then
		Call GF_writeTextAlign(Gbl_oPDF,x+10, y, texto, 590,PDF_ALIGN_LEFT)
	end if
end Function
'-----------------------------------------------------------------------------------------------------------
Function DibujarRegla(inicio,separacion,y)
	Dim resu,i
	for i = 0 to 588 step separacion
		Call GF_verticalLine  (Gbl_oPDF,inicio+i,y,10)
	next
end Function
'----------------------------------------------------------------------------------------------------------
Function NuevaHoja()
	nroHojas = nroHojas +1
	Call GF_newPage(Gbl_oPDF)	
	Call GF_setFont(Gbl_oPDF,"ARIAL", 8,8)
	Call GF_writeTextAlign(Gbl_oPDF,5 ,840, "Page " & nroHojas, 580,PDF_ALIGN_RIGHT)
	Call DibujarTitular()
	Call DibujarMarcaDeAgua() 	 
	Call GF_setFont(Gbl_oPDF,"ARIAL", 16,8)
	Call GF_squareBox(Gbl_oPDF,5,50,580 ,790,0,"#FFFFFF","#000000",1,PDF_SQUARE_ROUND) 	
	Call GF_writeTextAlign(Gbl_oPDF,5 ,50, "DETAIL", 590,PDF_ALIGN_CENTER)		
	NuevaHoja=70 'Devuelve la linea inicial de las pagnas complementarias del AFE.
end Function
'--------------------------------------------------------------------------------------
Function DibujarMarcaDeAgua()
 if (afe_Confirmado = "R") then
	'el afe esta rechazado	
	Call GF_writeImage(Gbl_oPDF, Server.MapPath("Images\compras\AFE_canceled_wathermark.png"),90, 220, 500, 100, 335)
 end if
end function
'--------------------------------------------------------------------------------------
function esAMano(pValue)
	if pValue = A_MANO then esAMano = true
end function
'**********************************************************************************************
'**********					INICIO DE LA PAGINA					 **********
'**********************************************************************************************
 Dim Gbl_oPDF,ds,idAfe,nroHojas

'SETEO EL IDIOMA EN INGLES
GF_SET_IDIOMA(2)
	
 nroHojas = 1

 idAfe = GF_Parametros7("IDAFE","",6)
 Call readAFE(IdAfe, 0, 0)

 hojas = 1
 filename = "test.pdf"

 Set Gbl_oPDF = GF_createPDF(Server.MapPath("temp\" & filename))
 Call GF_setPDFMode(PDF_STREAM_MODE)
 
if (CLng(Left(afe_Momento, 8)) >= 20121009) then
	Call DibujarTitular()
	Call DibujarContenedor1()		' CABECERA
	Call DibujarContenedor2() 		' DESCRIPCION
	Call DibujarContenedor3() 		' TOTALES
	Call DibujarContenedor4() 		' FIRMAS
	Call DibujarMarcaDeAgua() 
 else
	'Call DibujarEncabezadoFV ()
	'Call DibujarContenedor1FV() 'category/type
	'Call DibujarContenedor2FV() 'company,no, general account, etc
	'Call DibujarContenedor3FV() ' firmas y totales
	'Call DibujarContenedor4FV() 'additional info
	'Call DibujarMarcaDeAguaFV() 
	'Call DibujarDescripcionFV()
 end if

 Call GF_closePDF(Gbl_oPDF)

'VUELVO A SETEAR EL IDIOMA EN ESPAÑOL
GF_SET_IDIOMA(1)

'**********************************************************************************************
'**********					FIN DE LA PAGINA					 **********
'**********************************************************************************************

%>


