<%
Const MAX_RENGLONESFV_PAGINA1FV = 16
Const MAX_RENGLONESFV = 94
Const MAX_LOCAL_LIMITFV = 5000000 'Maximo importe en centavos de dolar que permite no enviar el AFE a Hambuergo.
'----------------------------------------------------------------------------------------------------------
Function getLocation(pIdDivision) 
	Dim rs, conn, strSQL, rtrn

	strSQL="Select M.DESCPC, M.AUXIPC from MERFL.MER142F1 M inner join TOEPFERDB.TBLDIVISIONES T ON M.CODIPC =T.CODIGOPOSTAL where T.IDDIVISION=" & pIdDivision
	'Response.Write strSQL
	Call GF_BD_COMPRAS(rs, conn, "OPEN", strSQL)
	if rs.recordCount > 0 then 
		'EAB VERRRRRRR
		while not rs.eof 
			if ucase(left(rs("DESCPC"),3)) = "ARR" or ucase(left(rs("DESCPC"),10)) = "PUERTO SAN" then
				rtrn = rs("DESCPC")
			end if	
			rs.movenext
		wend
	else
		if not rs.eof then
			rtrn = rs("DESCPC")
		end if
	end if
	Call GF_BD_COMPRAS(rs, conn, "CLOSE", strSQL)
	getLocation = rtrn
End Function
'----------------------------------------------------------------------------------------------------------
Function DibujarEncabezadoFV()

	'logo
	Call GF_writeImage(Gbl_oPDF, Server.MapPath("Images\kogge64.gif"),5, 2, 40, 40, 0)
	
	'Titulo
	Call GF_setFont(Gbl_oPDF,"ARIAL", 16,8)
	Call GF_writeTextAlign(Gbl_oPDF,0, 10, "AUTHORIZATION FOR EXPENDITURE", 590,PDF_ALIGN_CENTER)
	
	
	Call GF_setFont(Gbl_oPDF,"ARIAL", 6,8)
	Call GF_writeTextAlign(Gbl_oPDF,0, 20, "AFE NO.: " & afe_CdAFE , 590,PDF_ALIGN_RIGHT)
	

	Call GF_setFont(Gbl_oPDF,"ARIAL", 8,8)
	Call GF_writeTextAlign(Gbl_oPDF,0, 30, "ALFRED C.TOEPFER S.R.L", 590,PDF_ALIGN_CENTER)
	
	Call GF_setFont(Gbl_oPDF,"ARIAL", 6,0)
	Call GF_writeTextAlign(Gbl_oPDF,0, 30, "EST. COMPLETION DATE: " & mid(afe_ObraFechaFin,5,2) & mid(afe_ObraFechaFin,3,2), 590,PDF_ALIGN_RIGHT)
	
	Call GF_setFont(Gbl_oPDF,"ARIAL", 6,0)
	Call GF_writeTextAlign(Gbl_oPDF,0, 40, "(MMYY)", 590,PDF_ALIGN_RIGHT)
	
End Function 
'----------------------------------------------------------------------------------------------------------
Function DibujarContenedor1FV()

	 '-----------------------------------------------------------------------------------
	 'Dibujo primer contenedor
	 Call GF_horizontalLine(Gbl_oPDF,5  ,50 ,585)
	 Call GF_horizontalLine(Gbl_oPDF,5  ,80 ,585)
	 Call GF_horizontalLine(Gbl_oPDF,5  ,130,585)
	 
	 Call GF_verticalLine  (Gbl_oPDF,5  ,50 ,80 )
	 Call GF_verticalLine  (Gbl_oPDF,590,50 ,80 )
	 '-----------------------------------------------------------------------------------
	 
	Dim Col(4),seleccion, cdAFE
	Dim PxlFila
	col(0)=100
	col(1)=220
	col(2)=340
	col(3)=480
	
	seleccion = 0
	Call GF_setFont(Gbl_oPDF,"ARIAL", 8,8)
	Call GF_writeTextAlign(Gbl_oPDF,10, 53, "CATEGORY", 590,PDF_ALIGN_LEFT)
	
	Call GF_setFont(Gbl_oPDF,"ARIAL", 8,0)
	
	'primera fila de radios
		if afe_Categoria = "C" then seleccion = 1 else seleccion = 0 end if
		Call DibujarRadioFV("CAPITAL",col(0)-10,53,seleccion)

		if afe_Categoria = "G" then seleccion = 1 else seleccion = 0 end if
		Call DibujarRadioFV("EXPENSE",col(1)-10,53,seleccion)
		
		if afe_Categoria = "I" then seleccion = 1 else seleccion = 0 end if
		Call DibujarRadioFV("INVESTMENT",col(2)-10,53,seleccion)
				
		if afe_Categoria = "A" then seleccion = 1 else seleccion = 0 end if
		Call DibujarRadioFV("SUPPLEMENT TO AFE NO.:",col(3)-10,53,seleccion)
		cdAFE = getCdAFE(afe_NroAFEComplID)
		if  (cdAFE = "") then
			Call GF_writeTextAlign(Gbl_oPDF,col(3), 63, "NA", 590,PDF_ALIGN_LEFT)
		else
			Call GF_writeTextAlign(Gbl_oPDF,col(3), 63, cdAFE, 590,PDF_ALIGN_LEFT)
		end if
		
	
	'segunda fila de radios
	
		if afe_Categoria = "Q" then seleccion = 1 else seleccion = 0 end if
		Call DibujarRadioFV("LEASE",col(0)-10,63,seleccion)
				
		if afe_Categoria = "S" then seleccion = 1 else seleccion = 0 end if
		Call DibujarRadioFV("CONSULTING SERVICES",col(1)-10,63,seleccion)
				
		if afe_Categoria = "O" then seleccion = 1 else seleccion = 0 end if
		Call DibujarRadioFV("OTHER:",col(2)-10,63,seleccion)
				
		if afe_Categoria = "O" then
			Call GF_writeTextAlign(Gbl_oPDF,col(2)+40, 63,afe_CatOtros, 590,PDF_ALIGN_LEFT)
		end if
	
	'***************
	Call GF_setFont(Gbl_oPDF,"ARIAL", 8,8)
	Call GF_writeTextAlign(Gbl_oPDF,10, 83, "TYPE", 590,PDF_ALIGN_LEFT)
	
	Call GF_setFont(Gbl_oPDF,"ARIAL", 8,0)
	
	'Primero fila Call DibujarCheckFVs
		PxlFila = 83
		if inSTR(afe_Tipo,"E") <> 0 then seleccion = 1 else seleccion = 0 end if
		Call DibujarCheckFV(col(0)-10,PxlFila,seleccion)
		Call GF_writeTextAlign(Gbl_oPDF,col(0), PxlFila, "IMPROVED EFFICIENCY", 590,PDF_ALIGN_LEFT)
		
		if inSTR(afe_Tipo,"R") <> 0 then seleccion = 1 else seleccion = 0 end if
		Call DibujarCheckFV(col(1)-10,PxlFila,seleccion)
		Call GF_writeTextAlign(Gbl_oPDF,col(1), PxlFila, "SPARE PARTS", 590,PDF_ALIGN_LEFT)
		
		if inSTR(afe_Tipo,"C") <> 0 then seleccion = 1 else seleccion = 0 end if
		Call DibujarCheckFV(col(2)-10,PxlFila,seleccion)
		Call GF_writeTextAlign(Gbl_oPDF,col(2), PxlFila, "COMPLIANCE (Check ONE>)", 590,PDF_ALIGN_LEFT)
		
		if inSTR(afe_TipoCC,"A") <> 0 then seleccion = 1 else seleccion = 0 end if
		Call DibujarRadioFV("",col(3)-10,PxlFila,seleccion)
		Call GF_writeTextAlign(Gbl_oPDF,col(3), PxlFila, "ENVIRONMENTAL", 590,PDF_ALIGN_LEFT)
	
	'Segunda fila Call DibujarCheckFVs
		PxlFila = 93
		if inSTR(afe_Tipo,"I") <> 0 then seleccion = 1 else seleccion = 0 end if
		Call DibujarCheckFV(col(0)-10,PxlFila,seleccion)
		Call GF_writeTextAlign(Gbl_oPDF,col(0), PxlFila, "INCREASED CAPACITY", 590,PDF_ALIGN_LEFT)
		
		if inSTR(afe_Tipo,"M") <> 0 then seleccion = 1 else seleccion = 0 end if
		Call DibujarCheckFV(col(1)-10,PxlFila,seleccion)
		Call GF_writeTextAlign(Gbl_oPDF,col(1), PxlFila, "MAINTENANCE", 590,PDF_ALIGN_LEFT)
		
		if inSTR(afe_Tipo,"V") <> 0 then seleccion = 1 else seleccion = 0 end if
		Call DibujarCheckFV(col(2)-10,PxlFila,seleccion)
		Call GF_writeTextAlign(Gbl_oPDF,col(2), PxlFila, "VEHICLE", 590,PDF_ALIGN_LEFT)
		
		if inSTR(afe_TipoCC,"S") <> 0 then seleccion = 1 else seleccion = 0 end if
		Call DibujarRadioFV("",col(3)-10,PxlFila,seleccion)
		Call GF_writeTextAlign(Gbl_oPDF,col(3), PxlFila, "HEALTH & SAFETY", 590,PDF_ALIGN_LEFT)
	
	'Tercera fila Call DibujarCheckFVs
		PxlFila = 103
		if inSTR(afe_Tipo,"D") <> 0 then seleccion = 1 else seleccion = 0 end if
		Call DibujarCheckFV(col(0)-10,PxlFila,seleccion)
		Call GF_writeTextAlign(Gbl_oPDF,col(0), PxlFila, "CHANGE OF SCOPE", 590,PDF_ALIGN_LEFT)
		
		if inSTR(afe_Tipo,"Y") <> 0 then seleccion = 1 else seleccion = 0 end if
		Call DibujarCheckFV(col(1)-10,PxlFila,seleccion)
		Call GF_writeTextAlign(Gbl_oPDF,col(1), PxlFila, "OVERSPEND", 590,PDF_ALIGN_LEFT)
		
		if inSTR(afe_TipoCC,"N") <> 0 then seleccion = 1 else seleccion = 0 end if
		Call DibujarRadioFV("",col(3)-10,PxlFila,seleccion)
		Call GF_writeTextAlign(Gbl_oPDF,col(3), PxlFila, "QUALITY ASSURANCE", 590,PDF_ALIGN_LEFT)
	
	'Cuarta fila Call DibujarCheckFVs
		if inSTR(afe_Tipo,"T") <> 0 then seleccion = 1 else seleccion = 0 end if
		PxlFila = 113
		Call DibujarCheckFV(col(0)-10,PxlFila,seleccion)
		Call GF_writeTextAlign(Gbl_oPDF,col(0), PxlFila, "IT/TELECOMMUNICATIONS", 590,PDF_ALIGN_LEFT)
		
		if inSTR(afe_Tipo,"O") <> 0 then seleccion = 1 else seleccion = 0 end if
		Call DibujarCheckFV(col(1)-10,PxlFila,seleccion)
		Call GF_writeTextAlign(Gbl_oPDF,col(1), PxlFila, "OTHER:", 590,PDF_ALIGN_LEFT)
	
		if  inSTR(afe_Tipo,"O") <> 0 then
			Call GF_writeTextAlign(Gbl_oPDF,col(1)+40, PxlFila, afe_TipoOtros, 590,PDF_ALIGN_LEFT)
		end if
end Function 
'----------------------------------------------------------------------------------------------------------
Function DibujarContenedor2FV()
	dim myCompania,myNo,myCuentaGeneral,myResp,myDept,mySub,myLoc,myDiv,myDepartment,myLocation,myJob,myAfe
	 '-----------------------------------------------------------------------------------
	 'Dibujo segundo contenedor
	 Call GF_horizontalLine(Gbl_oPDF,5 ,150,585)
	 Call GF_horizontalLine(Gbl_oPDF,5 ,180,585)
	 Call GF_horizontalLine(Gbl_oPDF,5 ,210,585)
	 Call GF_horizontalLine(Gbl_oPDF,5 ,240,585)
	 Call GF_verticalLine  (Gbl_oPDF,5  ,150,90)
	 Call GF_verticalLine  (Gbl_oPDF,245,150,60)
	 Call GF_verticalLine  (Gbl_oPDF,275,150,30)
	 Call GF_verticalLine  (Gbl_oPDF,405,150,60)
	 Call GF_verticalLine  (Gbl_oPDF,445,150,30)
	 Call GF_verticalLine  (Gbl_oPDF,495,150,60)
	 Call GF_verticalLine  (Gbl_oPDF,545,150,30)
	 Call GF_verticalLine  (Gbl_oPDF,590,150,90)
	 '-----------------------------------------------------------------------------------	
	
	myCompania	= getDescripcionProveedor(CD_TOEPFER)
	myNo        = "-"
	mySub       = "-"
	myLoc       = "-"
	
	if isnull(afe_ObraCuentaDS) then
		myCuentaGeneral = "-"
	else
		myCuentaGeneral = afe_ObraCuentaDS
	end if
	
	if afe_ObraRespCD = "" then
		myResp= "-"
	else
		myResp= afe_ObraRespCD
	end if
	
	if afe_RespSectorDS = "" then
		myDept= "-"
	else
		myDept= afe_RespSectorDS
	end if
	
	if afe_ObraDivDs = "" then
		myDiv= "-"
	else
		myDiv= ucase(afe_ObraDivDs)
		if (afe_ImporteDolares >= MAX_LOCAL_LIMITFV) then
			myDiv= "ARGENTINA (" & myDiv & ")"
		end if
		
	end if
	
	if afe_Departamento = "" then
		myDepartment= "-"
	else
		myDepartment= ucase(afe_Departamento)
	end if

	if afe_IdDivision = "0" then
		myLocation= "-"
	else
		myLocation= getLocation(afe_IdDivision)
		if myLocation = "" then myLocation = "-"
	end if
	
	if afe_ObraCD = "" then
		myJob= "-"
	else
		if (isnull(afe_IDArea)) then
			myJob= afe_ObraCD
		else
			myJob= afe_ObraCD & "-" & afe_IDArea & "-" & afe_IDDetalle
		end if
	end if
	
	if afe_Titulo = "" then
		myAfe= "-"
	else
		myAfe= ucase(afe_Titulo)
	end if
	
	
	'**********************
	'TITULOS
	Call GF_setFont(Gbl_oPDF,"ARIAL", 8,8)	
	Call GF_writeTextAlign(Gbl_oPDF,10 , 151, "COMPANY"        	, 590,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,250, 151, "NO."            	, 590,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,280, 151, "GENERAL ACCOUNT"	, 590,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,410, 151, "RESP."          	, 590,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,450, 151, "DEPT."          	, 590,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,500, 151, "SUB."           	, 590,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,550, 151, "LOC."           	, 590,PDF_ALIGN_LEFT)
	
	Call GF_writeTextAlign(Gbl_oPDF,10 , 181, "DIVISION"          	, 590,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,250, 181, "DEPARTMENT"        	, 590,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,410, 181, "LOCATION"          	, 590,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,497, 181, "JOB/WORK ORDER NO."	, 590,PDF_ALIGN_LEFT)
	
	Call GF_writeTextAlign(Gbl_oPDF,10 , 211, "AFE TITLE"			, 590,PDF_ALIGN_LEFT)
	
	'**********************************
	
	Call GF_setFont(Gbl_oPDF,"ARIAL", 8,0)
	
	'Completo Primera Fila
	Call GF_writeTextAlign(Gbl_oPDF,10 , 165, myCompania      ,245,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(Gbl_oPDF,245, 165, myNo            ,30 ,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(Gbl_oPDF,275, 165, myCuentaGeneral ,130,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(Gbl_oPDF,405, 165, myResp          ,40 ,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(Gbl_oPDF,445, 165, myDept          ,50 ,PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(Gbl_oPDF,495, 165, mySub           ,50 ,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(Gbl_oPDF,545, 165, myLoc           ,50 ,PDF_ALIGN_CENTER)
		
	'Completo Segunda Fila
	Call GF_writeTextAlign(Gbl_oPDF,10 , 195, myDiv           ,245,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(Gbl_oPDF,245, 195, myDepartment    ,160,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(Gbl_oPDF,405, 195, trim(myLocation),90 ,PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(Gbl_oPDF,495, 195, myJob           ,100,PDF_ALIGN_CENTER)

	
	'Completo tercer fila
	Call GF_writeTextAlign(Gbl_oPDF,10 , 225, myAfe           ,580,PDF_ALIGN_LEFT)

end Function
'----------------------------------------------------------------------------------------------------------
Function DibujarContenedor3FV()

	Dim total,local,code,rate,arr,irr,rona,payback,genericSignature	
	

	total   = GF_EDIT_DECIMALS(cdbl(afe_ImporteDolares),2)
	local   = GF_EDIT_DECIMALS(cdbl(afe_ImportePesos)  ,2)
	code    = getSimboloMonedaLetras(MONEDA_PESO)
	rate    = afe_TipoCambio
	if afe_Arr = "0" then
		arr     = "NA"
	else
		arr     = GF_EDIT_DECIMALS(cdbl(afe_Arr),2)
	end if
	if afe_Irr = "0" then
		irr     = "NA"
	else
		irr     = GF_EDIT_DECIMALS(cdbl(afe_Irr),2)
	end if
	if afe_RONA = "" then
		rona    = "NA"
	else
		rona    = GF_EDIT_DECIMALS(cdbl(afe_RONA),2)
	end if
	
	if afe_PAYBACK = "" then
		payback = "NA"
	else
		payback = GF_EDIT_DECIMALS(cdbl(afe_PAYBACK),2)
	end if
	
	Call GF_setFont(Gbl_oPDF,"ARIAL", 8,8)	
	Call GF_writeTextAlign(Gbl_oPDF,5 , 420,"NOTE:" , 590,PDF_ALIGN_LEFT)
	
	Call GF_setFont(Gbl_oPDF,"ARIAL", 8,0)	
	Call GF_writeTextPlus(Gbl_oPDF,5, 430, "NOTIFY TREASURER IF PROJECT WILL INCLUDE PURCHASES PAYABLE IN NON-LOCAL CURRENCY.<br>NOTIFY GLOBAL FIXED ASSET ACCOUNTING BY APPROPIATE FROM OF THE STATUS OF ANY EQUIPMENT THAT IS BEING REPLACED.", 180, 8, PDF_ALIGN_LEFT)
	
	'***********************************************************
	'dibujo contenedores
	'precios
	
	Call GF_squareBox(Gbl_oPDF,5,520,180,30,0,"#FFFFFF","#000000",1,0) 
	Call GF_squareBox(Gbl_oPDF,5,550,180,30,0,"#FFFFFF","#000000",1,0) 
	Call GF_squareBox(Gbl_oPDF,5,580,180,30,0,"#FFFFFF","#000000",1,0) 
	Call GF_squareBox(Gbl_oPDF,5,610,180,30,0,"#FFFFFF","#000000",1,0) 
	Call GF_squareBox(Gbl_oPDF,5,640,180,30,0,"#FFFFFF","#000000",1,0) 
	
	Call GF_verticalLine (Gbl_oPDF,95,580,90)
		
	'firmas
	

	Call GF_squareBox(Gbl_oPDF,215,410,375,10,0,"#FFFFFF","#000000",1,0) 
	
	Call GF_squareBox(Gbl_oPDF,215,420,375,50,0,"#FFFFFF","#000000",1,0) 
	Call GF_squareBox(Gbl_oPDF,215,470,375,50,0,"#FFFFFF","#000000",1,0) 
	Call GF_squareBox(Gbl_oPDF,215,520,375,50,0,"#FFFFFF","#000000",1,0) 
	Call GF_squareBox(Gbl_oPDF,215,570,375,50,0,"#FFFFFF","#000000",1,0) 
	Call GF_squareBox(Gbl_oPDF,215,620,375,50,0,"#FFFFFF","#000000",1,0) 
	
	Call GF_verticalLine (Gbl_oPDF,355,420,250)
	Call GF_verticalLine (Gbl_oPDF,405,420,250)
	Call GF_verticalLine (Gbl_oPDF,545,420,250)
	
	'dibujo las firmas
	genericSignature = Server.MapPath("Images\Firmas\signature-48x48.png")
	
	Call GF_setFont(Gbl_oPDF,"ARIAL", 4,0)
	if (afe_PreparedByHkey <> "")  then
		if not esAManoFV(afe_PreparedByHkey) then Call GF_writeImage(Gbl_oPDF, server.MapPath(".") & "\images\firmas\" & obtenerFirma(afe_PreparedByCD)    , 245, 428, 80, 30, 0)
		Call GF_writeTextAlign(Gbl_oPDF,215, 465,armarTextoPlanoFirma(afe_PreparedByHkey, afe_PreparedByHkeyDate), 140,PDF_ALIGN_CENTER)
	end if
	if (afe_OfficerHkey <> "") then
		if not esAManoFV(afe_OfficerHkey) then Call GF_writeImage(Gbl_oPDF, server.MapPath(".") & "\images\firmas\" & obtenerFirma(afe_OfficerCD) 	   , 435, 428, 80, 30, 0)
		Call GF_writeTextAlign(Gbl_oPDF,405, 465,armarTextoPlanoFirma(afe_OfficerHkey, afe_OfficerHkeyDate), 140,PDF_ALIGN_CENTER)
	end if
	if (afe_RequestedByHkey <> "") then
		if not esAManoFV(afe_RequestedByHkey) then Call GF_writeImage(Gbl_oPDF, server.MapPath(".") & "\images\firmas\" & obtenerFirma(afe_RequestedByCD)   , 245, 478, 80, 30, 0)
		Call GF_writeTextAlign(Gbl_oPDF,215, 515,armarTextoPlanoFirma(afe_RequestedByHkey, afe_RequestedByHkeyDate), 140,PDF_ALIGN_CENTER)
	end if
	if (afe_VicePresidentHkey <> "") then
		if not esAManoFV(afe_VicePresidentHkey) then Call GF_writeImage(Gbl_oPDF, server.MapPath(".") & "\images\firmas\" & obtenerFirma(afe_VicePresidentCD) , 435, 478, 80, 30, 0)
		Call GF_writeTextAlign(Gbl_oPDF,405, 515,armarTextoPlanoFirma(afe_VicePresidentHkey, afe_VicePresidentHkeyDate), 140,PDF_ALIGN_CENTER)
	end if
	if (afe_EngReviewHkey <> "") then
		if not esAManoFV(afe_EngReviewHkey) then Call GF_writeImage(Gbl_oPDF, server.MapPath(".") & "\images\firmas\" & obtenerFirma(afe_EngReviewCD)     , 245, 528, 80, 30, 0)
		Call GF_writeTextAlign(Gbl_oPDF,215, 565,armarTextoPlanoFirma(afe_EngReviewHkey, afe_EngReviewHkeyDate), 140,PDF_ALIGN_CENTER)
	end if
	if (afe_ControllerHkey <> "") then
		if not esAManoFV(afe_ControllerHkey) then Call GF_writeImage(Gbl_oPDF, server.MapPath(".") & "\images\firmas\" & obtenerFirma(afe_ControllerCD)    , 245, 578, 80, 30, 0)
		Call GF_writeTextAlign(Gbl_oPDF,215, 615,armarTextoPlanoFirma(afe_ControllerHkey, afe_ControllerHkeyDate), 140,PDF_ALIGN_CENTER)
	end if
	if (afe_AuditorHkey <> "") then
		if not esAManoFV(afe_AuditorHkey) then Call GF_writeImage(Gbl_oPDF, server.MapPath(".") & "\images\firmas\" & obtenerFirma(afe_AuditorCD)       , 245, 628, 80, 30, 0)
		Call GF_writeTextAlign(Gbl_oPDF,215, 665,armarTextoPlanoFirma(afe_AuditorHkey, afe_AuditorHkeyDate), 140,PDF_ALIGN_CENTER)
	end if
	'**********************************************************
	'texto precios
	Call GF_setFont(Gbl_oPDF,"ARIAL", 8,8)	
	Call GF_writeTextAlign(Gbl_oPDF,5   , 522, "TOTAL EXPENDITURE(US$)", 180,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(Gbl_oPDF,5   , 552, "LOCAL CURRENCY AMOUNT" , 180,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(Gbl_oPDF,10  , 582, "CURRENCY CODE"         , 90 ,PDF_ALIGN_LEFT  )
	Call GF_writeTextAlign(Gbl_oPDF,100 , 582, "EXCHANGE RATE"         , 90 ,PDF_ALIGN_LEFT  )
	Call GF_writeTextAlign(Gbl_oPDF,10  , 612, "ARR"                   , 90 ,PDF_ALIGN_LEFT  )
	Call GF_writeTextAlign(Gbl_oPDF,100 , 612, "IRR"                   , 90 ,PDF_ALIGN_LEFT  )
	Call GF_writeTextAlign(Gbl_oPDF,10  , 642, "RONA"                  , 90 ,PDF_ALIGN_LEFT  )
	Call GF_writeTextAlign(Gbl_oPDF,100 , 642, "PAYBACK"               , 90 ,PDF_ALIGN_LEFT  )
	
	Call GF_setFont(Gbl_oPDF,"ARIAL", 8,0)	
	Call GF_writeTextAlign(Gbl_oPDF,5  , 540, total                   , 180,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(Gbl_oPDF,5  , 570, local                   , 180,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(Gbl_oPDF,5  , 600, code                    , 90 ,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(Gbl_oPDF,90 , 600, RATE                    , 90 ,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(Gbl_oPDF,5  , 630, ARR                     , 90 ,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(Gbl_oPDF,90 , 630, IRR                     , 90 ,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(Gbl_oPDF,5  , 660, RONA                    , 90 ,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(Gbl_oPDF,90 , 660, PAYBACK                 , 90 ,PDF_ALIGN_CENTER)
	'**********************************************************
	'texto firmas
	Call GF_setFont(Gbl_oPDF,"ARIAL", 8,8)	
	Call GF_writeTextAlign(Gbl_oPDF,215, 410, "REVIEW AND APPROVAL",375,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(Gbl_oPDF,355, 421, "DATE"               , 50,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(Gbl_oPDF,545, 421, "DATE"               , 50,PDF_ALIGN_CENTER)
	
	Call GF_writeTextAlign(Gbl_oPDF,220, 421, "ESTIMATE PREPARED BY"      , 140,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,410, 421, "OFFICER IN CHARGE"         , 140,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,220, 471, "EXPENDITURE REQUESTED BY"  , 140,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,410, 471, "VICE PRESIDENT"			  , 140,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,220, 520, "ENGINEERING REVIEW"        , 140,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,410, 520, "PRESIDENT"                 , 140,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,220, 571, "CONTROLLER  REVIEW"        , 140,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,410, 571, "GENERAL MANAGER ACTI"      , 140,PDF_ALIGN_LEFT)
	'Auditoria no firma mas desde el 08/08/2012. Se conserva su por compatibilidad con AFEs viejos.
	if (afe_AuditorHkey <> "") then	Call GF_writeTextAlign(Gbl_oPDF,220, 620, "AUDIT DEPARTMENT"		  , 140,PDF_ALIGN_LEFT)
	
	Call GF_setFont(Gbl_oPDF,"ARIAL", 6.8,8)	
	Call GF_writeTextAlign(Gbl_oPDF,410, 620, "MANAGING DIRECTOR " , 140,PDF_ALIGN_LEFT)		
	Call GF_writeTextAlign(Gbl_oPDF,410, 630, "FINANCE, BUSINESS DEV, TAX OF ACTI", 140,PDF_ALIGN_LEFT)
	
	Call GF_setFont(Gbl_oPDF,"ARIAL", 6,0)	
	
	Call GF_writeTextAlign(Gbl_oPDF,215, 458, afe_PreparedBy   , 140,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(Gbl_oPDF,405, 458, afe_Officer      , 140,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(Gbl_oPDF,215, 508, afe_RequestedBy  , 140,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(Gbl_oPDF,405, 508, afe_VicePresident, 140,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(Gbl_oPDF,215, 558, afe_EngReview    , 140,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(Gbl_oPDF,405, 558, afe_President    , 140,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(Gbl_oPDF,215, 608, afe_Controller   , 140,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(Gbl_oPDF,405, 608, "Ulrich Litterscheid", 140,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(Gbl_oPDF,215, 658, afe_Auditor      , 140,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(Gbl_oPDF,405, 658, "Kevin Brassington", 140,PDF_ALIGN_CENTER)
	
	'fecha firmas
	Call dibujarFechaFirmas()
		
	'**********************************************************
end Function
'----------------------------------------------------------------------------------------------------------
Function dibujarFechaFirmas()
	Call GF_setFont(Gbl_oPDF,"ARIAL", 10,0)
	if (afe_PreparedByHkeyDate <> "")	then Call GF_writeTextAlign(Gbl_oPDF,355, 455, left(GF_FN2DTE(left(afe_PreparedByHkeyDate,8)),5), 50,PDF_ALIGN_CENTER)
	if (afe_OfficerHkeyDate <> "")		then Call GF_writeTextAlign(Gbl_oPDF,545, 455, left(GF_FN2DTE(left(afe_OfficerHkeyDate,8)),5), 50,PDF_ALIGN_CENTER)
	if (afe_RequestedByHkeyDate <> "")	then Call GF_writeTextAlign(Gbl_oPDF,355, 505, left(GF_FN2DTE(left(afe_RequestedByHkeyDate,8)),5), 50,PDF_ALIGN_CENTER)
	if (afe_VicePresidentHkeyDate <> "")then Call GF_writeTextAlign(Gbl_oPDF,545, 505, left(GF_FN2DTE(left(afe_VicePresidentHkeyDate,8)),5), 50,PDF_ALIGN_CENTER)
	if (afe_EngReviewHkeyDate <> "")	then Call GF_writeTextAlign(Gbl_oPDF,355, 555, left(GF_FN2DTE(left(afe_EngReviewHkeyDate,8)),5), 50,PDF_ALIGN_CENTER)
	if (afe_ControllerHkeyDate <> "")	then Call GF_writeTextAlign(Gbl_oPDF,355, 605, left(GF_FN2DTE(left(afe_ControllerHkeyDate,8)),5), 50,PDF_ALIGN_CENTER)
	if (afe_AuditorHkeyDate <> "")		then Call GF_writeTextAlign(Gbl_oPDF,355, 655, left(GF_FN2DTE(left(afe_AuditorHkeyDate,8)),5), 50,PDF_ALIGN_CENTER)
end Function
'----------------------------------------------------------------------------------------------------------
Function DibujarContenedor4FV()
	Dim espacio,inicio,aux
	espacio = 10
	inicio = 680
	
	'****************************************
	'dibujo cuarto contenedor
	'lineas horizontales
	Call GF_horizontalLine(Gbl_oPDF,5 ,inicio+(espacio),585)
	Call GF_horizontalLine(Gbl_oPDF,5 ,inicio+(espacio*2),585)
	Call GF_horizontalLine(Gbl_oPDF,5 ,inicio+(espacio*3),585)
	Call GF_horizontalLine(Gbl_oPDF,5 ,inicio+(espacio*5),585)
	Call GF_horizontalLine(Gbl_oPDF,5 ,inicio+(espacio*6),585)
	Call GF_horizontalLine(Gbl_oPDF,5 ,inicio+(espacio*8),585)
	Call GF_horizontalLine(Gbl_oPDF,5 ,inicio+(espacio*9),585)
	Call GF_horizontalLine(Gbl_oPDF,5 ,inicio+(espacio*10),585)
	Call GF_horizontalLine(Gbl_oPDF,5 ,inicio+(espacio*11),585)
	Call GF_horizontalLine(Gbl_oPDF,5 ,inicio+(espacio*12),585)
	Call GF_horizontalLine(Gbl_oPDF,5 ,inicio+(espacio*13),585)
	Call GF_horizontalLine(Gbl_oPDF,5 ,inicio+(espacio*14),585)
	Call GF_horizontalLine(Gbl_oPDF,5 ,inicio+(espacio*16),585)
		
	'lineas verticales
	call GF_verticalLine  (Gbl_oPDF,5,inicio+espacio,150)
	call GF_verticalLine  (Gbl_oPDF,590,inicio+espacio,150)
	'*****************************************
	
	Call GF_setFont(Gbl_oPDF,"ARIAL", 8,8)	
	Call GF_writeTextAlign(Gbl_oPDF,10 , inicio+espacio+1, "ADDITIONAL AFE INFORMATION", 590,PDF_ALIGN_LEFT)
	
	'escribo las preguntas
	Call GF_setFont(Gbl_oPDF,"ARIAL", 8,0)	
	Call GF_writeTextAlign(Gbl_oPDF,10 , inicio+(espacio*2)+1, "WAS GLOBAL SOURCING USED?", 590,PDF_ALIGN_LEFT)
	
	Call GF_writeTextAlign(Gbl_oPDF,10,  inicio+(espacio*3)+1,"WILL THIS PROJECT REQUIRE ADDITIONAL INCREMENTAL WORKING CAPITAL (I.E. MORE ",590,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,10,  inicio+(espacio*4)+1,"INVENTORY,MORERECEIVABLES)?",590,PDF_ALIGN_LEFT)
	
	Call GF_writeTextAlign(Gbl_oPDF,10 , inicio+(espacio*5)+1,"WAS CAPITALIZED INTEREST INCLUDED IN PROJECT ESTIMATE?" , 590,PDF_ALIGN_LEFT)
	
	Call GF_writeTextAlign(Gbl_oPDF,10 , inicio+(espacio*6 )+1,"HAS PROJECT BEEN PR-APPROVED THROGH BOARD OF DIRECTORS RESOLUTION OR BUSINESS PLAN APPROVAL?" , 590,PDF_ALIGN_LEFT)
	Call GF_setFont(Gbl_oPDF,"ARIAL", 8,4)	
	Call GF_writeTextAlign(Gbl_oPDF,15 , inicio+(espacio*7 )+1,afe_Preg4_Text , 590,PDF_ALIGN_LEFT)
	Call GF_setFont(Gbl_oPDF,"ARIAL", 8,0)	
	Call GF_writeTextAlign(Gbl_oPDF,10 , inicio+(espacio*8 )+1,"DOES PROJECT ANTICIPED TRADE IN VALUE OF ASSETS?" , 590,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,10 , inicio+(espacio*9 )+1,"WAS THERE A LEGAL DEPARTMENT REVIEW?" , 590,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,10 , inicio+(espacio*10)+1,"DOES THE FINANCIAL CALCULATION DEPEND OF HEAD COUNT?" , 590,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,10 , inicio+(espacio*11)+1,"IS THE SPEND IN A FOREIGN CURRENCY?" , 590,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,10 , inicio+(espacio*12)+1,"ARE THE QUOTES LESS THAN 90 DAYS OLD (AFEs GREATER THAN $5,000,000 ONLY)?" , 590,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,10 , inicio+(espacio*13)+1,"ARE QUOTES INTERNAL OR EXTERNAL?" , 590,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(Gbl_oPDF,10 , inicio+(espacio*14)+1,"WHAT ARE THE INTERNAL COST OF THE PROJECT IN US DOLLARS?" , 590,PDF_ALIGN_LEFT)
	Call GF_setFont(Gbl_oPDF,"ARIAL", 8,4)	
	Call GF_writeTextAlign(Gbl_oPDF,15 , inicio+(espacio*15)+1,afe_Preg11 , 590,PDF_ALIGN_LEFT)
	Call GF_setFont(Gbl_oPDF,"ARIAL", 8,0)
	'pongo los radios
	Dim inicioRadios
	inicioRadios = 500
	
	Call DibujarPreguntasFV(afe_Preg1,espacio*2)
	Call DibujarPreguntasFV(afe_Preg2,espacio*3)
	Call DibujarPreguntasFV(afe_Preg3,espacio*5)
	Call DibujarPreguntasFV(afe_Preg4,espacio*6)
	Call DibujarPreguntasFV(afe_Preg5,espacio*8)
	Call DibujarPreguntasFV(afe_Preg6,espacio*9)
	Call DibujarPreguntasFV(afe_Preg7,espacio*10)
	Call DibujarPreguntasFV(afe_Preg8,espacio*11)
	Call DibujarPreguntasFV(afe_Preg9,espacio*12)
	
	
	'pregunta 10 ********************************************
	if afe_Preg10 = "I" then
		Call DibujarRadioFV("INT",inicioRadios    ,inicio+(espacio*13)+1,1)
	else
		Call DibujarRadioFV("INT",inicioRadios    ,inicio+(espacio*13)+1,0)
	end if
	if afe_Preg10 = "E" then
		Call DibujarRadioFV("EXT" ,inicioRadios+30,inicio+(espacio*13)+1,1)
	else
		Call DibujarRadioFV("EXT" ,inicioRadios+30,inicio+(espacio*13)+1,0)
	end if
	if afe_Preg10 = "A" then
		Call DibujarRadioFV("NA" ,inicioRadios+60 ,inicio+(espacio*13)+1,1)
	else
		Call DibujarRadioFV("NA" ,inicioRadios+60 ,inicio+(espacio*13)+1,0)
	end if
	
end Function
'----------------------------------------------------------------------------------------------------------
Function DibujarPreguntasFV(Respuesta,Espacio)
	dim inicio
	Dim inicioRadios
	inicioRadios = 500
	inicio = 680
	if Respuesta = "S" then
		Call DibujarRadioFV("YES",inicioRadios    ,inicio+Espacio+1,1)
	else
		Call DibujarRadioFV("YES",inicioRadios    ,inicio+Espacio+1,0)
	end if
	if Respuesta = "N" then
		Call DibujarRadioFV("NO" ,inicioRadios+30 ,inicio+Espacio+1,1)
	else
		Call DibujarRadioFV("NO" ,inicioRadios+30 ,inicio+Espacio+1,0)
	end if
	if Respuesta = "A" then
		Call DibujarRadioFV("NA" ,inicioRadios+60 ,inicio+Espacio+1,1)
	else
		Call DibujarRadioFV("NA" ,inicioRadios+60 ,inicio+Espacio+1,0)
	end if
end Function
'----------------------------------------------------------------------------------------------------------
Function DibujarDescripcionFV()	
	Dim vecRenglones, idxRenglones, idxPagina, idxRenglonesPagina, vecPaginas(), textoPagina
	
	Call GF_setFont(Gbl_oPDF,"ARIAL", 8,8)	
	Call GF_writeTextAlign(Gbl_oPDF,5 , 250, "DESCRIPTION:", 590,PDF_ALIGN_LEFT)
	Call GF_setFont(Gbl_oPDF,"ARIAL", 8,0)			

	vecRenglones = generarRenglones(afe_Descripcion, ENTER_SYMBOL, 160)	
	'Calculo la pantidad de paginas			
	if (UBound(vecRenglones) <= MAX_RENGLONESFV_PAGINA1FV) then
		Redim vecPaginas(0)
	else
		Redim vecPaginas(Ceil((UBound(vecRenglones)-MAX_RENGLONESFV_PAGINA1FV)/MAX_RENGLONESFV))
	end if
	'Se arman las paginas	
	idxPagina = 0
	idxRenglones = 0
	while (idxRenglones <= UBound(vecRenglones))
		'Determino el maximo nro de renglones para la pagina.
		maxRenglones = MAX_RENGLONESFV
		if (idxPagina = 0) then maxRenglones = MAX_RENGLONESFV_PAGINA1FV
		'Armo el texto de la pagina.
		textoPagina=""
		idxRenglonesPagina=0		
		while ((idxRenglones <= UBound(vecRenglones)) and (idxRenglonesPagina <= maxRenglones))
			textoPagina = textoPagina & vecRenglones(idxRenglones)			
			idxRenglones = idxRenglones + 1
			idxRenglonesPagina = idxRenglonesPagina + 1
		wend		
		vecPaginas(idxPagina)= textoPagina
		idxPagina = idxPagina + 1
	wend
	Call ImprimirDescripcionFV(vecPaginas)
end Function
'----------------------------------------------------------------------------------------------------------
Function ImprimirDescripcionFV(vDS)
	Dim yInicial
	
	
	Call GF_setFont(Gbl_oPDF,"ARIAL", 8,0)
	yInicial = 260 
	for i = 0 to ubound(vDS)-1	
		Call GF_writeTextPlus(Gbl_oPDF,10, yInicial, vDS(i), 570, 8, PDF_ALIGN_LEFT)				
		yInicial = NuevaHojaFV()		
	next	
	Call GF_writeTextPlus(Gbl_oPDF,10, yInicial, vDS(ubound(vDS)), 570, 8, PDF_ALIGN_LEFT)
end Function
'-----------------------------------------------------------------------------------------------------------
Function DibujarRadioFV(texto,x,y,chequeado)
	Call GF_writeImage(Gbl_oPDF, Server.MapPath("Images\Radio_Chk"&chequeado&".gif"),x,y,8,8, 0)
	
	if texto <> "" then
		Call GF_writeTextAlign(Gbl_oPDF,x+10, y, texto, 590,PDF_ALIGN_LEFT)
	end if
end Function
'-----------------------------------------------------------------------------------------------------------
Function DibujarCheckFV(x,y,chequeado)
	Call GF_writeImage(Gbl_oPDF, Server.MapPath("Images\Item_Chk"&chequeado&".gif"),x,y,8,8, 0)
end Function
'----------------------------------------------------------------------------------------------------------
Function DibujarReglaFV(inicio,separacion,y)
	Dim resu,i
	for i = 0 to 588 step separacion
		Call GF_verticalLine  (Gbl_oPDF,inicio+i,y,10)
	next
end Function
'----------------------------------------------------------------------------------------------------------
Function NuevaHojaFV()
	nroHojas = nroHojas +1
	Call GF_newPage(Gbl_oPDF)	
	Call GF_setFont(Gbl_oPDF,"ARIAL", 8,8)
	Call GF_writeTextAlign(Gbl_oPDF,5 ,840, "Page " & nroHojas, 580,PDF_ALIGN_RIGHT)
	Call DibujarEncabezadoFV()
	Call DibujarMarcaDeAguaFV() 	 
	Call GF_setFont(Gbl_oPDF,"ARIAL", 16,8)
	Call GF_squareBox(Gbl_oPDF,5,50,580 ,790,0,"#FFFFFF","#000000",1,PDF_SQUARE_ROUND) 	
	Call GF_writeTextAlign(Gbl_oPDF,5 ,50, "DETAIL", 590,PDF_ALIGN_CENTER)		
	NuevaHojaFV=70 'Devuelve la linea inicial de las pagnas complementarias del AFE.
end Function
'--------------------------------------------------------------------------------------
Function DibujarMarcaDeAguaFV()
 if (afe_Confirmado = "R") then
	'el afe esta rechazado
	Call GF_writeImage(Gbl_oPDF, Server.MapPath("Images\compras\AFE_canceled_wathermark.png"),90, 220, 500, 100, 335)
 end if
end function
'--------------------------------------------------------------------------------------
function esAManoFV(pValue)
	if pValue = A_MANO then esAManoFV = true
end function

%>


