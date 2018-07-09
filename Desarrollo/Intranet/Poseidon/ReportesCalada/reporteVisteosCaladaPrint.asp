<!--#include file="../../Includes/procedimientosUnificador.asp"-->
<!--#include file="../../Includes/procedimientosParametros.asp"-->
<!--#include file="../../Includes/procedimientosPuertos.asp"-->
<!--#include file="../../Includes/procedimientosSQL.asp"-->
<!--#include file="../../Includes/procedimientos.asp"-->
<!--#include file="../../Includes/procedimientosPDF.asp"-->
<!--#include file="../../Includes/procedimientostraducir.asp"-->
<!--#include file="../../Includes/procedimientosfechas.asp"-->
<!--#include file="../../Includes/procedimientosuser.asp"-->
<!--#include file="../../Includes/procedimientosMG.asp"-->
<!--#include file="../../Includes/procedimientosFormato.asp"-->
<!--#include file="reporteVisteosCaladaCommon.asp"-->

<%
'-----------------------------------------------------------------------------------------
Function getNombreRubroCalada(pcdRubro, pto)
	dim strSQL, rs, rtrn
	strSQL = "Select DSRUBRO from dbo.RUBROS where cdrubro = "&pcdRubro
	GF_BD_Puertos pto, rs, "OPEN",strSQL
	rtrn = pcdRubro
	if (not rs.eof) then rtrn = rs("DSRUBRO")
	getNombreRubroCalada = rtrn
end Function
'-----------------------------------------------------------------------------------------
Function armarTitulos(pidcamion, ppto, pnuCartaPorte, codVerificador, pemision, phasta,pcdProducto,pcdVendedor,pcdCorredor,pcdCliente,pcdEntregador, pcdEstado)
	dim strSQL, rs, myWhere, diaHoy, filtroEstado
	'Analizo los filtros.
	if (pnuCartaPorte <> "") 	then Call mkWhere(myWhere, "A.NUCARTAPORTE", pnuCartaPorte, "LIKE", 3)								  
	if (codVerificador <> "") 	then Call mkWhere(myWhere, "A.NUCTAPTEDIG", codVerificador, "LIKE", 3)
	if (pidcamion > 0)			then Call mkWhere(myWhere, "A.IDCAMION", pidcamion, "=", 0)
	if (pemision <> "") 		then Call mkWhere(myWhere, "A.DTCONTABLE", pemision, ">=", 0)
	if (phasta <> "") 			then Call mkWhere(myWhere, "A.DTCONTABLE", phasta, "<=", 0)	
	if (pcdProducto <> 0) 		then Call mkWhere(myWhere, "A.CDPRODUCTO", pcdProducto, "=", 1)
	if (pcdCliente <> "")	    then Call mkWhere(myWhere, "A.CDCLIENTE", pcdCliente, "=", 1)
	if (pcdCorredor <> "") 		then Call mkWhere(myWhere, "A.CDCORREDOR", pcdCorredor, "=", 1)
	if (pcdVendedor <> "") 		then Call mkWhere(myWhere, "A.CDVENDEDOR", pcdVendedor, "=", 1)
	if (pcdEntregador <> "")	then Call mkWhere(myWhere, "A.CDENTREGADOR", pcdEntregador, "=", 1)
	if (pcdEntregador <> "")	then Call mkWhere(myWhere, "A.ESTADO", pcdEntregador, "=", 1)
	filtroEstado = CAMIONES_ESTADO_EGRESADOOK & ", " & CAMIONES_ESTADO_PESADOTARA
	if (CInt(pcdEstado) <> 0)	then filtroEstado = pcdEstado
	
	'Armo la fecha del dia la SQL diaria.
	diaHoy = Year(Now()) & GF_nDigits(Month(Now()), 2) & GF_nDigits(Day(Now()), 2) 
	'Consulta a la tabla diaria.
	strSQL = "Select"
	strSQL = strSQL & "	A.NUCARTAPORTE,"
	strSQL = strSQL & " A.NUCTAPTEDIG,"
	strSQL = strSQL & " A.CDCLIENTE,"
	strSQL = strSQL & "	A.CDCHAPACAMION,"
	strSQL = strSQL & " A.CDCHAPAACOPLADO,"
	strSQL = strSQL & " A.CDPRODUCTO,"
	strSQL = strSQL & "	A.IDCAMION,"	
	strSQL = strSQL & "	A.DTCONTABLE,"	
	strSQL = strSQL & " D.DSCORREDOR,A.CDCORREDOR," 
	strSQL = strSQL & "	F.DSVENDEDOR,A.CDVENDEDOR,"
	strSQL = strSQL & " G.DSENTREGADOR,A.CDENTREGADOR"	
	strSQL = strSQL & " from"
	
	strSQL = strSQL & "((SELECT '" & diaHoy & "' AS DTCONTABLE," 
	strSQL = strSQL & "	A.NUCARTAPORTE,A.NUCTAPTEDIG,"
	strSQL = strSQL & " A.CDCLIENTE,"
	strSQL = strSQL & "	B.CDCHAPACAMION,B.CDCHAPAACOPLADO,"
	strSQL = strSQL & " B.CDPRODUCTO,"
	strSQL = strSQL & "	A.IDCAMION,"	
	strSQL = strSQL & "	A.CDCORREDOR,"
	strSQL = strSQL & "	A.CDVENDEDOR,"
	strSQL = strSQL & " A.CDENTREGADOR"
	strSQL = strSQL & " FROM dbo.CAMIONESDESCARGA A"
	strSQL = strSQL & "	INNER JOIN dbo.CAMIONES B ON A.IDCAMION = B.IDCAMION"
	strSQL = strSQL & " where B.CDESTADO in (" & filtroEstado & "))"
	strSQL = strSQL & " UNION " 	
	strSQL = strSQL & " (SELECT (YEAR(A.DTCONTABLE)*10000 + Month(A.DTCONTABLE)*100 + DAY(A.DTCONTABLE)) DTCONTABLE,"
	strSQL = strSQL & "	A.NUCARTAPORTE,A.NUCTAPTEDIG,"
	strSQL = strSQL & " A.CDCLIENTE,"
	strSQL = strSQL & "	B.CDCHAPACAMION,B.CDCHAPAACOPLADO,"
	strSQL = strSQL & "	B.CDPRODUCTO,"
	strSQL = strSQL & "	A.IDCAMION,"
	strSQL = strSQL & "	A.CDCORREDOR,"
	strSQL = strSQL & "	A.CDVENDEDOR,"
	strSQL = strSQL & " A.CDENTREGADOR"
	strSQL = strSQL & " FROM dbo.HCAMIONESDESCARGA A"
	strSQL = strSQL & "	INNER JOIN dbo.HCAMIONES B ON A.IDCAMION = B.IDCAMION AND A.DTCONTABLE =B.DTCONTABLE"
	strSQL = strSQL & " where B.CDESTADO in (" & filtroEstado & "))) A"	
	
	strSQL = strSQL & "	INNER JOIN dbo.PRODUCTOS C ON A.CDPRODUCTO = C.CDPRODUCTO"			
	strSQL = strSQL & " INNER JOIN dbo.CORREDORES D ON A.CDCORREDOR = D.CDCORREDOR"
	strSQL = strSQL & "	INNER JOIN dbo.CLIENTES E ON A.CDCLIENTE = E.CDCLIENTE"
	strSQL = strSQL & "	INNER JOIN dbo.VENDEDORES F ON A.CDVENDEDOR = F.CDVENDEDOR"
	strSQL = strSQL & "	INNER JOIN dbo.ENTREGADORES G ON A.CDENTREGADOR= G.CDENTREGADOR"
	
	strSQL = strSQL & myWhere & " ORDER BY DTCONTABLE, IDCAMION"	

	Call GF_BD_Puertos(ppto, rs, "OPEN", strSQL)
	set armarTitulos = rs
end function
'-----------------------------------------------------------------------------------------
Function armarDetalle(psqcalada, pidcamion,pdtcontable,ppto) 		
	dim rs,myWhere,strSQL,diaHoy
	diaHoy = GF_nDigits(Year(Now()), 4)&GF_nDigits(Month(Now()), 2)&GF_nDigits(Day(Now()), 2) 
	if (pidcamion > 0) then Call mkWhere(myWhere, "IDCAMION", pidcamion, "=", 0)		
	
							Call mkWhere(myWhere, "SQCALADA", psqcalada, "=", 1)
							Call mkWhere(myWhere, "DTCONTABLE", pdtcontable, "=", 0)		
	strSQL = "Select * from "
	strSQL = strSQL & " ((select '" & diaHoy & "' AS DTCONTABLE ,cdrubro, vlbonrebaja, vlmerma, vlpesorubro, pcpesorubro,idcamion,sqcalada "
	strSQL = strSQL & " FROM dbo.audrubrosvisteocamiones A) "
	strSQL = strSQL & " UNION "
	strSQL = strSQL & " (select (YEAR(DtContable)*10000 + Month(DtContable)*100 + DAY(DtContable)) DTCONTABLE,cdrubro, vlbonrebaja, vlmerma, vlpesorubro, pcpesorubro,idcamion,sqcalada "
	strSQL = strSQL & " FROM dbo.haudrubrosvisteocamiones A)) as Tabla "
	strSQL = strSQL & myWhere
	Call GF_BD_Puertos(ppto, rs, "OPEN", strSQL)
Set armarDetalle = rs
end function
'-----------------------------------------------------------------------------------------
Function armarCabecera(v_idcamion,pto,fecha)
	dim rs,myWhere,strSQL	
	if(v_idcamion > 0)then Call mkWhere(myWhere, "idcamion", v_idcamion, "=", 0)	
	if (fecha <> "") then Call mkWhere(myWhere, "DTCONTABLE", fecha, "=", 0)		 
	'Armo la fecha del dia la SQL diaria.
    diaHoy = Year(Now()) & GF_nDigits(Month(Now()), 2) & GF_nDigits(Day(Now()), 2) 
	strSQL = "Select * from "
	strSQL = strSQL & "((SELECT '" & diaHoy & "' AS DTCONTABLE, A.sqauditoria, A.SQCALADA ,B.CDUSERNAME, B.CDTERMINAL,A.IDCAMION, C.CDSUPERVISOR "
	strSQL = strSQL & " FROM  dbo.audcaladadecamiones A INNER JOIN dbo.auditoriacamiones B "
	strSQL = strSQL & " ON A.idcamion = B.idcamion "	
	strSQL = strSQL & " AND A.sqauditoria = B.sqauditoria"
	strSQL = strSQL & " AND B.cdtransaccion = " & VISTEO_CALADA
	strSQL = strSQL & " left join AuditoriaCamiones C on A.IDCAMION=C.IDCAMION and A.SQAUDITORIA=C.SQAUDITORIA)"
	strSQL = strSQL & " UNION "
	strSQL = strSQL & "(SELECT (YEAR(A.DTCONTABLE)*10000 + Month(A.DTCONTABLE)*100 + DAY(A.DTCONTABLE)) dtcontable, A.sqauditoria, A.SQCALADA ,B.CDUSERNAME, B.CDTERMINAL,A.IDCAMION, C.CDSUPERVISOR "
	strSQL = strSQL & " FROM  dbo.haudcaladadecamiones A INNER JOIN dbo.hauditoriacamiones B "
	strSQL = strSQL & " ON A.idcamion = B.idcamion"	
	strSQL = strSQL & " AND A.sqauditoria = B.sqauditoria "
	strSQL = strSQL & " AND B.cdtransaccion = " & VISTEO_CALADA
	strSQL = strSQL & " AND A.dtcontable = B.dtcontable"
	strSQL = strSQL & " left join HAuditoriaCamiones C on A.IDCAMION=C.IDCAMION and A.SQAUDITORIA=C.SQAUDITORIA and A.DTCONTABLE=C.DTCONTABLE"
	strSQL = strSQL & ")) as Tabla "
	strSQL = strSQL & myWhere	
	strSQL = strSQL & " ORDER BY DTCONTABLE, IDCAMION, SQCALADA"		
	
	Call GF_BD_Puertos(pto, rs, "OPEN",strSQL)
	set armarCabecera =	rs
End function
'-----------------------------------------------------------------------------------------
function dibujarTitulo(pTitulo)
	Call GF_squareBox(oPDF, 2, 2, 590, 833, 0, "", "#0B3B0B", 2, PDF_SQUARE_ROUND)
	Call GF_writeImage(oPDF, Server.MapPath("images\kogge64.gif"), 20, 10, 48, 48, 0)
	call GF_setFont(oPDF,"ARIAL",16,8)
	Call GF_writeTextAlign(oPDF,2, 25, pTitulo, 590, PDF_ALIGN_CENTER)
	Call GF_horizontalLine(oPDF,2,65,590)
	call GF_setFont(oPDF,"ARIAL",8,0)
	Call GF_writeTextAlign(oPDF, 10 , 840, "Pagina  " & nroPagina		 , 580 , PDF_ALIGN_RIGHT)		
	Call GF_setFont(oPDF,"COURIER",8,0)	
	Call GF_writeTextAlign(oPDF,5,5,GF_FN2DTE(session("MmtoSistema")), 580 , PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,5,15,session("Usuario"), 580 , PDF_ALIGN_RIGHT)
end function
'-----------------------------------------------------------------------------------------
function validarFiltros(py,pto,fechaD,fechaH,pidcamion,pnuCartaPorte,pcodVerificador,pcdProducto,pcdVendedor,pcdCorredor,pcdCliente,pcdEntregador, pcdEstado)
	Dim auxCamion,auxCartaPorte
	py = 80
	Call GF_writeTextAlign(oPDF,20, py, "Puerto.........: "&pto, 590, PDF_ALIGN_LEFT)
	py = py + SEPARATION
	myFormatFecha = GF_FN2DTE(fechaD)
	Call GF_writeTextAlign(oPDF,20, py, "Fecha Desde....: "& myFormatFecha , 590, PDF_ALIGN_LEFT)
	py = py + SEPARATION 
	myFormatFecha = GF_FN2DTE(fechaH)
	Call GF_writeTextAlign(oPDF,20, py, "Fecha Hasta....: "& myFormatFecha, 590, PDF_ALIGN_LEFT)
	py = py + SEPARATION 
	if(pidcamion > 0)then 
		auxCamion = pidcamion 
	else
		if (nuCartaPorte <> "") then
			auxCamion = GF_EDIT_CTAPTE(GF_nChars(nuCartaPorte, 16, "0", CHR_AFT))
		else
			auxCamion = "Todos" 
		end if
	end if
	Call GF_writeTextAlign(oPDF,20, py, "Camion.........: "& auxCamion, 590, PDF_ALIGN_LEFT)
	py = py + SEPARATION 
	
	auxCartaPorte = "Todas"
	if(pnuCartaPorte&pcodVerificador <> "") then auxCartaPorte = pnuCartaPorte&pcodVerificador 
	Call GF_writeTextAlign(oPDF,20, py, "Carta de Porte.: "& auxCartaPorte, 590, PDF_ALIGN_LEFT)
	py = 80
	
	auxProducto = "Todos"
	if(pcdProducto > 0)then auxProducto = Trim(pcdProducto)&" - "&Trim(getDsProducto(pcdProducto))	
	Call GF_writeTextAlign(oPDF,270, py, "Producto.......: "& auxProducto, 590, PDF_ALIGN_LEFT)
	py = py + SEPARATION 
	auxCorredor = "Todos"
	if(pcdCorredor > 0)then
		auxCorredor = Trim(pcdCorredor)&" - "&Trim(getDsCorredor(pcdCorredor))	
		if(len(auxCorredor) > 44)then auxCorredor = left(auxCorredor,44) & "..." 
	end if	
	Call GF_writeTextAlign(oPDF,270, py, "Corredor.......: "& auxCorredor, 590, PDF_ALIGN_LEFT)
	py = py + SEPARATION 
	
	auxVendedor = "Todos"
	if(pcdVendedor > 0)then 
		auxVendedor = Trim(pcdVendedor)&" - "&Trim(getDsVendedor(pcdVendedor))	
		if(len(auxVendedor) > 44)then auxVendedor = left(auxVendedor,44) & "..." 
	end if
	Call GF_writeTextAlign(oPDF,270, py, "Vendedor.......: "& auxVendedor, 590, PDF_ALIGN_LEFT)
	py = py + SEPARATION 
	
	auxCliente = "Todos"
	if(pcdCliente > 0)then 
		auxCliente = Trim(pcdCliente)&" - "&Trim(getDsCliente(pcdCliente))	
		if(len(auxCliente) > 44)then auxCliente = left(auxCliente,44) & "..." 
	end if
	Call GF_writeTextAlign(oPDF,270, py, "Cliente........: "& auxCliente, 590, PDF_ALIGN_LEFT)
	py = py + SEPARATION 
	
	auxEntregador = "Todos"
	if(pcdEntregador > 0)then 
		auxEntregador = Trim(pcdEntregador)&" - "&Trim(getDsEntregador(pcdEntregador))	
		if(len(auxEntregador) > 44)then auxEntregador = left(auxEntregador,44) & "..." 
	end if
	Call GF_writeTextAlign(oPDF,270, py, "Entregador.....: "& auxEntregador, 590, PDF_ALIGN_LEFT)
	py = py + SEPARATION 
	if(pcdEstado > 0)then 
		auxEstado = Trim(pcdEstado)&" - "&Trim(getDsEstado(pcdEstado))	
		if(len(auxEstado) > 30)then auxEntregador = left(auxEstado,30)
    else
        auxEstado = "Descargados OK"
	end if
	Call GF_writeTextAlign(oPDF,270, py, "Estado.........: "& auxEstado , 590, PDF_ALIGN_LEFT)
	py = py + SEPARATION
	
	validarFiltros = py
end function
'-----------------------------------------------------------------------------------------
Function armarPDFGeneral(idcamion, pto, fechaD, fechaH, nuCartaPorte, codVerificador,pcdProducto,pcdVendedor,pcdCorredor,pcdCliente,pcdEntregador, pcdEstado)
Dim totalReg, i, cambioPagina,rsCab,myFecha, cont, myCamion,rs, myFormatFechaVieja,myFormatFechaNueva
cont = 0
cambioPagina = false
Call dibujarTitulo("DATOS VISTEOS CALADA")	
currentY = validarFiltros(currentY,pto,fechaD,fechaH,idcamion,nuCartaPorte,codVerificador,pcdProducto,pcdVendedor,pcdCorredor,pcdCliente,pcdEntregador, pcdEstado)
Set rs = armarTitulos(idCamion, pto, nuCartaPorte, codVerificador, fechaD, fechaH,pcdProducto,pcdVendedor,pcdCorredor,pcdCliente,pcdEntregador, pcdEstado)
while not rs.eof	
	if currentY > PAGE_HEIGHT_SIZE  then	
		Call nuevaPagina()					 
		cambioPagina = true
		currentY = 80
	end if
	Call dibujarTitulosTitularPrimario(currentY)
	currentY = currentY + SEPARATION + 3
	currentY = getTitulosDescripcionPrimaria(currentY,rs("IDCAMION"),rs("CDPRODUCTO"),rs("NUCARTAPORTE"),rs("CDCHAPACAMION"),rs("CDCHAPAACOPLADO"),rs("CDCLIENTE"),rs("DTCONTABLE"))
	currentY = currentY + SEPARATION 
	if currentY > PAGE_HEIGHT_SIZE  then	
		Call nuevaPagina()					 
		cambioPagina = true
		currentY = 80
	end if	
	Call dibujarTitulosTitularSecundario(currentY)
	currentY = currentY + SEPARATION + 3
	currentY = getTitulosDescripcionSecundaria(currentY,rs("CDCORREDOR"),rs("DSCORREDOR"),rs("CDENTREGADOR"),rs("DSENTREGADOR"),rs("CDVENDEDOR"),rs("DSVENDEDOR"))
	currentY = currentY + SEPARATION 
	'LE DOY EL FORMATO AAAA-MM-DD PARA PODER TRABAJAR CON LA BASE DA DATOS	
	'myFormatFechaVieja = split(rs("DTCONTABLE"),"/")	
	'myFormatFechaNueva = GF_nDigits(myFormatFechaVieja(2), 4)&"-"&GF_nDigits(myFormatFechaVieja(0), 2)&"-"&GF_nDigits(myFormatFechaVieja(1), 2)	
	set rsCab = armarCabecera(rs("IDCAMION"), pto, rs("DTCONTABLE"))		
	totalReg1 = rsCab.RecordCount
	currentAuxY = currentY
	while not rsCab.eof		
		 if currentY > PAGE_HEIGHT_SIZE  then			 	 			
		 	 Call nuevaPagina()					 
		 	 cambioPagina = true
		 	 currentY = 80
		 end if
		 call dibujarTitulosDetalle(currentY)	
		 currentY = currentY + SEPARATION + 4			 
		 'myFormatFechaVieja = split(rsCab("DTCONTABLE"),"/")
		 'myfecha = GF_nDigits(myFormatFechaVieja(2), 4)&"-"&GF_nDigits(myFormatFechaVieja(0), 2)&"-"&GF_nDigits(myFormatFechaVieja(1), 2)			 
		 myCamion = rs("IDCAMION")		 
		 currentY = dibujarDatoCalada(currentY,rsCab("DTCONTABLE"),rsCab("SQCALADA"),rsCab("CDUSERNAME"),rsCab("CDTERMINAL"),rsCab("CDSUPERVISOR"),pto)
	 	 cambioPagina = false
		 Set rsDet = armarDetalle(rsCab("SQCALADA"), myCamion,rsCab("DTCONTABLE"),pto) 
		 currentY = currentY + (SEPARATION)		 
		 Call dibujarTitularBoxDetalle(currentY)
		 currentY = currentY + SEPARATION + 1
		 while not rsDet.eof 
			if currentY > PAGE_HEIGHT_SIZE then					
				Call nuevaPagina()					 
				cambioPagina = true
				currentY = 80
			end if						
			currentY = dibujarDetalleVale(currentY, rsDet("cdrubro"), rsDet("vlbonrebaja"), rsDet("vlmerma"),pto)
			rsDet.moveNext								
		wend	
		rsCab.moveNext		
	wend
	rs.moveNext
wend
currentY = currentY + SEPARATION + 1
Call GF_horizontalLine(oPDF,20,currentY + SEPARATION/2 ,558)
Call GF_writeTextAlign(oPDF, 20, currentY + SEPARATION/2 + 10, GF_TRADUCIR("FIN DEL REPORTE"), 550, PDF_ALIGN_CENTER)		
end function
'-----------------------------------------------------------------------------------------
function nuevaPagina()
	Call GF_newPage(oPDF)
	currentY = PAGE_TOP_INIT + (SEPARATION) + 4		
	nroPagina = nroPagina + 1
	call dibujarTitulo("DATOS CALADA")
	currentAuxY = PAGE_TOP_INIT 
end function
'-----------------------------------------------------------------------------------------
Function getTitulosDescripcionPrimaria(pY,pIdCamion,pCodProducto,pNPorte,pChapa,pAcoplado,pCodCli,pdtcont)	
	Call GF_squareBox(oPDF, 18, pY, 562, 10, 0, color, color, 1, PDF_SQUARE_NORMAL)
	Call GF_writeTextAlign(oPDF, 18, pY+1 	, GF_FN2DTE(pdtcont)	, 47	, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, 65, pY+1	, pIdCamion, 55	, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, 120, pY+1	, GF_EDIT_CTAPTE(pNPorte&pNPorte3), 80 , PDF_ALIGN_CENTER)
	auxPRO = Trim(pCodProducto)&" - "&Trim(getDsProducto(pCodProducto))	
	if (len(auxPRO) > 20) then auxPRO = left(auxPRO,20) & "..."
	Call GF_writeTextAlign(oPDF, 200, pY+1	, auxPRO, 100 , PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF, 300, pY+1	, pChapa, 45 , PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, 345, pY+1	, pAcoplado, 55 , PDF_ALIGN_CENTER)	
	auxCli = Trim(pCodCli)&" - "&Trim(getDsCliente(pCodCli))	
	if (len(auxCli) > 39) then auxCli = left(auxCli,39) & "..."
	Call GF_writeTextAlign(oPDF, 400, pY+1	, auxCli, 180 , PDF_ALIGN_LEFT)			
	getTitulosDescripcionPrimaria = pY
End function
'-----------------------------------------------------------------------------------------
Function getTitulosDescripcionSecundaria(pY,pCdCorredor,pDsCorredor,pCdEntregador,pDsEntregador,pCdVendedor,pDsVendedor)	
	Call GF_squareBox(oPDF, 18, pY, 562, 10, 0, color, color, 1, PDF_SQUARE_NORMAL)
	auxCorredor = Trim(pCdCorredor)&" - "&Trim(pDsCorredor)
	if(len(auxCorredor) > 40)then auxCorredor = left(auxCorredor,41) & "..."
	Call GF_writeTextAlign(oPDF, 20, pY+1 	, auxCorredor , 182	, PDF_ALIGN_LEFT)
	auxVendedor = Trim(pCdVendedor)&" - "&Trim(pDsVendedor)	
	if(len(auxVendedor) > 40)then auxVendedor = left(auxVendedor,41) & "..."
	Call GF_writeTextAlign(oPDF, 202, pY+1	, auxVendedor , 186	, PDF_ALIGN_LEFT)
	auxEntregador = Trim(pCdEntregador)&" - "&Trim(pDsEntregador)
	if(len(auxEntregador) > 40)then auxEntregador = left(auxEntregador,41) & "..."
	Call GF_writeTextAlign(oPDF, 388, pY+1	, auxEntregador, 190 , PDF_ALIGN_LEFT)
	getTitulosDescripcionSecundaria = pY		
End function
'-----------------------------------------------------------------------------------------
Function dibujarDatoCalada(currentY,pdtcontable,psqcalada,pcdusername,pcdterminal, cdSupervisor, pto)
	Dim auxValeRelacionado, dsPartSector, proveedor, dsSupervisor
	
	dsSupervisor = ""
	if (cdSupervisor <> "") then dsSupervisor = getNombreApellidoCalada(cdSupervisor)
	
	call GF_setFont(oPDF,"COURIER",8,0)
	color = "#CCFFCC"
	Call GF_squareBox(oPDF, 60, currentY, 465, 10, 0, color, color, 1, PDF_SQUARE_NORMAL)
	Call GF_writeTextAlign(oPDF, 60, currentY, GF_FN2DTE(pdtcontable)	, 55	, PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(oPDF,115, currentY, psqcalada								, 20	, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,135, currentY, getNombreApellidoCalada(pcdusername)     , 155	, PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(oPDF,290, currentY, pcdterminal								, 80	, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,370, currentY, dsSupervisor								, 155	, PDF_ALIGN_CENTER)		
	dibujarDatoCalada = currentY		
end function
'-----------------------------------------------------------------------------------------
Function dibujarDetalleVale(currentY, pcdrubro,pvlbonrebaja,pvlmerma,pto)	
	call GF_setFont(oPDF,"COURIER",6,0)	
	Call GF_squareBox(oPDF, 170, currentY, 280, 10, 0, "#F5F6CE", "#F5F6CE", 1, PDF_SQUARE_NORMAL)
	Call GF_writeTextAlign(oPDF,172, currentY,pcdrubro, 60, PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(oPDF,232, currentY,getDsRubro(pcdrubro) , 110, PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(oPDF,342, currentY, pvlbonrebaja, 60, PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(oPDF,402, currentY,  pvlmerma, 50, PDF_ALIGN_CENTER)	
	currentY = currentY + SEPARATION		
	dibujarDetalleVale = currentY
end Function
'-----------------------------------------------------------------------------------------
function dibujarTitulosDetalle(pY)
	Call GF_squareBox(oPDF, 60	,pY , 55	, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 115	,pY , 20	, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)	
	Call GF_squareBox(oPDF, 135	,pY , 155	, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 290	,pY , 80	, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 370	,pY , 155	, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	call GF_setFont(oPDF,"ARIAL",8,8)
	call GF_setFontColor("FFFFFF")
	Call GF_writeTextAlign(oPDF, 62,  pY+2 	, GF_TRADUCIR("FECHA")		, 55	, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, 115, pY+2	, GF_TRADUCIR("SQ")			, 20	, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, 137, pY+2	, GF_TRADUCIR("USUARIO")	, 155	, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, 294, pY+2	, GF_TRADUCIR("PC TERMINAL"), 80	, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, 374, pY+2	, GF_TRADUCIR("SUPERVISOR")	, 155	, PDF_ALIGN_CENTER)		
	call GF_setFontColor("000000")
end function
'------------------------------------------------------------------------------------------
function dibujarTitulosTitularPrimario(pY)	
	Call GF_squareBox(oPDF, 18	,pY , 47	, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 65	,pY , 55	, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 120	,pY , 80	, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 200	,pY , 100	, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 300	,pY , 45	, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 345	,pY , 55	, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 400	,pY , 180	, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	call GF_setFont(oPDF,"ARIAL",8,8)
	call GF_setFontColor("FFFFFF")
	Call GF_writeTextAlign(oPDF, 18, pY+1 	, GF_TRADUCIR("FECHA"), 45	, PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(oPDF, 65, pY+1	, GF_TRADUCIR("CAMION"), 55	, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, 120, pY+1	, GF_TRADUCIR("CARTA PORTE"), 80 , PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, 200, pY+1	, GF_TRADUCIR("PRODUCTO"), 100 , PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, 300, pY+1	, GF_TRADUCIR("CHASIS"), 45 , PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, 345, pY+1	, GF_TRADUCIR("ACOPLADO"), 55 , PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(oPDF, 400, pY+1	, GF_TRADUCIR("CLIENTE"), 180 , PDF_ALIGN_CENTER)	
		
	call GF_setFontColor("000000")
end function
'------------------------------------------------------------------------------------------
function dibujarTitulosTitularSecundario(pY)	
	Call GF_squareBox(oPDF, 18	,pY , 182	, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 200	,pY , 186	, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 386	,pY , 194	, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	call GF_setFont(oPDF,"ARIAL",8,8)
	call GF_setFontColor("FFFFFF")
	Call GF_writeTextAlign(oPDF, 18, pY+1 	, GF_TRADUCIR("CORREDOR"), 182	, PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(oPDF, 200, pY+1	, GF_TRADUCIR("VENDEDOR"), 186	, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, 386, pY+1	, GF_TRADUCIR("ENTREGADOR"), 194 , PDF_ALIGN_CENTER)			
	call GF_setFontColor("000000")
end function
'------------------------------------------------------------------------------------------
Function dibujarTitularBoxDetalle(pY)
	Call GF_squareBox(oPDF, 170, pY, 60, 10, 0, "#F2F5A9", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 230, pY, 110, 10, 0, "#F2F5A9", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 340, pY, 60, 10, 0, "#F2F5A9", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 400, pY, 50, 10, 0, "#F2F5A9", "#000000", 1, PDF_SQUARE_NORMAL)
	call GF_setFont(oPDF,"ARIAL",6,8)
	call GF_setFontColor("000000")
	Call GF_writeTextAlign(oPDF, 172, pY+1	, GF_TRADUCIR("COD RUBRO"), 60	, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, 232, pY+1	, GF_TRADUCIR("RUBRO"), 110	, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, 342, pY+1	, GF_TRADUCIR("BONREBAJA"), 60	, PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(oPDF, 402, pY+1	, GF_TRADUCIR("MERMA")	, 50	, PDF_ALIGN_CENTER)			
	call GF_setFontColor("000000")
end Function
'/********************************************************************\
' **********                 COMIENZA PAGINA                 *********
'\********************************************************************/
Dim pto

SEPARATION = 10
MARGIN = 0
PAGE_HEIGHT_SIZE = 800
PAGE_TOP_INIT = 82
nroPagina = 1

pto = GF_PARAMETROS7("pto", "", 6)
idcamion = GF_PARAMETROS7("idcamion", 0, 6)
if (idcamion <> "") then idcamion = GF_nDigits(idcamion, 10)

fecContableD = GF_PARAMETROS7("fecContableD", "", 6)
fecContableM = GF_PARAMETROS7("fecContableM", "", 6)
fecContableA = GF_PARAMETROS7("fecContableA", "", 6)

fecContableDH = GF_PARAMETROS7("fecContableDH", "", 6)
fecContableMH = GF_PARAMETROS7("fecContableMH", "", 6)
fecContableAH = GF_PARAMETROS7("fecContableAH", "", 6)

nuCartaPorte1 = GF_PARAMETROS7("nuCartaPorte1", "", 6)
if (nuCartaPorte1 <> "") then nuCartaPorte1 = GF_nDigits(nuCartaPorte1, 4)
nuCartaPorte2 = GF_PARAMETROS7("nuCartaPorte2", "", 6)
if (nuCartaPorte2 <> "") then nuCartaPorte2 = GF_nDigits(nuCartaPorte2, 8)
nuCartaPorte3 = GF_PARAMETROS7("nuCartaPorte3", "", 6)
if (nuCartaPorte3 <> "") then nuCartaPorte3 = GF_nDigits(nuCartaPorte3, 4)

'------------------------Nuevos Filtros----------------------
cdProducto = GF_PARAMETROS7("cdProducto", 0, 6)

cdVendedor = GF_PARAMETROS7("cdVendedor", "", 6)
dsVendedor = GF_PARAMETROS7("dsVendedor", "", 6)

cdCorredor = GF_PARAMETROS7("cdCorredor", "", 6)
dsCorredor = GF_PARAMETROS7("dsCorredor", "", 6)

cdCliente = GF_PARAMETROS7("cdCliente", "", 6)
dsCliente = GF_PARAMETROS7("dsCliente", "", 6)

cdEntregador = GF_PARAMETROS7("cdEntregador", "", 6)
dsEntregador = GF_PARAMETROS7("dsEntregador", "", 6)

cdEstado = GF_PARAMETROS7("cdEstado", "", 6)
'---------------------------------------------------------
g_strPuerto = pto

nuCartaPorte = nuCartaPorte1 & nuCartaPorte2

Call GF_STANDARIZAR_FECHA(fecContableD, fecContableM, fecContableA)
Call GF_STANDARIZAR_FECHA(fecContableDH, fecContableMH, fecContableAH)
'db2FechaD = fecContableA & "-" & fecContableM & "-" & fecContableD
'db2FechaH = fecContableAH & "-" & fecContableMH & "-" & fecContableDH
db2FechaD = fecContableA & fecContableM & fecContableD
db2FechaH = fecContableAH & fecContableMH & fecContableDH

Set oPDF = GF_createPDF("")
Call GF_setPDFMODE(PDF_STREAM_MODE)
Call armarPDFGeneral(idcamion, pto, db2FechaD, db2FechaH, nuCartaPorte, nuCartaPorte3,cdProducto,cdVendedor,cdCorredor,cdCliente,cdEntregador, cdEstado)
Call GF_closePDF(oPDF)	
%>
