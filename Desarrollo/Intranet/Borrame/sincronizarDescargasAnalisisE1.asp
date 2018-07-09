<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientos.asp"-->
<!--#include file="Includes/procedimientosPuertos.asp"-->
<!--#include file="Includes/procedimientosformato.asp"-->
<!--#include file="Includes/procedimientosLog.asp"-->
<!--#include file="Includes/procedimientosTraducir.asp"-->
<%'--------------------------------------------------------------------------------------------------------------
Function migrarAnalisis(pMmto,pPto, p_transporte)
    Dim strSQL, rs, auxCdBolsa
    
    auxCdBolsa	= getCodigoBolsa(pPto)
    
    'Borro los datos del puerto    
	strSQL = "DELETE FROM MERFL.MER582F1 WHERE COBEC1 = "& auxCdBolsa & " and FANAC1=" & pMmto
    Call executeQuery(rs, "EXEC", strSQL)    
    strSQL = "DELETE FROM MERFL.MER582F2 WHERE COBEC2= " & auxCdBolsa & " and FANAC2=" & pMmto
    Call executeQuery(rs, "EXEC", strSQL)
    strSQL = "DELETE FROM MERFL.MER583F1 WHERE COBEC1 = "& auxCdBolsa & " and FANAC1=" & pMmto
    Call executeQuery(rs, "EXEC", strSQL)    
    strSQL = "DELETE FROM MERFL.MER583F2 WHERE COBEC2= " & auxCdBolsa & " and FANAC2=" & pMmto
    Call executeQuery(rs, "EXEC", strSQL)
    
    if ((p_transporte = TIPO_TRANSPORTE_CAMION) or (p_transporte = TIPO_TRANSPORTE_CAMVAG)) then Call migrarAnalisisCamiones(pMmto,pPto)		
	if ((p_transporte = TIPO_TRANSPORTE_VAGON) or (p_transporte = TIPO_TRANSPORTE_CAMVAG)) then Call migrarAnalisisVagones(pMmto,pPto)
		
End Function
'-----------------------------------------------------------------------------------------------------------------
Function migrarAnalisisCamiones(pFecha,pPto)
	Dim auxIdCamionOld,auxCtaPteOld,auxNroPto,rs,auxCdBolsa, fechaCamion
	auxNroPto	= getNumeroPuerto(pPto)
	auxCdBolsa	= getCodigoBolsa(pPto)

    'Se migran los nuevos datos.
	Set rs = armarSQLAnalisisCamiones(pFecha,pPto)	
	while (not rs.Eof)
	    'Debo verificar si ya tiene an{alisis.	    
	    auxIdCamionOld  = CDbl(rs("CAMION"))			
	    auxCtaPteOld	= CDbl(rs("CARTAPORTE"))
	    auxProductoOld	= Cdbl(rs("CDPRODUCTO"))		
	    fechaCamion = GF_DTE2FN(GF_STANDARIZAR_FECHA_RTRN(rs("FECHA")))
	    'if (not existeAnalisisCargado(fechaCamion, rs("CARTAPORTE"))) then
		    Call grabarCabeceraAnalisisCamiones(pPto,auxCdBolsa,auxNroPto,rs("CDPRODUCTO"),rs("CAMION"), Right(Trim(rs("NUINFOANALISIS")), 6),fechaCamion,rs("CARTAPORTE"),rs("KILOSNETOS"), rs("ICCAMARA"))
		    logMig.Info("NRO.PUERTO : "& auxNroPto&" | NRO.ANALISIS: "& rs("CAMION") &" | PRODUCTO: "&rs("CDPRODUCTO")&" | CARTA PORTE: "& rs("CARTAPORTE") & " | KILOS NETOS: " & GF_EDIT_DECIMALS(cdbl(rs("KILOSNETOS")),0) & " Kg.")		
		    while (controlarCamionProducto(auxIdCamionOld, auxCtaPteOld, auxProductoOld, rs))		
			    if (Cdbl(rs("RUBRO")) <> 0) then
				    Call grabarDetalleAnalisisCamiones(pPto,auxCdBolsa,auxNroPto,rs("CDPRODUCTO"),rs("CAMION"),fechaCamion,rs("RUBRO"),rs("BONREBAJA"), rs("ICCAMARA"))
				    logMig.Info("--> RUBRO : "& rs("RUBRO") &" | VLBONREBAJA: "&rs("BONREBAJA"))
			    end if	
			    rs.MoveNext()
		    wend		
		'else
		'    logMig.Info("ERROR!!! - YA EXISTE ANALISIS PARA LA CARTA DE PORTE. NRO.PUERTO : "& auxNroPto&" | NRO.ANALISIS: "& rs("CAMION") &" | PRODUCTO: "&rs("CDPRODUCTO")&" | CARTA PORTE: "& rs("CARTAPORTE") & " | KILOS NETOS: " & GF_EDIT_DECIMALS(cdbl(rs("KILOSNETOS")),0) & " Kg.")		
		'    while (controlarCamionProducto(auxIdCamionOld, auxCtaPteOld, auxProductoOld, rs))		
		'        rs.MoveNext()
		'    wend
        'end if		    
	wend
End Function
'-----------------------------------------------------------------------------------------------------------------
Function existeAnalisisCargado(pFecha, pCartaPorte)
    Dim rs, strSQL, ret
    
    strSQL="Select * from MERFL.MER591CA where FANACA=" & pFecha & " and CPORCA=" & CDbl(pCartaPorte)
    Call executeQuery(rs, "OPEN", strSQL)
    ret = false
    if (not rs.eof) then ret= true
    existeAnalisisCargado = ret
    
End Function
'-----------------------------------------------------------------------------------------------------------------
Function armarSQLAnalisisCamiones(p_mmto,pPto)
	Dim strSQL,myDtContable
	myDtContable = Left(p_mmto, 4) & "-" & mid(p_mmto, 5, 2) & "-" & Right(p_mmto, 2)
	strSQL = "SELECT HC.CDPRODUCTO, "&_
			 "		 CASE WHEN HC.idcamion IS NOT NULL THEN HC.idcamion ELSE '' END AS CAMION, "&_			 
			 "		 HC.DTCONTABLE AS FECHA, "&_
			 "		 CASE WHEN HCD.NUCARTAPORTE IS NOT NULL THEN HCD.NUCARTAPORTE ELSE '' END AS CARTAPORTE, "&_
			 "		((SELECT PC.VLPESADA FROM HPESADASCAMION PC WHERE PC.DTCONTABLE = HCD.DTCONTABLE AND PC.IDCAMION = HCD.IDCAMION AND PC.CDPESADA = 1 AND PC.SQPESADA = (SELECT MAX(SQPESADA) FROM HPESADASCAMION WHERE PC.DTCONTABLE = DTCONTABLE AND PC.IDCAMION = IDCAMION AND CDPESADA = 1)) "&_
			 "		  -   "&_
			 "		 (SELECT PC.VLPESADA FROM HPESADASCAMION PC WHERE PC.DTCONTABLE = HCD.DTCONTABLE AND PC.IDCAMION = HCD.IDCAMION AND PC.CDPESADA = 2 AND PC.SQPESADA = (SELECT MAX(SQPESADA) FROM HPESADASCAMION WHERE PC.DTCONTABLE = DTCONTABLE AND PC.IDCAMION = IDCAMION AND CDPESADA = 2)) "&_
			 "		  - "&_
			 "		 (SELECT CASE WHEN HMC.VLMERMAKILOS IS NULL THEN 0 ELSE HMC.VLMERMAKILOS END FROM HMERMASCAMIONES HMC WHERE HMC.DTCONTABLE=HCD.DTCONTABLE AND HMC.IDCAMION = HCD.IDCAMION AND HMC.SQPESADA= (SELECT MAX(SQPESADA) FROM HMERMASCAMIONES WHERE DTCONTABLE=HCD.DTCONTABLE AND IDCAMION = HCD.IDCAMION)) "&_
			 "		) AS KILOSNETOS, "&_
			 "		 HCC.ICCAMARA, "&_
			 "		 R.CDRUBRO AS RUBRO, "&_	
			 "       HCD.NUINFOANALISIS, " &_
			 "		 HRVC.VLBONREBAJA AS BONREBAJA "&_			 
			 "FROM ( SELECT * "&_ 
			 "		 FROM HCAMIONES  "&_
			 "		 WHERE DTCONTABLE = '"& myDtContable &"' AND CDESTADO IN(6,8)) AS HC "&_
			 "INNER JOIN HCAMIONESDESCARGA HCD ON HC.IDCAMION = HCD.IDCAMION AND HC.DTCONTABLE = HCD.DTCONTABLE "&_
			 "INNER JOIN (Select CDRUBRO,DTCONTABLE,IDCAMION,SQCALADA,VLBONREBAJA "&_
			 "			   FROM (Select * from hrubrosvisteocamiones where DTCONTABLE='" & myDtContable & "') A "&_
			 "			   WHERE A.SQCALADA = (SELECT MAX(SQCALADA) "&_
             "								   FROM   hrubrosvisteocamiones "&_
             "								   WHERE  idcamion = A.idcamion "&_
             "										AND dtcontable = A.dtcontable) "&_
			 "		    ) HRVC ON HRVC.DTCONTABLE = HC.DTCONTABLE AND HRVC.IDCAMION = HC.IDCAMION "&_
			 "INNER JOIN HCALADADECAMIONES HCC on HRVC.DTCONTABLE=HCC.DTCONTABLE and HRVC.IDCAMION=HCC.IDCAMION and HRVC.SQCALADA=HCC.SQCALADA "&_
			 "INNER JOIN RUBROS R ON R.CDRUBRO = HRVC.CDRUBRO "&_
			 "INNER JOIN PRODUCTOS P ON P.CDPRODUCTO = HC.CDPRODUCTO "&_
			 "INNER JOIN CLIENTES C ON C.CDCLIENTE = HCD.CDCLIENTE " &_
	         " ORDER BY HC.idcamion,HC.DTCONTABLE "  
	'response.Write strSQL & "<br>"  
	'Response.End 
	Call GF_BD_Puertos(pPto, rs, "OPEN",strSQL)	 
	Set armarSQLAnalisisCamiones = rs
End Function
'-----------------------------------------------------------------------------------------------------------------
Function grabarCabeceraAnalisisCamiones(pPto,pCodigoBolsa,pNroPto,pCdProducto,pNroAnalisis, pNroSolicitudAnalisis,pFecha,pCtaPte,pKilosNetos, pICCamara)
	Dim strSQL 
	strSQL = "INSERT INTO MERFL.MER583F1(COBEC1,CDCAC1,CDESC1,CPROC1,NROAC1,FANAC1,CPORC1,NSANC1,OBSEC1,GRADC1,GRASC1,IMPAC1,KGMOC1,NPSTC1,PTEMC1,USERC1,FECHC1,HORAC1,ESTAC1) VALUES "&_
			 " ("&pCodigoBolsa&","& pNroPto &","& pNroPto &","& pCdProducto &","& pNroAnalisis &","& pFecha &","& pCtaPte &","& pNroSolicitudAnalisis &",'',0,0,0,"& pKilosNetos &",0,0,'"& session("Usuario") &"',"& Left(session("MmtoDato"),8) &","& Right(session("MmtoDato"),6) &",'') "
	Call executeQuery(rs, "EXEC", strSQL)	
	if (pICCamara = "N") then
	    strSQL = "INSERT INTO MERFL.MER582F1(COBEC1,CDCAC1,CDESC1,CPROC1,NROAC1,FANAC1,CPORC1,NSANC1,OBSEC1,GRADC1,GRASC1,IMPAC1,KGMOC1,NPSTC1,PTEMC1,USERC1,FECHC1,HORAC1,ESTAC1) VALUES "&_
			     " ("&pCodigoBolsa&","& pNroPto &","& pNroPto &","& pCdProducto &","& pNroAnalisis &","& pFecha &","& pCtaPte &","& pNroSolicitudAnalisis &",'',0,0,0,"& pKilosNetos &",0,0,'"& session("Usuario") &"',"& Left(session("MmtoDato"),8) &","& Right(session("MmtoDato"),6) &",'') "
	    Call executeQuery(rs, "EXEC", strSQL)	
    end if	    
End Function
'-----------------------------------------------------------------------------------------------------------------
Function grabarDetalleAnalisisCamiones(pPto,pCodigoBolsa,pNroPto,pCdProducto,pNroAnalisis,pFecha,pCdRubro,pKilosReb, pICCamara)
	Dim strSQL, tipo
	
	'Detemrino el tipo de analisis
	strSQL="Select TIPAAE from MERFL.MER2E9F1 where CAMAAE=" & pCodigoBolsa & " and PRODAE in (0, " & pCdProducto & ") and ANCAAE=" & pCdRubro & " order by PRODAE desc"
	Call executeQuery(rs, "OPEN", strSQL)	
	tipo = 3
	if (not rs.eof) then tipo = rs("TIPAAE")
	pKilosReb = Replace(pKilosReb, ",", ".")
	strSQL = "INSERT INTO MERFL.MER583F2(COBEC2,CDESC2,CPROC2,NROAC2,FANAC2,COANC2,VACAC2,PREBC2,PBONC2,ESTAC2,COCAC2,TIPAC2) VALUES "&_
			 " ("&pCodigoBolsa&",0,"& pCdProducto &","& pNroAnalisis &","& pFecha &",0,"& pKilosReb &",0,0,'',"& pCdRubro &"," & tipo & ")"	
	Call executeQuery(rs, "EXEC", strSQL)	
	if (pICCamara = "N") then
	    strSQL = "INSERT INTO MERFL.MER582F2(COBEC2,CDESC2,CPROC2,NROAC2,FANAC2,COANC2,VACAC2,PREBC2,PBONC2,ESTAC2,COCAC2,TIPAC2) VALUES "&_
			 " ("&pCodigoBolsa&",0,"& pCdProducto &","& pNroAnalisis &","& pFecha &",0,"& pKilosReb &",0,0,'',"& pCdRubro &"," & tipo & ")"	
	Call executeQuery(rs, "EXEC", strSQL)	
	end if 
End Function
'-----------------------------------------------------------------------------------------------------------------
Function controlarCamionProducto(pIdCamionOld,pCtaPteOld,pProductoOld, pRs)
	Dim rtrn
	rtrn  = false	
	if (not pRs.eof) then
	    if((pIdCamionOld = CDbl(pRs("CAMION")))and(pCtaPteOld = CDbl(pRs("CARTAPORTE")))and(pProductoOld = CDbl(pRs("CDPRODUCTO"))))then rtrn = true	
	end if
	controlarCamionProducto = rtrn
End Function
'----------------------------------------------------------------------------------------------------------------
Function armarSQLAnalisisVagones(p_mmto,pPto)
	Dim strSQL,myDtContable
	myDtContable = Left(p_mmto, 4) & "-" & mid(p_mmto, 5, 2) & "-" & Right(p_mmto, 2)
	strSQL = "SELECT HC.CDPRODUCTO, "&_
			 "		 HC.CDVAGON, "&_
			 "		 HC.DTCONTABLE, "&_
			 "		 HC.NUCARTAPORTESERIE, "&_
			 "		 HC.NUCARTAPORTE, "&_
			 "		((SELECT PC.VLPESADA FROM HPESADASVAGON PC WHERE PC.DTCONTABLE = HC.DTCONTABLE AND PC.NUCARTAPORTE = HC.NUCARTAPORTE AND PC.CDVAGON = HC.CDVAGON AND PC.CDPESADA = 1 AND PC.SQPESADA = (SELECT MAX(SQPESADA) FROM HPESADASVAGON WHERE PC.DTCONTABLE = DTCONTABLE AND PC.NUCARTAPORTE = NUCARTAPORTE AND PC.CDVAGON = CDVAGON AND CDPESADA = 1)) "&_
			 "		  -   "&_
			 "		 (SELECT PC.VLPESADA FROM HPESADASVAGON PC WHERE PC.DTCONTABLE = HC.DTCONTABLE AND PC.NUCARTAPORTE = HC.NUCARTAPORTE AND PC.CDVAGON = HC.CDVAGON AND PC.CDPESADA = 2 AND PC.SQPESADA = (SELECT MAX(SQPESADA) FROM HPESADASVAGON WHERE PC.DTCONTABLE = DTCONTABLE AND PC.NUCARTAPORTE = NUCARTAPORTE AND PC.CDVAGON = CDVAGON AND CDPESADA = 2)) "&_
			 "		  - "&_
			 "		 (SELECT CASE WHEN HMC.VLMERMAKILOS IS NULL THEN 0 ELSE HMC.VLMERMAKILOS END FROM HMERMASVAGONES HMC WHERE HMC.DTCONTABLE=HC.DTCONTABLE AND HMC.NUCARTAPORTE = HC.NUCARTAPORTE AND HMC.CDVAGON = HC.CDVAGON AND HMC.SQPESADA= (SELECT MAX(SQPESADA) FROM HMERMASVAGONES WHERE DTCONTABLE=HC.DTCONTABLE AND NUCARTAPORTE = HC.NUCARTAPORTE AND CDVAGON = HC.CDVAGON )) "&_
			 "		) AS KILOSNETOS, "&_
			 "       ICCAMARA, " &_
			 "		 R.CDRUBRO AS RUBRO, "&_
			 "		 HRVC.VLBONREBAJA AS BONREBAJA "&_
			 "FROM ( SELECT * "&_ 
			 "		 FROM HVAGONES  "&_
			 "		 WHERE DTCONTABLE = '"& myDtContable &"' AND CDESTADO IN(6, 8)) AS HC "&_
			 "INNER JOIN HOPERATIVOS HO ON HO.DTCONTABLE = HC.DTCONTABLE AND HO.NUCARTAPORTE = HC.NUCARTAPORTE "&_
			 "INNER JOIN (Select CDRUBRO,DTCONTABLE,NUCARTAPORTE,CDVAGON,SQCALADA,VLBONREBAJA "&_
			 "			   FROM (Select * from HRUBROSVISTEOVAGONES where DTCONTABLE='" & myDtContable & "') A "&_
			 "			   WHERE A.SQCALADA = (SELECT MAX(SQCALADA) "&_
             "								   FROM   HRUBROSVISTEOVAGONES "&_
             "								   WHERE  NUCARTAPORTE = A.NUCARTAPORTE "&_
             "										  AND CDVAGON = A.CDVAGON "&_
             "										  AND dtcontable = A.dtcontable) "&_
			 "		    ) HRVC ON HRVC.DTCONTABLE = HC.DTCONTABLE AND HRVC.NUCARTAPORTE = HC.NUCARTAPORTE AND HRVC.CDVAGON = HC.CDVAGON "&_
			 "INNER JOIN HCALADADEVAGONES HCC on HRVC.DTCONTABLE=HCC.DTCONTABLE AND HRVC.NUCARTAPORTE = HCC.NUCARTAPORTE and HRVC.CDVAGON=HCC.CDVAGON and HRVC.SQCALADA=HCC.SQCALADA "&_
			 "INNER JOIN RUBROS R ON R.CDRUBRO = HRVC.CDRUBRO "&_
			 "INNER JOIN PRODUCTOS P ON P.CDPRODUCTO = HC.CDPRODUCTO "&_
			 "INNER JOIN CLIENTES C ON C.CDCLIENTE = HO.CDCLIENTE "&_
	         " ORDER BY HC.NUCARTAPORTE,HC.CDVAGON,R.CDRUBRO "
	Call GF_BD_Puertos(pPto, rs, "OPEN",strSQL)
	'Response.Write strSQL
	'Response.End
	Set armarSQLAnalisisVagones = rs
End Function
'---------------------------------------------------------------------------------------------------------------
Function migrarAnalisisVagones(pFecha,pPto)
	Dim auxIdCamion,auxFecha,auxCtaPte,auxNroPto,rs,auxCdBolsa
	auxNroPto	= getNumeroPuerto(pPto)    
	Set rs = armarSQLAnalisisVagones(pFecha,pPto)
	while not rs.Eof
		auxCtaPteOld	= CDbl(rs("NUCARTAPORTE"))
		auxCdVagon		= CDbl(rs("CDVAGON"))
		auxCdBolsa		= getCodigoBolsa(pPto)
		fechaVagon = GF_DTE2FN(GF_STANDARIZAR_FECHA_RTRN(rs("DTCONTABLE")))
		myCPorte = rs("NUCARTAPORTESERIE") & Left(rs("NUCARTAPORTE"), 8)
		Call grabarCabeceraAnalisisVagones(pPto,auxCdBolsa,auxNroPto,rs("CDPRODUCTO"),auxCdVagon, fechaVagon,myCPorte,rs("KILOSNETOS"), rs("ICCAMARA"))
		logMig.Info("NRO.PUERTO : "& auxNroPto&" | NRO.ANALISIS: "&auxCdVagon&" | PRODUCTO: "&rs("CDPRODUCTO")&" | CARTA PORTE: "& myCPorte & " | KILOS NETOS: " & GF_EDIT_DECIMALS(cdbl(rs("KILOSNETOS")),0) & " Kg.")
		while (controlarCtaPteVagon(auxCdVagon,auxCtaPteOld,rs))
			Call grabarDetalleAnalisisVagones(pPto,auxCdBolsa,auxNroPto,rs("CDPRODUCTO"),auxCdVagon,fechaVagon,rs("RUBRO"),rs("BONREBAJA"), rs("ICCAMARA"))
			logMig.Info("--> RUBRO : "& rs("RUBRO") &" | VLBONREBAJA: "&rs("BONREBAJA"))
			rs.MoveNext()
		wend
	wend	
End Function
'-----------------------------------------------------------------------------------------------------------------
Function grabarCabeceraAnalisisVagones(pPto,pCodigoBolsa,pNroPto,pCdProducto,pNroAnalisis,pFecha,pCtaPte,pKilosNetos, pICCamara)
	Dim strSQL
	strSQL = "INSERT INTO MERFL.MER583F1 (COBEC1,CDCAC1,CDESC1,CPROC1,NROAC1,FANAC1,CPORC1,NSANC1,OBSEC1,GRADC1,GRASC1,IMPAC1,KGMOC1,NPSTC1,PTEMC1,USERC1,FECHC1,HORAC1,ESTAC1) VALUES "&_
			 " ("&pCodigoBolsa&","& pNroPto &","& pNroPto &","& pCdProducto &","& pNroAnalisis &","& pFecha &","& pCtaPte &","& pNroAnalisis &",'',0,0,0,"& pKilosNetos &",0,0,'"& session("Usuario") &"',"& Left(session("MmtoDato"),8) &","& Right(session("MmtoDato"),6) &",'') "	
	Call executeQuery(rs, "EXEC", strSQL)	
	if (pICCamara = "N") then
	    strSQL = "INSERT INTO MERFL.MER582F1 (COBEC1,CDCAC1,CDESC1,CPROC1,NROAC1,FANAC1,CPORC1,NSANC1,OBSEC1,GRADC1,GRASC1,IMPAC1,KGMOC1,NPSTC1,PTEMC1,USERC1,FECHC1,HORAC1,ESTAC1) VALUES "&_
			     " ("&pCodigoBolsa&","& pNroPto &","& pNroPto &","& pCdProducto &","& pNroAnalisis &","& pFecha &","& pCtaPte &","& pNroAnalisis &",'',0,0,0,"& pKilosNetos &",0,0,'"& session("Usuario") &"',"& Left(session("MmtoDato"),8) &","& Right(session("MmtoDato"),6) &",'') "	
	    Call executeQuery(rs, "EXEC", strSQL)	
    end if	    
End Function
'-----------------------------------------------------------------------------------------------------------------
Function grabarDetalleAnalisisVagones(pPto,pCodigoBolsa,pNroPto,pCdProducto,pNroAnalisis,pFecha,pCdRubro,pKilosReb, pICCamara)
	Dim strSQL
	'Detemrino el tipo de analisis
	strSQL="Select TIPAAE from MERFL.MER2E9F1 where CAMAAE=" & pCodigoBolsa & " and PRODAE in (0, " & pCdProducto & ") and ANCAAE=" & pCdRubro & " order by PRODAE desc"
	Call executeQuery(rs, "OPEN", strSQL)	
	tipo = 3
	if (not rs.eof) then tipo = rs("TIPAAE")
	pKilosReb = Replace(pKilosReb, ",", ".")
	strSQL = "INSERT INTO MERFL.MER583F2 (COBEC2,CDESC2,CPROC2,NROAC2,FANAC2,COANC2,VACAC2,PREBC2,PBONC2,ESTAC2,COCAC2,TIPAC2) VALUES "&_
			 " ("&pCodigoBolsa&",0,"& pCdProducto &","& pNroAnalisis &","& pFecha &",0,"& pKilosReb &",0,0,'',"& pCdRubro &", " & tipo & ") "		
	Call executeQuery(rs, "EXEC", strSQL)
	if (pICCamara = "N") then
	    strSQL = "INSERT INTO MERFL.MER582F2 (COBEC2,CDESC2,CPROC2,NROAC2,FANAC2,COANC2,VACAC2,PREBC2,PBONC2,ESTAC2,COCAC2,TIPAC2) VALUES "&_
			     " ("&pCodigoBolsa&",0,"& pCdProducto &","& pNroAnalisis &","& pFecha &",0,"& pKilosReb &",0,0,'',"& pCdRubro &", " & tipo & ") "		
	    Call executeQuery(rs, "EXEC", strSQL)
	end if
End Function
'-----------------------------------------------------------------------------------------------------------------
Function controlarCtaPteVagon(pCdVagonOld,pCtaPteOld,pRs)
	Dim rtrn		
	rtrn  = false
	if (not pRs.EoF) then
		if((pCtaPteOld = Cdbl(pRs("NUCARTAPORTE")))and(pCdVagonOld = Cdbl(pRs("CDVAGON"))))then rtrn = true
	end if
	controlarCtaPteVagon = rtrn
End Function
'-----------------------------------------------------------------------------------------------------------------
'****************************************************
'*****          COMIENZO DE LA PAGINA           *****
'***************************************************
Dim myHoy,logMig, transporte,myNext,flagGenerarFecha, origenFile

on error resume next

origen = GF_PARAMETROS7("p", "", 6)
origenFile = origen
if (origenFile = "") then origenFile = "ERROR"
myHoy = GF_PARAMETROS7("f", 0, 6)
transporte = GF_PARAMETROS7("t", 0, 6)

Set logMig = new classLog
Call startLog(HND_VIEW+HND_FILE,MSG_INF_LOG+MSG_ERR_LOG+MSG_WRN_LOG)
logMig.fileName = "ANALISIS-SYNC-"& origenFile &"-"& left(session("MmtoDato"),8)		

logMig.info("####################################################")
logMig.info("******* MIGRACION DE ANALISIS *********")
logMig.info("		-PUERTO       :  "& origen)
logMig.info("		-MOMENTO      :  "& GF_FN2DTE(Left(session("MmtoSistema"),8)))
logMig.info("		-TRANSPORTE   :  "& transporte )	
logMig.info("		-USUARIO      :  "& session("Usuario"))	
logMig.info("		-FECHA MIGRADA:  "& GF_FN2DTE(myHoy))
logMig.info("####################################################")	

if (origen = "") then Response.end
g_strPuerto = origen
Call migrarAnalisis(myHoy,origen, transporte)
myBolsa = getCodigoBolsa(origen)
if (myBolsa <> -1) then
    logMig.info("EJECUTANDO STORED.PGM_MER59A...")
    Call executeSP(rs, "STORED.PGM_MER59A", myBolsa)
else
    logMig.info("ERROR MIGRANDO BOLSA DATOS!! Puerto: " & g_strPuerto & ", Bolsa:" & myBolsa)	
end if
myNext = GF_DTEADD(myHoy,1,"D")

%>
<HTML>
	<HEAD>
		<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
		<script type="text/javascript">
			parent.generateSegment_callback('<%=myNext%>');
		</script>
	</HEAD>
	<BODY>
		<P>&nbsp;</P>
	</BODY>
</HTML>
