<%
dim myTurno, myOperativo, myFecContableD, myFecContableM, myFecContableA, myFecContableH, myFecContableN, myFecContableDesde
dim myFecContableDH, myFecContableMH, myFecContableAH, myFecContableHH, myFecContableNH, myFecContableHasta
dim myCdCoordinador, myDsCoordinador, myCdCoordinado, myDsCoordinado, myCdProducto, myIdVagon,myGrado(13,6)
dim myCdCorredor, myDsCorredor, myCdVendedor, myDsVendedor, myCdEntregador, myDsEntregador,sortBy,myKiloNeto
dim strSQLInnerH, strSQLInner,strPath,myFileCode,params,myCartaPorte,orderByCorredor,orderByEntregador,orderByVendedor,orderByLocalidad
dim orderByTurno,orderByOperativo,orderByCartaPorte,orderByFecha,orderByCoordinado,orderByProducto,myOrder,orderByInicio,orderByEstado,myEstado,orderByFechaInicio
'---------------------------------------------------------
sub getParametros()
		myOperativo = GF_PARAMETROS7("Operativo", "", 6)
		call addParam("Operativo", myOperativo, params)
				
		myCartaPorte = GF_PARAMETROS7("CartaPorte", "", 6)
		call addParam("CartaPorte", myCartaPorte, params)
		
		myTurno = GF_PARAMETROS7("turno", "", 6)
		call addParam("turno", myTurno, params)
		
		myIdVagon = GF_PARAMETROS7("nroVagon", "", 6)
		Call addParam("nroVagon", myIdVagon, params)
				
		myFecContableD = GF_PARAMETROS7("fecContableD", "", 6)
		if (myFecContableD = "") then myFecContableD=Day(Now())
		call addParam("fecContableD", myFecContableD, params)
		myFecContableM = GF_PARAMETROS7("fecContableM", "", 6)
		if (myFecContableM = "") then myFecContableM=Month(Now())
		Call addParam("fecContableM", myFecContableM, params)
		myFecContableA = GF_PARAMETROS7("fecContableA", "", 6)
		if (myFecContableA = "") then myFecContableA=Year(Now())
		Call addParam("fecContableA", myFecContableA, params)
		myFecContableDesde = myFecContableA & myFecContableM & myFecContableD

		myFecContableDH = GF_PARAMETROS7("fecContableDH", "", 6)
		if (myFecContableDH = "") then myFecContableDH=Day(Now())
		call addParam("fecContableDH", myFecContableDH, params)
		myFecContableMH = GF_PARAMETROS7("fecContableMH", "", 6)
		if (myFecContableMH = "") then myFecContableMH=Month(Now())
		call addParam("fecContableMH", myFecContableMH, params)
		myFecContableAH = GF_PARAMETROS7("fecContableAH", "", 6)
		if (myFecContableAH = "") then myFecContableAH=Year(Now())
		call addParam("fecContableAH", myFecContableAH, params)
		
		ret = GF_CONTROL_PERIODO(myFecContableD, myFecContableDH, myFecContableM, myFecContableMH, myFecContableA, myFecContableAH)
		flagCall = false
		Select case (ret)
			case 0	
				flagCall=true
				myFecContableDesde = myFecContableA  & myFecContableM  & myFecContableD
				myFecContableHasta = myFecContableAH & myFecContableMH & myFecContableDH
			case 1
				Call setError(FECHA_INICIO_INCORRECTA)
			case 2
				Call setError(FECHA_FIN_INCORRECTA)
			case 3
				Call setError(PERIODO_ERRONEO)
		end select
		
		myCdCoordinador = GF_PARAMETROS7("cdCoordinador", "", 6)
		call addParam("cdCoordinador", myCdCoordinador, params)
		myDsCoordinador = Trim(GF_PARAMETROS7("dsCoordinador", "", 6))
		call addParam("dsCoordinador", myDsCoordinador, params)

		myCdCoordinado = GF_PARAMETROS7("cdCoordinado", "", 6)
		call addParam("cdCoordinado", myCdCoordinado, params)
		myDsCoordinado = Trim(GF_PARAMETROS7("dsCoordinado", "", 6))
		call addParam("dsCoordinado", myDsCoordinado, params)

		myCdProducto = GF_PARAMETROS7("cdProducto", "", 6)
		call addParam("cdProducto", myCdProducto, params)

		myCdCorredor = GF_PARAMETROS7("cdCorredor", "", 6)
		call addParam("cdCorredor", myCdCorredor, params)
		myDsCorredor = Trim(GF_PARAMETROS7("dsCorredor", "", 6))
		call addParam("dsCorredor", myDsCorredor, params)

		myCdVendedor = GF_PARAMETROS7("cdVendedor", "", 6)
		call addParam("cdVendedor", myCdVendedor, params)
		myDsVendedor = Trim(GF_PARAMETROS7("dsVendedor", "", 6))
		call addParam("dsVendedor", myDsVendedor, params)

		myCdEntregador = GF_PARAMETROS7("cdEntregador", "", 6)
		call addParam("cdEntregador", myCdEntregador, params)
		myDsEntregador = Trim(GF_PARAMETROS7("dsEntregador", "", 6))
		call addParam("dsEntregador", myDsEntregador, params)
		
		myEstado = GF_PARAMETROS7("cmbEstado", 0, 6)
		call addParam("cmbEstado", myEstado, params)
		
		myFileCode = GF_PARAMETROS7("fileCode", "", 6)
		call addParam("fileCode", myFileCode, params)
		myOrder = GF_PARAMETROS7("myOrder", "" ,6)
		call addParam("myOrder", myOrder, params)
		if myOrder = "" then myOrder = " T1.DTCONTABLE DESC "
				
		orderByTurno = GF_PARAMETROS7("orderByTurno","",6)
		orderByOperativo = GF_PARAMETROS7("orderByOperativo","",6)
		orderByCartaPorte = GF_PARAMETROS7("orderByCartaPorte","",6)
		orderByFechaInicio = GF_PARAMETROS7("orderByFechaInicio","",6)
		orderByCoordinado = GF_PARAMETROS7("orderByCoordinado","",6)
		orderByProducto = GF_PARAMETROS7("orderByProducto","",6)
		orderByCorredor = GF_PARAMETROS7("orderByCorredor","",6)		
		orderByVendedor = GF_PARAMETROS7("orderByVendedor","",6)		
		orderByEstado = GF_PARAMETROS7("orderByEstado","",6)
end sub		
'-------------------------------------------------------------------------------------------------------------
Function filtrarOperativosPuertos()
	Dim strSqlFiltro
	If myOperativo <> "" Then 	    
	    auxOpe = Trim(myOperativo)
		IF Len(myOperativo) < 12 Then 
		    auxOpe = Trim(myOperativo)
			auxSer = "0000"
		else
		    auxOpe = Mid(Trim(myOperativo), 5, 12) & "0000"
			auxSer = Mid(Trim(myOperativo), 1, 4)	    	        
		end if
		strSqlFiltro = strSqlFiltro & " and T1.cdOperativo ='" & auxOpe & "'"
		strSqlFiltro = strSqlFiltro & " and T1.cdOperativoSerie ='" &auxSer& "'"
		
	end if
	If myCartaPorte <> "" Then 
		IF Len(myCartaPorte) < 12 Then 
			auxOpe = Trim(myCartaPorte)
			auxSer = "0000"
		else
			auxOpe = Mid(Trim(myCartaPorte), 5, 12)
			auxSer = Mid(Trim(myCartaPorte), 1, 4)
		end if
		strSqlFiltro = strSqlFiltro & " and T1.cdOperativo ='" & auxOpe & "'"
		strSqlFiltro = strSqlFiltro & " and T1.cdOperativoSerie ='" &auxSer& "'"
	end if
	If myFecContableDesde <> "" Then 		
		strSqlFiltro = strSqlFiltro & " and T1.DTINICIO >= " & myFecContableDesde
	end if
	If myFecContableHasta <> "" Then
		strSqlFiltro = strSqlFiltro & " and T1.DTINICIO <= " & myFecContableHasta
	end if	
	If myTurno <> "" Then strSqlFiltro = strSqlFiltro & " and T1.sqTurno = " & Trim(myTurno)
	If myCdCoordinador <> "" Then strSqlFiltro = strSqlFiltro & " and T1.cdEmpresa = " & myCdCoordinador
	If myCdCoordinado <> "" Then strSqlFiltro = strSqlFiltro & " and T1.cdCliente = " & myCdCoordinado
	If myCdProducto <> "" Then strSqlFiltro = strSqlFiltro & " and T1.cdProducto = " & myCdProducto
	If myCdVendedor <> "" Then strSqlFiltro = strSqlFiltro & " and T1.cdVendedor = " & CLng(myCdVendedor)
	If myCdCorredor <> "" Then strSqlFiltro = strSqlFiltro & " and T1.cdCorredor = " & CLng(myCdCorredor)
	If myCdEntregador <> "" Then strSqlFiltro = strSqlFiltro & " and T1.cdEntregador = " & CLng(myCdEntregador)
	if (myIdVagon <> "") then 
	    'Se busca el vagon y se filtra por todos los operativos que tengan vagones con este vagon.
	    strSQLTemp = "Select CDOPERATIVO from ((Select CDOPERATIVO, CDVAGON from VAGONES where CDVAGON='" & CLng(myIdVagon) & "') union (Select CDOPERATIVO, CDVAGON from HVAGONES where CDVAGON='" & CLng(myIdVagon) & "')) T "
	    strSqlFiltro = strSqlFiltro & " and T1.CDOPERATIVO in (" & strSQLTemp & ")" 
	end if
	If myEstado <> 0 then strSqlFiltro = strSqlFiltro & " and EST.CDESTADO = " & myEstado	
	filtrarOperativosPuertos = strSqlFiltro
End function
'-------------------------------------------------------------------------------------------------------------
Function loadOperativosPuertos()
	Dim strSQL ,diaHoy,myWhere
	myWhere = filtrarOperativosPuertos()
	diaHoy = Year(Now()) & GF_nDigits(Month(Now()), 2) & GF_nDigits(Day(Now()), 2)
	strSQL = " SELECT  T1.*, T2.DSPROCEDENCIA,P.DSPRODUCTO,CL.DSCLIENTE,CO.DSCORREDOR,VE.DSVENDEDOR,EM.DSEMPRESA,ENT.DSENTREGADOR,EST.DSESTADO "&_
			 "  from (( "&_
			 "			SELECT  '"& diaHoy &"' as DTCONTABLE,O.CDOPERATIVO,O.NUCARTAPORTE, ((YEAR(O.DTEMISION)*10000) + (MONTH(O.DTEMISION)*100) + DAY(O.DTEMISION)) as DTEMISION,O.CDPRODUCTO, "&_
			 "					O.CDEMPRESA,O.CDCLIENTE,O.CDVENDEDOR,O.CDPROCEDENCIA,O.CDENTREGADOR,O.CDDESTINO, "&_
			 "					O.CDTRANSPORTISTA,O.CDCORREDOR,O.CDTIPOTRANSPORTE, ((YEAR(O.DTARRIBO)*10000) + (MONTH(O.DTARRIBO)*100) + DAY(O.DTARRIBO)) as DTARRIBO, ((YEAR(O.dtinicio)*10000) + (MONTH(O.dtinicio)*100) + DAY(O.dtinicio)) as dtinicio, ((YEAR(O.DTFIN)*10000) + (MONTH(O.DTFIN)*100) + DAY(O.DTFIN)) as DTFIN,O.QTVAGONES, "&_
			 "					O.CDESTADO,O.SQOPERATIVO,O.SQTURNO,O.VLNETOTOTAL,O.NUCARTAPORTESERIE,O.CDOPERATIVOSERIE, "&_			 
			 "					CASE O.NUCARTAPORTESERIE WHEN '0000' THEN O.NUCARTAPORTE ELSE O.NUCARTAPORTESERIE + O.NUCARTAPORTE END CartaPorte, "&_
			 "					CASE O.CDOPERATIVOSERIE WHEN '0000' THEN O.CDOPERATIVO ELSE O.CDOPERATIVOSERIE + O.CDOPERATIVO END Operativo "&_
			 "			FROM dbo.operativos O "&_
			 "			WHERE O.CDTIPOTRANSPORTE = 2 ) "&_
			 " UNION  "&_
			 "		 (  SELECT  ((YEAR(O.DTCONTABLE)*10000) + (MONTH(O.DTCONTABLE)*100) + DAY(O.DTCONTABLE)) as DTCONTABLE,O.CDOPERATIVO,O.NUCARTAPORTE,((YEAR(O.DTEMISION)*10000) + (MONTH(O.DTEMISION)*100) + DAY(O.DTEMISION)) as DTEMISION,O.CDPRODUCTO,O.CDEMPRESA,O.CDCLIENTE, "&_
			 "					O.CDVENDEDOR,O.CDPROCEDENCIA,O.CDENTREGADOR,O.CDDESTINO,O.CDTRANSPORTISTA,O.CDCORREDOR, "&_
			 "					O.CDTIPOTRANSPORTE, ((YEAR(O.DTARRIBO)*10000) + (MONTH(O.DTARRIBO)*100) + DAY(O.DTARRIBO)) as DTARRIBO, ((YEAR(O.dtinicio)*10000) + (MONTH(O.dtinicio)*100) + DAY(O.dtinicio)) as dtinicio, ((YEAR(O.DTFIN)*10000) + (MONTH(O.DTFIN)*100) + DAY(O.DTFIN)) as DTFIN,O.QTVAGONES,O.CDESTADO,O.SQOPERATIVO, "&_
             "				    O.SQTURNO,O.VLNETOTOTAL,O.NUCARTAPORTESERIE,O.CDOPERATIVOSERIE, "&_             
             "					CASE O.NUCARTAPORTESERIE WHEN '0000' THEN O.NUCARTAPORTE ELSE O.NUCARTAPORTESERIE + O.NUCARTAPORTE END CartaPorte, "&_
             "					CASE O.CDOPERATIVOSERIE WHEN '0000' THEN O.CDOPERATIVO ELSE O.CDOPERATIVOSERIE + O.CDOPERATIVO END Operativo "&_
			 "			FROM dbo.hoperativos O "&_
			 "			WHERE O.CDTIPOTRANSPORTE = 2 )) T1 " &_
			 " LEFT JOIN dbo.PROCEDENCIAS T2 on T2.CDPROCEDENCIA = T1.CDPROCEDENCIA "&_
			 " INNER JOIN dbo.PRODUCTOS P ON T1.CDPRODUCTO=P.CDPRODUCTO "&_
		     " INNER JOIN dbo.CLIENTES CL ON T1.CDCLIENTE=CL.CDCLIENTE "&_
		     " INNER JOIN dbo.CORREDORES CO ON T1.CDCORREDOR=CO.CDCORREDOR "&_
		     " INNER JOIN dbo.VENDEDORES VE ON T1.CDVENDEDOR=VE.CDVENDEDOR "&_
		     " INNER JOIN dbo.EMPRESAS EM ON T1.CDEMPRESA=EM.CDEMPRESA "&_
		     " INNER JOIN dbo.ENTREGADORES ENT ON T1.CDENTREGADOR=ENT.CDENTREGADOR "&_
		     " INNER JOIN dbo.ESTADOSOPERATIVOS EST ON T1.CDESTADO=EST.CDESTADO "&_		     
		     " WHERE 1 = 1 "& myWhere &_
		     " ORDER BY "& myOrder
	Call GF_BD_Puertos(g_strPuerto, rs, "OPEN", strSQL)	

	Set loadOperativosPuertos= rs
End Function
'---------------------------------------------------------------------------------------------------------------------------
Function loadVagonesByOperativosPuertos(pCdOperativo, pDtContable, pCartaPorte, pPto)
	Dim strSQL, diaHoy, rsVag
	diaHoy = Year(Now()) & "-" & GF_nDigits(Month(Now()), 2) & "-" & GF_nDigits(Day(Now()), 2) 
	iEstadoTara = 8          
	'Response.Write strSqlFiltro	
		strSQL = " Select	o.SQTURNO Turno, "
		strSQL = strSQL & " o.cdoperativo, "		
		strSQL = strSQL & " o.NUCARTAPORTE, "
		strSQL = strSQL & " va.dtContableVagon Fecha, "
		strSQL = strSQL & " pe.dtPesada, "
		strSQL = strSQL & " va.cdvagon NoVagon, "
		strSQL = strSQL & " CASE o.NUCARTAPORTESERIE WHEN '0000' THEN o.NUCARTAPORTE ELSE o.NUCARTAPORTESERIE + o.NUCARTAPORTE END CartaPorte, "
		strSQL = strSQL & " CASE o.cdOperativoSERIE WHEN '0000' THEN o.cdOperativo ELSE o.cdOperativoSerie + o.cdOperativo END Operativo, "
		strSQL = strSQL & " (select pv.vlPesada from dbo.HPesadasVagon pv where pv.dtcontable = va.dtcontable and pv.cdOperativo = va.cdOperativo and pv.cdOperativoSERIE = va.cdOperativoSERIE"
		strSQL = strSQL & " and pv.nuCartaPorte = va.nuCartaPOrte and pv.nuCartaPorteSERIE = va.nuCartaPOrteSERIE and pv.cdVagon = va.cdVagon "
		strSQL = strSQL & " and pv.cdPesada = 1 and pv.sqpesada = (Select Max(sqPesada) from dbo.HPesadasVagon"
		strSQL = strSQL & " where pv.dtcontable = dtcontable and pv.cdOperativo = cdOperativo and pv.nuCartaporte = nuCartaPorte and pv.cdOperativoSERIE = cdOperativoSERIE and pv.nuCartaporteSERIE = nuCartaPorteSERIE "
		strSQL = strSQL & " and pv.cdVagon = cdVagon and cdPesada = 1)) Bruto,"
		strSQL = strSQL & " (select pv.vlPesada from dbo.HPesadasVagon pv where pv.dtcontable = va.dtcontable and pv.cdOperativo = va.cdOperativo and pv.cdOperativoSERIE = va.cdOperativoSERIE"
		strSQL = strSQL & " and pv.nuCartaPorte = va.nuCartaPorte and pv.nuCartaPorteSERIE = va.nuCartaPorteSERIE and pv.cdVagon = va.cdVagon "
		strSQL = strSQL & " and pv.cdPesada = 2 and pv.sqpesada = (Select Max(sqPesada) from dbo.HPesadasVagon"
		strSQL = strSQL & " where pv.dtcontable = dtcontable and pv.cdOperativo = cdOperativo and pv.nuCartaporte = nuCartaPorte and pv.cdOperativoSERIE = cdOperativoSERIE and pv.nuCartaporteSERIE = nuCartaPorteSERIE "
		strSQL = strSQL & " and pv.cdVagon = cdVagon and cdPesada = 2)) Tara,"
		strSQL = strSQL & " (select mv.vlMermaKilos from dbo.HMermasVagones mv where mv.dtcontable = va.dtcontable and mv.cdOperativo = va.cdOperativo and mv.cdOperativoSERIE = va.cdOperativoSERIE"
		strSQL = strSQL & " and mv.nuCartaPorte = va.nuCartaPOrte and mv.nuCartaPorteSERIE = va.nuCartaPOrteSERIE and mv.cdVagon = va.cdVagon "
		strSQL = strSQL & " and mv.sqpesada = (Select Max(sqPesada) from dbo.HPesadasVagon"
		strSQL = strSQL & " where mv.dtcontable = dtcontable and mv.cdOperativo = cdOperativo and mv.nuCartaporte = nuCartaPorte and mv.cdOperativoSERIE = cdOperativoSERIE and mv.nuCartaporteSERIE = nuCartaPorteSERIE "
		strSQL = strSQL & " and mv.cdVagon = cdVagon and cdPesada = 2)) Merma,"
		strSQL = strSQL & " 0 Netos,"
		strSQL = strSQL & " cld.nuBarras Barras,"		
		strSQL = strSQL & " cld.cdAceptacion, "
		strSQL = strSQL & " ac.dsaceptacion Aceptacion, "
		strSQL = strSQL & " pro.dsProducto, "
		strSQL = strSQL & " o.cdProducto "
		strSQL = strSQL & " From dbo.HVagones va "
		strSQL = strSQL & " Left Join dbo.HCaladadeVagones cld on cld.dtContable = va.dtContable "
		strSQL = strSQL & " and cld.cdOperativo = va.cdoperativo and cld.nuCartaPorte = va.nuCartaporte and cld.cdOperativoSERIE = va.cdoperativoSERIE and cld.nuCartaPorteSERIE = va.nuCartaporteSERIE "
		strSQL = strSQL & " and cld.cdVagon = va.cdVagon "
		strSQL = strSQL & " Left Join dbo.HPesadasVagon pe on pe.dtContable = va.dtContable "
		strSQL = strSQL & " and pe.cdOperativo = va.cdoperativo and pe.nuCartaPorte = va.nuCartaporte and pe.cdOperativoSERIE = va.cdoperativoSERIE and pe.nuCartaPorteSERIE = va.nuCartaporteSERIE "
		strSQL = strSQL & " and pe.cdVagon = va.cdVagon "
		strSQL = strSQL & " Left Join dbo.AceptacionCalidad ac on cld.cdaceptacion = ac.cdaceptacion"		
		strSQL = strSQL & " , dbo.HOperativos o  "
		strSQL = strSQL & " left join dbo.productos pro on o.cdproducto = pro.cdproducto "
		strSQL = strSQL & " Where va.cdestado = 8 "
		strSQL = strSQL & " and va.dtcontable = o.dtcontable and va.cdOperativo = o.cdOperativo and va.cdOperativoSERIE = o.cdOperativoSERIE "
		strSQL = strSQL & " and va.nuCartaPorte = o.nuCartaPorte and va.nuCartaPorteSERIE = o.nuCartaPorteSERIE "
		strSQL = strSQL & " and pe.sqPesada = (Select Max(SqPesada) from dbo.HPesadasVagon "
		strSQL = strSQL & " where dtcontable = pe.dtcontable and cdOperativo = pe.cdOperativo and cdOperativoSERIE = pe.cdOperativoSERIE "
		strSQL = strSQL & " and nuCartaPorte = pe.nuCartaPorte and nuCartaPorteSERIE = pe.nuCartaPorteSERIE and cdVagon = pe.cdVAgon and cdPesada=2)"
		strSQL = strSQL & " and cld.sqcalada = (Select Max(SqCalada) from dbo.HCaladadeVagones "
		strSQL = strSQL & " where dtcontable = cld.dtcontable and cdOperativo = cld.cdOperativo and cdOperativoSERIE = cld.cdOperativoSERIE "
		strSQL = strSQL & " and nuCartaPorte = cld.nuCartaPorte and nuCartaPorteSERIE = cld.nuCartaPorteSERIE and cdVagon = cld.cdVagon) "
		strSQL = strSQL & " and o.cdOperativo = '" & pCdOperativo &"' "
		strSQL = strSQL & " and o.NUCARTAPORTE = '" & pCartaPorte &"' "
		strSQL = strSQL & " and va.dtContableVagon =	'" & pDtContable & "'"
		strSQL = strSQL & " Union Select "
		strSQL = strSQL & " o.SQTURNO Turno, "
		strSQL = strSQL & " o.cdoperativo, "
		strSQL = strSQL & " o.NUCARTAPORTE, va.dtContableVagon Fecha,"		
		strSQL = strSQL & " pe.dtPesada,"
		strSQL = strSQL & " va.cdvagon NoVagon,"
		strSQL = strSQL & " CASE o.NUCARTAPORTESERIE WHEN '0000' THEN o.NUCARTAPORTE ELSE o.NUCARTAPORTESERIE + o.NUCARTAPORTE END CartaPorte, "
		strSQL = strSQL & " CASE o.cdOperativoSERIE WHEN '0000' THEN o.cdOperativo ELSE o.cdOperativoSerie + o.cdOperativo END Operativo,"
		strSQL = strSQL & " (select pv.vlPesada from dbo.PesadasVagon pv where "
		strSQL = strSQL & " pv.nuCartaPorte = va.nuCartaPOrte and pv.nuCartaPorteSERIE = va.nuCartaPOrteSERIE and pv.cdVagon = va.cdVagon "
		strSQL = strSQL & " and pv.cdPesada = 1 and pv.sqpesada = (Select Max(sqPesada) from dbo.PesadasVagon"
		strSQL = strSQL & " where pv.cdOperativo = cdOperativo and pv.nuCartaporte = nuCartaPorte and pv.cdOperativoSERIE = cdOperativoSERIE and pv.nuCartaporteSERIE = nuCartaPorteSERIE "
		strSQL = strSQL & " and pv.cdVagon = cdVagon and cdPesada = 1)) Bruto,"
		strSQL = strSQL & " (select pv.vlPesada from dbo.PesadasVagon pv where pv.cdOperativo = va.cdOperativo and pv.cdOperativoSERIE = va.cdOperativoSERIE"
		strSQL = strSQL & " and pv.nuCartaPorte = va.nuCartaPOrte and pv.nuCartaPorteSERIE = va.nuCartaPOrteSERIE and pv.cdVagon = va.cdVagon "
		strSQL = strSQL & " and pv.cdPesada = 2 and pv.sqpesada = (Select Max(sqPesada) from dbo.PesadasVagon"
		strSQL = strSQL & " where pv.cdOperativo = cdOperativo and pv.nuCartaporte = nuCartaPorte and pv.cdOperativoSERIE = cdOperativoSERIE and pv.nuCartaporteSERIE = nuCartaPorteSERIE "
		strSQL = strSQL & " and pv.cdVagon = cdVagon and cdPesada = 2)) Tara,"
		strSQL = strSQL & " (select mv.vlMermaKilos from dbo.MermasVagones mv where mv.cdOperativo = va.cdOperativo and mv.cdOperativoSERIE = va.cdOperativoSERIE"
		strSQL = strSQL & " and mv.nuCartaPorte = va.nuCartaPOrte and mv.nuCartaPorteSERIE = va.nuCartaPOrteSERIE and mv.cdVagon = va.cdVagon "
		strSQL = strSQL & " and mv.sqpesada = (Select Max(sqPesada) from dbo.PesadasVagon"
		strSQL = strSQL & " where mv.cdOperativo = cdOperativo and mv.nuCartaporte = nuCartaPorte and mv.cdOperativoSERIE = cdOperativoSERIE and mv.nuCartaporteSERIE = nuCartaPorteSERIE "
		strSQL = strSQL & " and mv.cdVagon = cdVagon and cdPesada = 2)) Merma,"
		strSQL = strSQL & " 0 Netos,"
		strSQL = strSQL & " cld.nuBarras Barras,"
		strSQL = strSQL & " cld.cdAceptacion, "
		strSQL = strSQL & " ac.dsaceptacion Aceptacion, "
		strSQL = strSQL & " pro.dsProducto, "
		strSQL = strSQL & " o.cdProducto "		
		strSQL = strSQL & " From dbo.Vagones va "
		strSQL = strSQL & " Left Join dbo.CaladadeVagones cld on "
		strSQL = strSQL & " cld.cdOperativo = va.cdoperativo and cld.nuCartaPorte = va.nuCartaporte and  cld.cdOperativoSERIE = va.cdoperativoSERIE and cld.nuCartaPorteSERIE = va.nuCartaporteSERIE "
		strSQL = strSQL & " and cld.cdVagon = va.cdVagon "
		strSQL = strSQL & " Left Join dbo.PesadasVagon pe on "
		strSQL = strSQL & " pe.cdOperativo = va.cdoperativo and pe.nuCartaPorte = va.nuCartaporte and pe.cdOperativoSERIE = va.cdoperativoSERIE and pe.nuCartaPorteSERIE = va.nuCartaporteSERIE "
		strSQL = strSQL & " and pe.cdVagon = va.cdVagon "
		strSQL = strSQL & " Left Join dbo.AceptacionCalidad ac on cld.cdaceptacion = ac.cdaceptacion"
		strSQL = strSQL & " ,dbo.Operativos o  "		
		strSQL = strSQL & " left join dbo.productos pro on o.cdproducto = pro.cdproducto "
		strSQL = strSQL & " Where va.cdestado = 8 "
		strSQL = strSQL & " and va.cdOperativo = o.cdOperativo and va.cdOperativoSERIE = o.cdOperativoSERIE "
		strSQL = strSQL & " and va.nuCartaPorte = o.nuCartaPorte and va.nuCartaPorteSERIE = o.nuCartaPorteSERIE "
		strSQL = strSQL & " and pe.sqPesada = (Select Max(SqPesada) from dbo.PesadasVagon "
		strSQL = strSQL & " where cdOperativo = pe.cdOperativo and cdOperativoSERIE = pe.cdOperativoSERIE "
		strSQL = strSQL & " and nuCartaPorte = pe.nuCartaPorte and nuCartaPorteSERIE = pe.nuCartaPorteSERIE and cdVagon = pe.cdVAgon and cdPesada=2)"
		strSQL = strSQL & " and cld.sqcalada = (Select Max(SqCalada) from dbo.CaladadeVagones "
		strSQL = strSQL & " where cdOperativo = cld.cdOperativo and cdOperativoSERIE = cld.cdOperativoSERIE "
		strSQL = strSQL & " and nuCartaPorte = cld.nuCartaPorte and nuCartaPorteSERIE = cld.nuCartaPorteSERIE and cdVagon = cld.cdVagon) "
		strSQL = strSQL & " and o.cdOperativo = '" & pCdOperativo &"' "
		strSQL = strSQL & " and o.NUCARTAPORTE = '" & pCartaPorte &"' "
		strSQL = strSQL & " and va.dtContableVagon ='" & pDtContable & "'"		
		Call GF_BD_Puertos(pPto, rsVag, "OPEN", strSQL)
		'Response.Write strSQL
	Set loadVagonesByOperativosPuertos = rsVag
End Function
'---------------------------------------------------------------------------------------------------------------------------
Function FormatFechaSQL(ByVal pFecha)
dim myFecha, myFechaIni
	myFecha = year(pFecha) & "/" & month(pFecha) & "/" & day(pFecha)
	myFechaIni = year(pFecha) & "/01/01"
    FormatFechaSQL = CStr(Year(myFecha) & DateDiff("d", myFechaIni, myFecha) + 1)
End Function
'---------------------------------------------------------------------------------------------------------------------------
function VerGrado(pPto,nAceptacion, sBarras, dFecha)
Dim rs, RS2, strSQL, rsensayo, rsGrado, rsresultado, rsCertificado, rsHonorarios, rsTipo, rsBonifreBaja, rsBaseoTolerancia, rsFueraStandard, auxResul, iEstadarBase, myDtContable
    		
myDtContable = GF_FN2DTCONTABLE(dFecha)

sEngra = "00000"
vCores = 1000
sGrado = "FAL"
If UCASE(pPto) = "PIEDRABUENA" Then        
    If nAceptacion = 2 then
        strSQL = "Select cdEnsayo,cdGrado,cdResultado,nuCertificado,vlHonorarios,icTipo,vlBonifreBaja, "
        strSQL = strSQL & " icFueraStandard, icBaseoTolerancia from dbo.ResultadosCamara "
        strSQL = strSQL & " where nuBArras = '" & sBarras & "' "
        strSQL = strSQL & " and dtcontable = '" & myDtContable & "' "
        strSQL = strSQL & " order by cdEnsayo "
        'Response.Write "<br>" & strSQL
        'Response.End
        call GF_BD_Puertos(pPto, rs, "OPEN", strSQL)
        TotEnsayos = 0
        TotEnsCond = 0
        TotEnsCondSi = 0
        TotEnsCondNo = 0
        If rs.EOF Then
            sGrado = "SAC"
        Else
            Do Until rs.EOF
                rsensayo = VerNull(rs("CdEnsayo"))
                rsGrado = VerNull(rs("CdGrado"))
                rsresultado = VerNull(rs("cdResultado"))
                rsCertificado = VerNull(rs("nuCertificado"))
                rsHonorarios = VerNull(rs("vlHonorarios"))
                rsTipo = VerNull(rs("icTipo"))
                rsBonifreBaja = VerNull(rs("vlBoniFreBaja"))
                rsBaseoTolerancia = VerNull(rs("icBaseoTolerancia"))
                rsFueraStandard = VerNull(rs("icFueraStandard"))
                If (VerNull(rs("CdEnsayo")) = sEngra) Or (VerNull(rs("CdEnsayo")) = "00001") Then Exit Do
                rs.MoveNext
                If rs.EOF Then Exit Do
            Loop
            sCertificado = rsCertificado
                
            If rsensayo = sEngra Then
                auxResul = rsresultado / vCores
                Select Case auxResul
                    Case 1
                        sGrado = "1"
                    Case 2
                        sGrado = "2"
                    Case 3
                        sGrado = "3"
                    Case Else
                        sGrado = "SAC"
                End Select
            Else
                If rsensayo = "00001" Then
                    auxResul = rsresultado / vCores
                    Select Case auxResul
                        Case 1
                            sGrado = "F1"
                        Case 2
                            sGrado = "F2"
                        Case 3
                            sGrado = "F3"
                        Case Else
                            sGrado = "SAC"
                    End Select
                Else
                    rs.MoveFirst
                                                             
                End If
            End If
        End If
    Else
        If nAceptacion = 1 then 'Val(GetParametro("CDACEPTACIONCONFORME", "C�digo de Aceptacion Conforme", "1")) Then
            sGrado = "CON"
        Else
            sGrado = "SAO"
        End If
    End If    
End If
VerGrado = sGrado
End function
'------------------------------------------------------------------------------------------------------------------------
Function VerEstandares(pPto, nProducto, sBarras, ByRef rsensayo, ByRef rsresultado)
Dim rs, strSQL
strSQL = "Select rc.cdEnsayo, rc.cdResultado from dbo.ResultadosCamara rc, dbo.EstandarEnsayos ee"
strSQL = strSQL & " where rc.nubarras = '" & sBarras & "' and ee.cdProducto =" & nProducto
strSQL = strSQL & " and rc.cdensayo = ee.cdensayo and (rc.cdresultado < ee.vlddecondok or rc.cdresultado > ee.vlhtacondok )"
strSQL = strSQL & " order by rc.cdensayo"
'Set rs = AdoConnection.Execute(strSQL)
call GF_BD_Puertos(pPto, rs, "OPEN", strSQL)
If Not rs.EOF Then
    rsensayo = VerNull(rs(0))
    rsresultado = VerNull(rs(1))
    VerEstandares = True
End If
Set rs = Nothing
End Function
'---------------------------------------------------------------------------------------------------------------------------
Function VerNull(dato)
    If IsNull(dato) Then
        VerNull = Empty
    Else
        VerNull = dato
    End If
End Function
'---------------------------------------------------------------------------------------------------------------------------
Sub SumarResumen(sGrado, pKilos)
Dim i
For i = 0 To 13
	If myGrado(i,0) = sGrado Then
		Exit For
    end If
Next
If i <= 13 Then
    myGrado(i,2) = clng(myGrado(i,2)) + 1
    myGrado(i,4) = clng(myGrado(i,4)) + CLng(pKilos)
    'ItemGrado(i).TotalKgNetoGrado = ItemGrado(i).TotalKgNetoGrado + CLng(lsv.ListItems(Fila).SubItems(15))
End If
End Sub
Sub ImprimirResumen()
Dim i
i = 1

For i = 0 To 13
	for j=0 to 5
		Response.Write "<hr>Indice i,j (" & i & "," & j & ") Valor(" & myGrado(i,j) & ")"
    next
Next
End Sub
'---------------------------------------------------------------------------------------------------------------------------
Sub Sumar_Totales(pKilos, pFila)
If pFila = 0 Then
	TotKgNeto = 0
	TotVagones = 0
Else
	TotVagones = TotVagones + 1
    If IsNumeric(pKilos) Then TotKgNeto = TotKgNeto + pKilos
End If
End Sub
'--------------------------------------------------------------------------------------------------------------------------------------------
Sub CargarGrados()
    myGrado(0,0) = "FAL"					'Codigo
    myGrado(0,1) = "Total faltante Grado"	'Descripcion
    myGrado(0,2) = 0						'Total Grado 
    myGrado(0,3) = 0						'Porcetaje Total Grado
    myGrado(0,4) = 0						'Total kilos grado
    myGrado(0,5) = 0						'Porcentaje Total kilos grado
    myGrado(1,0) = "CON" 
    myGrado(1,1) = "Total conforme (sin Analisis)"
    myGrado(1,2) = 0 
    myGrado(1,3) = 0 
    myGrado(1,4) = 0 
    myGrado(1,5) = 0 
    myGrado(2,0) = "1" 
    myGrado(2,1) = "Total Grado 1"
    myGrado(2,2) = 0 
    myGrado(2,3) = 0 
    myGrado(2,4) = 0 
    myGrado(2,5) = 0     
    myGrado(3,0) = "2" 
    myGrado(3,1) = "Total Grado 2"
    myGrado(3,2) = 0 
    myGrado(3,3) = 0 
    myGrado(3,4) = 0 
    myGrado(3,5) = 0     
    myGrado(4,0) = "3" 
    myGrado(4,1) = "Total Grado 3"
    myGrado(4,2) = 0 
    myGrado(4,3) = 0 
    myGrado(4,4) = 0 
    myGrado(4,5) = 0     
    myGrado(5,0) = "FE" 
    myGrado(5,1) = "Total Fuera Estanda"
    myGrado(5,2) = 0 
    myGrado(5,3) = 0 
    myGrado(5,4) = 0 
    myGrado(5,5) = 0     
    myGrado(6,0) = "F1" 
    myGrado(6,1) = "Total F/E Liquida Gr. 1"
    myGrado(6,2) = 0 
    myGrado(6,3) = 0 
    myGrado(6,4) = 0 
    myGrado(6,5) = 0   
    myGrado(7,0) = "F2" 
    myGrado(7,1) = "Total F/E Liquida Gr. 2"
    myGrado(7,2) = 0 
    myGrado(7,3) = 0 
    myGrado(7,4) = 0 
    myGrado(7,5) = 0   
    myGrado(8,0) = "F3" 
    myGrado(8,1) = "Total F/E Liquida Gr. 3"
    myGrado(8,2) = 0 
    myGrado(8,3) = 0 
    myGrado(8,4) = 0 
    myGrado(8,5) = 0   
    myGrado(9,0) = "F3" 
    myGrado(9,1) = "Total F/E Liquida Gr. 3"
    myGrado(9,2) = 0 
    myGrado(9,3) = 0 
    myGrado(9,4) = 0 
    myGrado(9,5) = 0   
    myGrado(10,0) = "DBT" 
    myGrado(10,1) = "Total Dentro base y/o tolerancia"
    myGrado(10,2) = 0 
    myGrado(10,3) = 0 
    myGrado(10,4) = 0 
    myGrado(10,5) = 0   
    myGrado(11,0) = "FBT" 
    myGrado(11,1) = "Total Fuera base y/o tolerancia"
    myGrado(11,2) = 0 
    myGrado(11,3) = 0 
    myGrado(11,4) = 0 
    myGrado(11,5) = 0   
    myGrado(12,0) = "SAC" 
    myGrado(12,1) = "Total Sin Analisis Cond. Camara"
    myGrado(12,2) = 0 
    myGrado(12,3) = 0 
    myGrado(12,4) = 0 
    myGrado(12,5) = 0   
    myGrado(13,0) = "SAO" 
    myGrado(13,1) = "Total Sin Analisis(Otros)"
    myGrado(13,2) = 0 
    myGrado(13,3) = 0 
    myGrado(13,4) = 0 
    myGrado(13,5) = 0   
End Sub
'--------------------------------------------------------------------------------------------------------------------------------------------
Sub SumarPorcentajeResumen(pTotalVagones, pTotalKilosNetos, byref pTotalVagonesRegistrados, byref pTotalKilosNetosRegistrados)
dim i
For i = 0 To 13
    If pTotalVagones <> 0 Then myGrado(i,3) = GF_EDIT_DECIMALS(((myGrado(i,2) / pTotalVagones) * 100)*100,2)
    If pTotalKilosNetos <> 0 Then myGrado(i,5) = GF_EDIT_DECIMALS(((myGrado(i,4) / pTotalKilosNetos) * 100)*100,2)
    'Response.Write "<hr>1(" & pTotalVagones & ")2(" & pTotalKilosNetos & ")"
    pTotalVagonesRegistrados = pTotalVagonesRegistrados + myGrado(i,2)
    pTotalKilosNetosRegistrados = pTotalKilosNetosRegistrados + myGrado(i,4)
Next 
end Sub          
'---------------------------------------------------------------------------------------------------------------------
Function getVagonesByOperativos(pCdOperativo,pDtContable,pNucartaPorte,pPto)
	Dim strSQL, strWhere
	if (pCdOperativo <> "") Then strWhere = strWhere & " AND VA.CDOPERATIVO = '" & pCdOperativo &"'"
strSQL ="	select A.cdvagon,(YEAR(A.dtcontablevagon)*10000 + Month(A.dtcontablevagon)*100 + DAY(A.dtcontablevagon)) AS dtcontablevagon, "&_
        "           CAST(A.FECHAPESADA as BIGINT)*1000000 + right('000000' + cast(A.HORAPESADA AS varchar(6)), 6) AS DTPESADA,"&_
		"	        Case when (A.Bruto is not null) then A.bruto else 0 end as bruto, "&_
		"	        Case when (A.tara is not null) then A.tara else 0 end as tara, "&_		
        "           CASE A.NUCARTAPORTESERIE WHEN '0000' THEN A.NUCARTAPORTE ELSE A.NUCARTAPORTESERIE + A.NUCARTAPORTE END CartaPorte, " &_
		"           CASE A.cdOperativoSERIE WHEN '0000' THEN A.cdOperativo ELSE A.cdOperativoSerie + A.cdOperativo END Operativo, " &_		
		"           CASE WHEN A.nurecibo is not null THEN A.nurecibo ELSE 0 END recibo " &_
		"	from(( "&_
		"	    SELECT va.dtcontablevagon,va.cdvagon,va.nucartaporteserie,va.nucartaporte,va.cdoperativoserie,va.cdoperativo, " &_
        "               ((Year(HPV.DTPESADA) * 10000) + (Month(HPV.DTPESADA) * 100) + Day(HPV.DTPESADA)) AS FECHAPESADA, "&_
		"		        ((DATEPART(HOUR, HPV.DTPESADA) * 10000) + (DATEPART(MINUTE, HPV.DTPESADA) * 100) + DATEPART(SECOND, HPV.DTPESADA)) AS HORAPESADA, "&_
		"	            va.nurecibo, "&_
		"	            (select pv.vlPesada from dbo.HPesadasVagon pv where pv.dtcontable = va.dtcontable and pv.cdOperativo = va.cdOperativo and pv.cdOperativoSERIE = va.cdOperativoSERIE "&_
		"	            and pv.nuCartaPorte = va.nuCartaPOrte and pv.nuCartaPorteSERIE = va.nuCartaPOrteSERIE and pv.cdVagon = va.cdVagon "&_ 
		"	            and pv.cdPesada = 1 and pv.sqpesada = (Select Max(sqPesada) from dbo.HPesadasVagon "&_
		"	            where pv.dtcontable = dtcontable and pv.cdOperativo = cdOperativo and pv.nuCartaporte = nuCartaPorte and pv.cdOperativoSERIE = cdOperativoSERIE and pv.nuCartaporteSERIE = nuCartaPorteSERIE  "&_
		"		        and pv.cdVagon = cdVagon and cdPesada = 1)) Bruto, "&_
		"		        (select pv.vlPesada from dbo.HPesadasVagon pv where pv.dtcontable = va.dtcontable and pv.cdOperativo = va.cdOperativo and pv.cdOperativoSERIE = va.cdOperativoSERIE "&_
		"		        and pv.nuCartaPorte = va.nuCartaPorte and pv.nuCartaPorteSERIE = va.nuCartaPorteSERIE and pv.cdVagon = va.cdVagon "&_ 
		"		        and pv.cdPesada = 2 and pv.sqpesada = (Select Max(sqPesada) from dbo.HPesadasVagon "&_
		"		        where pv.dtcontable = dtcontable and pv.cdOperativo = cdOperativo and pv.nuCartaporte = nuCartaPorte and pv.cdOperativoSERIE = cdOperativoSERIE and pv.nuCartaporteSERIE = nuCartaPorteSERIE  "&_
		"		        and pv.cdVagon = cdVagon and cdPesada = 2)) Tara, "&_
		"		        (select mv.vlMermaKilos from dbo.HMermasVagones mv where mv.dtcontable = va.dtcontable and mv.cdOperativo = va.cdOperativo and mv.cdOperativoSERIE = va.cdOperativoSERIE "&_
		"		        and mv.nuCartaPorte = va.nuCartaPOrte and mv.nuCartaPorteSERIE = va.nuCartaPOrteSERIE and mv.cdVagon = va.cdVagon "&_ 
		"		        and mv.sqpesada = (Select Max(sqPesada) from dbo.HPesadasVagon "&_
		"		        where mv.dtcontable = dtcontable and mv.cdOperativo = cdOperativo and mv.nuCartaporte = nuCartaPorte and mv.cdOperativoSERIE = cdOperativoSERIE and mv.nuCartaporteSERIE = nuCartaPorteSERIE  "&_
		"		        and mv.cdVagon = cdVagon and cdPesada = 2)) Merma "&_
		"	FROM   dbo.hvagones va "&_
		"	Left Join 	(select * from dbo.HPesadasVagon pv "&_
        "				 where sqPesada = (Select Max(sqPesada) from dbo.HPesadasVagon "&_
        "                         where cdOperativo = pv.cdOperativo and NUCARTAPORTE = pv.NUCARTAPORTE and dtContable = pv.dtContable and cdvagon = pv.cdvagon) "&_
        "				) HPV on HPV.cdOperativo = va.cdOperativo and HPV.NUCARTAPORTE = va.NUCARTAPORTE and HPV.dtContable = va.dtContable and HPV.cdvagon = va.cdvagon "&_
		"	WHERE  va.cdestado = 8 "& strWhere
		if (pDtContable <> "") Then strSQL = strSQL & " AND VA.DTCONTABLE = '" & GF_FN2DTCONTABLE(pDtContable) & "'"
strSQL=strSQL&"	)union "&_ 
		"	(SELECT va.dtcontablevagon, va.cdvagon, va.nucartaporteserie,va.nucartaporte,va.cdoperativoserie,va.cdoperativo, " &_
		"          ((Year(HPV.DTPESADA) * 10000) + (Month(HPV.DTPESADA) * 100) + Day(HPV.DTPESADA)) AS FECHAPESADA, "&_
		"		   ((DATEPART(HOUR, HPV.DTPESADA) * 10000) + (DATEPART(MINUTE, HPV.DTPESADA) * 100) + DATEPART(SECOND, HPV.DTPESADA)) AS HORAPESADA, "&_
		"	       va.nurecibo, "&_
		"	       (select pv.vlPesada from dbo.PesadasVagon pv where pv.cdOperativo = va.cdOperativo and pv.cdOperativoSERIE = va.cdOperativoSERIE  "&_
		"	          and pv.nuCartaPorte = va.nuCartaPOrte and pv.nuCartaPorteSERIE = va.nuCartaPOrteSERIE and pv.cdVagon = va.cdVagon  "&_ 
		"	          and pv.cdPesada = 1 and pv.sqpesada = (Select Max(sqPesada) from dbo.PesadasVagon  "&_
		"	         where pv.cdOperativo = cdOperativo and pv.nuCartaporte = nuCartaPorte and pv.cdOperativoSERIE = cdOperativoSERIE and pv.nuCartaporteSERIE = nuCartaPorteSERIE   "&_
		"		   and pv.cdVagon = cdVagon and cdPesada = 1)) Bruto,  "&_
		"		    (select pv.vlPesada from dbo.PesadasVagon pv where pv.cdOperativo = va.cdOperativo and pv.cdOperativoSERIE = va.cdOperativoSERIE  "&_
		"		    and pv.nuCartaPorte = va.nuCartaPorte and pv.nuCartaPorteSERIE = va.nuCartaPorteSERIE and pv.cdVagon = va.cdVagon  "&_ 
		"		    and pv.cdPesada = 2 and pv.sqpesada = (Select Max(sqPesada) from dbo.PesadasVagon  "&_
		"		    where pv.cdOperativo = cdOperativo and pv.nuCartaporte = nuCartaPorte and pv.cdOperativoSERIE = cdOperativoSERIE and pv.nuCartaporteSERIE = nuCartaPorteSERIE   "&_
		"		    and pv.cdVagon = cdVagon and cdPesada = 2)) Tara,  "&_
		"		    (select mv.vlMermaKilos from dbo.HMermasVagones mv where mv.cdOperativo = va.cdOperativo and mv.cdOperativoSERIE = va.cdOperativoSERIE  "&_
		"		    and mv.nuCartaPorte = va.nuCartaPOrte and mv.nuCartaPorteSERIE = va.nuCartaPOrteSERIE and mv.cdVagon = va.cdVagon  "&_ 
		"		    and mv.sqpesada = (Select Max(sqPesada) from dbo.PesadasVagon  "&_
		"		    where  mv.cdOperativo = cdOperativo and mv.nuCartaporte = nuCartaPorte and mv.cdOperativoSERIE = cdOperativoSERIE and mv.nuCartaporteSERIE = nuCartaPorteSERIE   "&_
		"		    and mv.cdVagon = cdVagon and cdPesada = 2)) Merma  "&_
		"	FROM   dbo.vagones va  "&_
		"	Left Join 	(select * from dbo.PesadasVagon pv "&_
        "				 where sqPesada = (Select Max(sqPesada) from dbo.PesadasVagon "&_
        "                         where cdOperativo = pv.cdOperativo and NUCARTAPORTE = pv.NUCARTAPORTE and cdvagon = pv.cdvagon) "&_
        "				) HPV on HPV.cdOperativo = va.cdOperativo and HPV.NUCARTAPORTE = va.NUCARTAPORTE and HPV.cdvagon = va.cdvagon "&_
		"	WHERE  va.cdestado = 8  "& strWhere & "	)) as A   "
	Call GF_BD_Puertos(pPto, rsVag, "OPEN", strSQL)
	'Response.Write strSQL
	'Response.End
	Set getVagonesByOperativos = rsVag
End Function 
'--------------------------------------------------------------------------------------------------------
Function getCaladasVagones(p_dtContable, p_idVagon,p_cdOperativo,p_cartaPorte)
	dim strSql, myDtContable, rs
	
	myDtContable = GF_FN2DTCONTABLE(p_dtContable)
	
	strSql = "SELECT CC.*, AC.*, ACC.dsname, ACC.dslastname "&_
			 "FROM   (SELECT "& Year(Now()) & GF_nDigits(Month(Now()),2) & GF_nDigits(Day(Now()),2) &" DTCONTABLE, "&_
             "				 CV.CDOPERATIVO, "&_
			 "				 CV.NUCARTAPORTE, "&_
             "				 CV.CDVAGON, "&_
             "			     CV.SQCALADA, "&_
             "			     CV.ICTIPOCALADA, "&_
             "			     CV.VLHUMEDAD, "&_
			 "		  	     CV.VLPROTEINA, "&_
			 "		  	     CV.CDACEPTACION, "&_
             "				 CV.CDRUBROPPAL, "&_
             "				 CV.CDGRADO, "&_
             "				 CV.ICCONDICIONFABRICA, "&_
             "				 CV.CDUSERNAME, "&_
             "				 CV.DTCALADA, "&_
             "			     CV.DSOBSERVACIONES, "&_
             "				 CV.NUBARRAS, "&_
             "				 CV.CDMOTIVORECHAZO, "&_
             "				 CV.PCMERMA, "&_
             "				 CV.ICCAMARA, "&_
             "				 CV.NUBALDE, "&_
             "				 CV.CDOPERATIVOSERIE, "&_
             "				 CV.NUCARTAPORTESERIE, "&_
             "				 VA.CDPRODUCTO "&_
			 "		  FROM   caladadevagones CV  "&_			 
			 "		  INNER JOIN vagones VA "&_
             "			ON VA.cdoperativo = CV.cdoperativo AND VA.cdvagon = CV.cdvagon AND VA.nucartaporte = CV.nucartaporte "&_
			 "	 	  WHERE  CV.cdvagon = '"&p_idVagon&"' "&_
			 "	 		AND  CV.cdoperativo = '"&p_cdOperativo&"' "&_
			 "	UNION "&_
			 " SELECT ((YEAR(HCV.DTCONTABLE)*10000) + (MONTH(HCV.DTCONTABLE)*100) + DAY(HCV.DTCONTABLE)) DTCONTABLE, "&_
             "				 HCV.CDOPERATIVO, "&_
			 "				 HCV.NUCARTAPORTE, "&_
             "				 HCV.CDVAGON, "&_
             "			     HCV.SQCALADA, "&_
             "			     HCV.ICTIPOCALADA, "&_
             "			     HCV.VLHUMEDAD, "&_
			 "		  	     HCV.VLPROTEINA, "&_
			 "		  	     HCV.CDACEPTACION, "&_
             "				 HCV.CDRUBROPPAL, "&_
             "				 HCV.CDGRADO, "&_
             "				 HCV.ICCONDICIONFABRICA, "&_
             "				 HCV.CDUSERNAME, "&_
             "				 HCV.DTCALADA, "&_
             "			     HCV.DSOBSERVACIONES, "&_
             "				 HCV.NUBARRAS, "&_
             "				 HCV.CDMOTIVORECHAZO, "&_
             "				 HCV.PCMERMA, "&_
             "				 HCV.ICCAMARA, "&_
             "				 HCV.NUBALDE, "&_
             "				 HCV.CDOPERATIVOSERIE, "&_
             "				 HCV.NUCARTAPORTESERIE, "&_
             "               HVA.CDPRODUCTO  "&_
			 " FROM   HCALADADEVAGONES HCV "&_
			 "	  INNER JOIN hvagones HVA "&_
             "			ON HVA.dtcontable = HCV.dtcontable AND HVA.cdoperativo = HCV.cdoperativo AND HVA.cdvagon = HCV.cdvagon AND HVA.nucartaporte = HCV.nucartaporte "&_
			 " WHERE  HCV.dtcontable = '"&myDtContable&"' "&_
			 "	  AND HCV.cdoperativo = '"&p_cdOperativo&"' "&_
             "	  AND HCV.CDVAGON = '"&p_idVagon&"') CC "&_
			 " INNER JOIN aceptacioncalidad AC "&_
             "	  ON CC.cdaceptacion = AC.cdaceptacion "&_
			 " LEFT JOIN accounts ACC "&_
             "	  ON Upper(CC.cdusername) = Upper(ACC.cdusername) "&_
			 " WHERE  dtcontable = " & p_dtContable &_
			 " ORDER  BY sqcalada DESC "	
	'response.write strSql
	'response.end
	Call GF_BD_Puertos(g_strPuerto, rs, "OPEN", strSql)
	Set getCaladasVagones = rs
End Function
'-----------------------------------------------------------------------------------------------------------
Function getDsEstadoOperativo(pCdOperativo,pPto)
	Dim strSQL ,aux
	strSQL = "SELECT DSESTADO FROM ESTADOSOPERATIVOS WHERE CDESTADO = "&pCdOperativo
	Call GF_BD_Puertos(pPto, rs, "OPEN", strSQL)
	aux = ""
	if not rs.Eof Then aux = Ucase(Trim(rs("DSESTADO")))
	getDsEstadoOperativo = aux
End Function

%>
