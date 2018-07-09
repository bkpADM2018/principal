<%
Const OBRA_ACTIVA = "ACTIVA"
Const OBRA_TERMINADA = "TERMINADA"
Const BUDGET_1ST_QUARTER = "1STQ"
Const BUDGET_2ND_QUARTER = "2NDQ"
Const BUDGET_3RD_QUARTER = "3RDQ"
Const BUDGET_4TH_QUARTER = "4THQ"
Const BUDGET_TOT_QUARTER = "TOTAL"

Const BUDGET_1ST_QUARTER_DATE = "1201"
Const BUDGET_2ND_QUARTER_DATE = "0301"
Const BUDGET_3RD_QUARTER_DATE = "0601"
Const BUDGET_4TH_QUARTER_DATE = "0901"
Const BUDGET_END_OF_PERIOD = "1130"

Const BUDGET_SIN_DETALLES = 0
Const BUDGET_SIN_AREA = 0

Const OBRA_INVERSION = "S"
Const OBRA_MANTENIMIENTO = "N"


Const OBRA_TIPO_INVERSION 	= "I"
Const OBRA_TIPO_MANT_TRIM 	= "M"
Const OBRA_TIPO_MANT_ANUAL 	= "N"
Const OBRA_TIPO_SERVICIO 	= "S"

Const OBRA_GEID = 99999 ' Gasto especial de imputacion directa
Const OBRA_GECD = "GEID"
Const OBRA_GEDS = "GASTO ESPECIAL DE IMPUTACION DIRECTA"
'***************** TIPOS DE FORMULARIO PARA LA PARTIDA PRESUPUESTARIA *****************
Const OBRA_FORM_ANUAL = 1 ' PARA PARTIDAS CON PRESUPUESTO ANUAL
Const OBRA_FORM_SERVICIOGRAL = 2 'PARA PARTIDAS DE SERVICIO GENEREAL
Const OBRA_FORM_TRIM = 3 ' PARA PARTIDAS CON PRESUPUESTO TRIMESTRAL

Const BUDGET_REASIGNACION_ACTIVA = 1
Const BUDGET_REASIGNACION_FINALIZADA = 3
'----------------------------------------------------------------
Function getDescripcionDivision(idDivision)
	Dim strSQL, rs, conn
	getDescripcionDivision = ""
	strSQL="Select * from TBLDIVISIONES where IDDIVISION=" & idDivision
	'Response.Write strsql
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then
		getDescripcionDivision = rs("DSDIVISION")
	end if
End Function
'----------------------------------------------------------------
Function getDivisionObra(pIdObra) 
	Dim strSQL, rs, rtrn
	rtrn = "vacio"
	strSQL = "Select DSDIVISION from TBLDATOSOBRAS DaOb inner join TBLDIVISIONES D on DaOb.IDDIVISION=D.IDDIVISION where DaOb.idObra = " & pIdObra
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if not rs.eof then rtrn = rs("DSDIVISION")
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
getDivisionObra = rtrn 	
End Function
'----------------------------------------------------------------
sub getDivisionObraFull(pIdObra, byref pIdDivision, byref pDsDivision) 
	Dim strSQL, rs, rtrn
	pIdDivision = ""
	pDsDivision = ""
	strSQL = "Select D.IDDIVISION, D.DSDIVISION from TBLDATOSOBRAS DaOb inner join TBLDIVISIONES D on DaOb.IDDIVISION=D.IDDIVISION where DaOb.idObra = " & pIdObra
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if not rs.eof then 
		pIdDivision = rs("IDDIVISION")
		pDsDivision = rs("DSDIVISION")
	end if
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
End sub
'----------------------------------------------------------------
Function getDescripcionObra(pIdObra) 
	Dim strSQL, rs, rtrn
	rtrn = ""
	strSQL = "Select DSOBRA from TBLDATOSOBRAS where idObra = " & pIdObra
	'Response.Write strsql
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if not rs.eof then rtrn = rs("DSOBRA")
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
getDescripcionObra = rtrn 	
End Function
'-----------------------------------------------------------------------------------------------
Function filtrarObrasSeguridad(ByRef myWhere)
	dim myDivisionesAdmin, cdUsuario
	
	'Filtro
	cdUsuario = session("Usuario")
	myDivisionesAdmin = getListaCargosAdmin()
	if (myDivisionesAdmin <> "") then
		myWhere = myWhere & " WHERE ((CDRESPONSABLE='" & cdUsuario & "')"
		myWhere = myWhere & " OR (obras.IDDIVISION IN (0," & myDivisionesAdmin & ")))"
	else
		myWhere = myWhere & " WHERE (CDRESPONSABLE='" & cdUsuario & "')"
	end if
	filtrarObrasSeguridad = myWhere
End Function
'----------------------------------------------------------------
Function filtrarObras(ByRef myWhere, idObra, cdObra, dsObra, idDivision, estadoObra) 
	Dim flagEstado,auxFecha		
	
	myWhere = filtrarObrasSeguridad(myWhere)
	
	if (idObra <> "") then Call mkWhere(myWhere, "IDOBRA", idObra, "=", 1)
	if (cdObra <> "") then Call mkWhere(myWhere, "CDOBRA", cdObra, "=", 3)
	if (dsObra <> "") then Call mkWhere(myWhere, "DSOBRA", dsObra, "LIKE", 3)
	if (estadoObra = OBRA_ACTIVA) then  
		if myWhere = "" then
			myWhere = " Where "
		else
			MyWhere = myWhere & " and"
		end if
			auxFecha = Left(session("MmtoSistema"),8)
			myWhere = myWhere & " ((FECHAINICIO <= " & auxFecha & " and FECHAFIN >= " & auxFecha & " and FECHAAJUSTADA = 0) or (FECHAINICIO <= " & auxFecha & " and FECHAAJUSTADA >= " & auxFecha & "))"
	end if	
	if ((idDivision <> "0") and (idDivision <> "")) then  Call mkWhere(myWhere, "obras.IDDIVISION", idDivision, "=", 1)		
	flagEstado= false	
End Function
'---------------------------------------------------------------------------------------------
sub loadDatosObra(ByRef pIdObra, byref pObraCD, byref pObraDS, byref pObraDivID, byref pObraDivDS, byref pObraImporte, byref pObraFechaBudget, byref pObraMonedaID, byref pObraFechaInicio, byref pObraFechaFin, byref pObraFechaAjustada, byref pObraRespCD, byref pObraRespDS)

dim strSQL, rs, rtrn

strSQL = "Select * from TBLDATOSOBRAS where idObra = " & pIdObra
Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if not rs.eof then 
		pObraDS = rs("DSOBRA")
		pObraCD = rs("CDOBRA")
		pObraDivID = rs("IDDIVISION")
		pObraDivDS = getDescripcionDivision(pObraDivID)		
		pObraImporte = calcularCostoEstimadoObra(MONEDA_DOLAR, pIdObra,0,0)
		pObraMonedaID = MONEDA_DOLAR 
		pObraFechaInicio = rs("FECHAINICIO") 
		pObraFechaFin = rs("FECHAFIN") 
		pObraFechaAjustada = rs("FECHAAJUSTADA")
		pObraFechaBudget = rs("FECHABUDGET")
		pObraRespCD = rs("CDRESPONSABLE")
		pObraRespDS = getUserDescription(pObraRespCD)
	else
		pIdObra = 0
		pObraDS = ""
		pObraCD = ""
		pObraDivID = 0
		pObraDivDS = ""		
		pObraImporte = 0
		pObraMonedaID = 0
		pObraFechaInicio = 0
		pObraFechaFin = 0
		pObraFechaAjustada = 0
		pObraFechaBudget = 0
		pObraRespCD = ""
		pObraRespDS = ""		
	end if		
end sub
'----------------------------------------------------------------
Function obtenerListaObras(idObra, cdObra, dsObra, idDivision, estadoObra) 
	Dim strSQL, rs, myWhere, conn, firstRecord
	'Ajusto Paginacion
	Call filtrarObras(myWhere, idObra, cdObra, dsObra, idDivision, estadoObra) 	
	strSQL = "Select * from TBLDATOSOBRAS as obras" & myWhere & " order by obras.IDDIVISION, FECHAFIN ASC"
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set obtenerListaObras = rs
End Function
'----------------------------------------------------------------
Function obtenerListaObrasOrdenado(idObra, cdObra, dsObra, idDivision, estadoObra,order) 
	Dim strSQL, rs, myWhere, conn, firstRecord
	
	if order = "" then order = " order by obras.IDDIVISION, FECHAFIN ASC"

	'Ajusto Paginacion
	Call filtrarObras(myWhere, idObra, cdObra, dsObra, idDivision, estadoObra) 	
		
	strSQL = "SELECT div.dsdivision, obras.* FROM tbldatosobras AS obras "
	strSQL = strSQL & "LEFT JOIN tbldivisiones AS div "
	strSQL = strSQL & "ON obras.iddivision = div.iddivision"
	strSQL = strSQL & myWhere & " " & order
	
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set obtenerListaObrasOrdenado = rs
End Function
'----------------------------------------------------------------
Function obtenerListaBudgetObra(idObra, idArea, idDetalle) 
	Dim strSQL, rs
	strSQL = "Select * from TBLBUDGETOBRAS where IDOBRA=" & idObra
	if (idArea <> 0) then strSQL = strSQL & " and IDAREA=" & idArea 
	if (idDetalle <> 0) then strSQL = strSQL & " and IDDETALLE=" & idDetalle 
	strSQL = strSQL & " order by IDAREA, IDDETALLE"
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set obtenerListaBudgetObra = rs
End Function

'----------------------------------------------------------------
Function getDescripcionEstadoObras(idEstado)
	Select Case idEstado
		case ESTADO_OBR_PENDIENTE:
			getDescripcionEstadoObras = GF_TRADUCIR("Propuesta")
		case ESTADO_OBR_APROBADA:
			getDescripcionEstadoObras = GF_TRADUCIR("Aprobada")
		case ESTADO_OBR_ACTIVA:
			getDescripcionEstadoObras = GF_TRADUCIR("En Curso") 
		case ESTADO_OBR_TERMINADA:
			getDescripcionEstadoObras = GF_TRADUCIR("Terminada") 
		case ESTADO_OBR_CANCELADA:
			getDescripcionEstadoObras = GF_TRADUCIR("Cancelada") 
	End Select
End Function
'----------------------------------------------------------------
'Totaliza el presupuesto estimado para la obra.
Function calcularCostoEstimadoObra(pIdMoneda, pIdObra, pIdArea, pIdDetalle)
	Dim rs, importe
		
	importe = 0
	if (pIdArea = "") then pIdArea = 0
	if (pIdDetalle = "") then pIdDetalle = 0
    Call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLBUDGETOBRAS_GET_SALDO_BY_IDOBRA", pIdObra & "||" & pIdArea & "||" & pIdDetalle)
	if (not rs.eof) then
	    if (pIdMoneda = MONEDA_DOLAR) then
	        importe = CDbl(rs("TOTALOBRADOLARES"))
        else
            importe = CDbl(rs("TOTALOBRAPESOS"))
        end if        
	end if
	calcularCostoEstimadoObra = importe

End Function
'-------------------------------------------------------------------
'Suma los AFE que se aprobaron para la obra.
'Se permite calcular el total teniendo en cuenta o no, el desvio autorizado para los gatos respecto al valor real del AFE.
Function totalizarAFESObra(idMoneda, idObra, pidArea, pidDetalle, pContarDesvio)
	Dim strSQL, rs, conn, importe, margen
	importe = 0
	
	if (idObra <> 0) then
		'strSQl = " Select sum(IMPORTEPESOS) IMPORTEPESOS, sum(IMPORTEDOLARES) IMPORTEDOLARES "
		strSQl = " Select * "
		strSQL = strSQL & " FROM tbldatosafe"
		strSQL = strSQL & " WHERE idobra   = " & idObra
		strSQL = strSQL & " AND confirmado ='" & AFE_APROBADO &"'"
		if (pidArea <> 0) then 
			strSQL = strSQL & " AND idarea in (0," & pidArea & ")"
			if (pidDetalle <> 0 ) then strSQL = strSQL & " AND iddetalle in (0," & pidDetalle & ")"		
		end if
		Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		'Response.Write strSQL
		while (not rs.eof) 
			if CLNG(rs("IDAREA")) = 0 then
				if (pidArea <> 0) then	
					'Obtener el porcentual de el area-detalle selecionado en la partida y en base a eso determinar 
					'el monto que le corresponde del AFE ya que el AFE es general.
					myPorcentajeEnPartida = getPorcentajeEnPartida(idMoneda, idObra, pidArea, pidDetalle)
					myImportePesos = myImportePesos + round( CDbl(myPorcentajeEnPartida) * CDbl(rs("IMPORTEPESOS")) / 100,0)
					myImporteDolares = myImporteDolares + round( CDbl(myPorcentajeEnPartida) * CDbl(rs("IMPORTEDOLARES")) / 100,0)
				else
					myImportePesos = myImportePesos + cdbl(rs("IMPORTEPESOS"))
					myImporteDolares = myImporteDolares + cdbl(rs("IMPORTEDOLARES"))
				end if
			else
				myImportePesos = myImportePesos + cdbl(rs("IMPORTEPESOS"))
				myImporteDolares = myImporteDolares + cdbl(rs("IMPORTEDOLARES"))
			end if
			rs.movenext
		wend
		if (myImportePesos <> "") then
			if (idMoneda = MONEDA_DOLAR) then 
				importe = cdbl(myImporteDolares)
			else
				importe = cdbl(myImportePesos)				
			end if	
		end if
		if (not IsNumeric(importe)) then importe = 0			
		if (pContarDesvio) then
			'Calculo el margen autorizado de exceso del AFE: Porcentual o absoluto, el menor de ambos (EN DOLARES!!).
			margen = importe*AFE_MAX_DESVIO_PCN
			if (margen > AFE_MAX_DESVIO_ABS) then margen=AFE_MAX_DESVIO_ABS
			importe = importe + margen
		end if	
	end if
	totalizarAFESObra = importe
End Function
'----------------------------------------------------------------
Function getPorcentajeEnPartida(pIdMoneda, pIdObra, pIdArea, pIdDetalle)
	Dim rs, rtrn, myAsignado, presupuestoTotal
		
	rtrn = 0
	if (pIdArea = "") then pIdArea = 0
	if (pIdDetalle = "") then pIdDetalle = 0
	Call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLBUDGETOBRAS_GET_SALDO_BY_IDOBRA", pIdObra & "||" & pIdArea & "||" & pIdDetalle)
	if (not rs.eof) then
	    if (pIdMoneda = MONEDA_DOLAR) then
	        myAsignado = CDbl(rs("TOTALOBRADOLARES"))
        else
            myAsignado = CDbl(rs("TOTALOBRAPESOS"))
        end if        
        presupuestoTotal = calcularCostoEstimadoObra(pIdMoneda, pIdObra, 0 , 0)
	    rtrn = round(CDbl(myAsignado) * 100 /  CDbl(presupuestoTotal),4)
	end if
		
	getPorcentajeEnPartida = rtrn
	
End Function
'----------------------------------------------------------------
Function calcularSaldoObra(pIdMoneda, pIdObra, pIdArea, pIdDetalle)
	Dim rs, importe
		
	importe = 0
	if (pIdArea = "") then pIdArea = 0
	if (pIdDetalle = "") then pIdDetalle = 0
	Call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLBUDGETOBRAS_GET_SALDO_BY_IDOBRA", pIdObra & "||" & pIdArea & "||" & pIdDetalle)
	if (not rs.eof) then
	    if (pIdMoneda = MONEDA_DOLAR) then
	        importe = CDbl(rs("TOTALOBRADOLARES")) - (Cdbl(rs("IMPORTEPICDOLARES")) + Cdbl(rs("IMPORTEVALESDOLARES")))
        else
            importe = CDbl(rs("TOTALOBRAPESOS")) - (Cdbl(rs("IMPORTEPICPESOS")) + Cdbl(rs("IMPORTEVALESPESOS")))
        end if        
	end if
	
	calcularSaldoObra = importe
End Function
'----------------------------------------------------------------
'Contabiliza los gastos  registrados para las obras contando las cotizaciones cargadas y asociadas a ella.
'pSoloAprobados : DEJO DE UTILIZARSE. No se suprime para mantener la interfaz del metodo.
'
'En todos los casos se ignoran los pics asociados a contratos, lo que se hace es directamente es sumar el total del contrato
'al gasto de la obra.
Function calcularGastosObra(pIdMoneda, pIdObra, pIdArea, pIdDetalle, pSoloAprobados)
	
	Dim rs, importe
		
	importe = 0
	if (pIdArea = "") then pIdArea = 0
	if (pIdDetalle = "") then pIdDetalle = 0
	Call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLBUDGETOBRAS_GET_SALDO_BY_IDOBRA", pIdObra & "||" & pIdArea & "||" & pIdDetalle)   
	if (not rs.eof) then
	    if (pIdMoneda = MONEDA_DOLAR) then
	        importe = Cdbl(rs("IMPORTEPICDOLARES")) + Cdbl(rs("IMPORTEVALESDOLARES"))
        else
            importe = Cdbl(rs("IMPORTEPICPESOS")) + Cdbl(rs("IMPORTEVALESPESOS"))	        
        end if        
	end if
	
	calcularGastosObra = importe
	
End Function
'----------------------------------------------------------------
'Contabiliza los gastos  registrados para las obras contando las cotizaciones cargadas y asociadas a ella.
'NO SE LA UTILIZA EN NINGUN LADO
Function calcularGastosTrimestreObra(pIdMoneda, pIdObra, pIdArea, pIdDetalle, pSoloAprobados,pTrimestre)
	Dim strSQL, rs, conn, importe,fechaInicio,fechaFin
	
	importe = 0
	
	select case pTrimestre
		case 0
			fechaInicio = "0101"
			fechaFin    = "0331"
		case 1
			fechaInicio = "0401"
			fechaFin    = "06030"
		case 2
			fechaInicio = "0701"
			fechaFin    = "0930"
		case 3
			fechaInicio = "1001"
			fechaFin    = "1231"
	end select	
	
	importe = 0
	
	strSQL = "Select sum(IMPORTEPESOS) IMPORTEPESOS, sum(IMPORTEDOLARES) IMPORTEDOLARES from "
	strSQL = strSQL & "( ("
	'-- Suma de PICs fuera de contratos.	
	strSQL = strSQL & " Select 1, sum(D.IMPORTEPESOS) IMPORTEPESOS, sum(D.IMPORTEDOLARES) IMPORTEDOLARES "
	strSQL = strSQL & " FROM (Select * from TBLCTZCABECERA WHERE IDOBRA=" & pIdObra 	
	'MODIFICACION: Solo sumo los Pic con contrato igual a cero
	strSQL = strSQL & " and IDCONTRATO = 0 "
	strSQL = strSQL & " and SUBSTRING(MOMENTO,5,4) >= " & fechaInicio
	strSQL = strSQL & " and SUBSTRING(MOMENTO,5,4) <= " & fechaFin
	if (pSoloAprobados) then 
		'Se toman solo los gastos aprobados o ya facturados
		strSQL = strSQL & " and estado in ('" & CTZ_FIRMADA & "', '" & CTZ_FACTURADA & "')"
	else
		'Se toman todos los gastos cargados y no anulados
		strSQL = strSQL & " AND estado <> '" & CTZ_ANULADA & "'"
	end if
	strSQL = strSQL & " ) C "	
	strSQL = strSQL & " inner join TBLCTZDETALLE D on C.IDCOTIZACION = D.IDCOTIZACION "
	if (pIdArea <> 0) then 
		strSQL = strSQL & " WHERE idarea in (" & pidArea & ")"
		if (pIdDetalle <> 0 ) then strSQL = strSQL & " AND iddetalle in (" & pidDetalle & ")"		
	end if
	'-- FIN Suma de PICs fuera de contratos.
	strSQL = strSQL & ") UNION ("
	'-- Suma de Contratos.
	strSQL = strSQL &  " Select 2, sum(IMPORTEPESOS) IMPORTEPESOS, sum(IMPORTEDOLARES) IMPORTEDOLARES "
	strSQL = strSQL & " FROM TBLOBRACONTRATOS  where ESTADO=" & ESTADO_CTC_AUTORIZADO & " and IDOBRA=" & pIdObra 
	if (pIdArea <> 0) then
		strSQL = strSQL & " AND idarea in (" & pidArea & ")"
		if (pIdDetalle <> 0 ) then strSQL = strSQL & " AND iddetalle in (" & pidDetalle & ")"		
	end if
	'-- FIN Suma de Contratos.
	strSQL = strSQL & ") ) GASTOS "
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then 		
		if (rs("IMPORTEPESOS") <> "") then
			importe = cdbl(rs("IMPORTEPESOS"))
			if (pIdMoneda = MONEDA_DOLAR) then importe = cdbl(rs("IMPORTEDOLARES"))		
		end if
	end if	
	
	calcularGastosTrimestreObra = importe
End Function
'---------------------------------------------------------------------------------
Function isBudgetProvisorio(pFechaBudget)
	dim ret
	isBudgetProvisorio = false
	if ((pFechaBudget = "0") or (pFechaBudget = "")) then
		isBudgetProvisorio = true
	end if	
End Function
'---------------------------------------------------------------------------------
'Funcion que determina si se pueden o no modificar los datos del presupuesto de una obra.
Function puedeModificarBudget(pCdResp, pFechaBudget, pIdDivision)	
	puedeModificarBudget = false
	if (isBudgetProvisorio(pFechaBudget) and ((pCdResp = session("Usuario")) or isAdmin(pIdDivision))) then
		puedeModificarBudget = true
	end if
End Function
'---------------------------------------------------------------------------------
'Funcion que determina si se pueden o no modificar los datos del presupuesto de una obra.
Function puedeReasignarBudget(cdResp, pIdDivision)	
	puedeReasignarBudget = false
	if (((cdResp = session("Usuario")) or isAdmin(pIdDivision))) then
		puedeReasignarBudget = true
	end if
End Function
'---------------------------------------------------------------------------------
Function getMmtoFinalizacionObra(pIdObra)
Dim lsobraCD, lsobraDS, lsobraDivID, lsobraDivDS, lsobraImporte, lsobraMonedaID, lsobraFechaInicio, lsobraFechaFin, lsobraFechaAjustada, lsobraRespCD, lsobraRespDS, lsFechaBudget
Call loadDatosObra(pIdObra, lsobraCD, lsobraDS, lsobraDivID, lsobraDivDS, lsobraImporte, lsFechaBudget, lsobraMonedaID, lsobraFechaInicio, lsobraFechaFin, lsobraFechaAjustada, lsobraRespCD, lsobraRespDS)

ret = lsobraFechaFin
if ((isNumeric(lsobraFechaAjustada)) or (vartype(lsobraFechaAjustada) = 14)) then
	if (CLng(lsobraFechaAjustada) <> 0) then ret = lsobraFechaAjustada
end if
getMmtoFinalizacionObra = CLng(ret)

End Function
'---------------------------------------------------------------------------------
Function checkControlObra(idObra)
	Dim rsObra, cdUsuario
	
	checkControlObra = false
	Set rsObra = obtenerListaObras(idObra, "", "", "", "") 
	if (not rsObra.eof) then		
		if (isAdmin(rsObra("IDDIVISION")) or (rsObra("CDRESPONSABLE") = session("Usuario"))) then
			checkControlObra = true
		end if		
	end if
	
End Function
'-----------------------------------------------------------------------------------------------
Function obtenerReasignaciones(pIdObra,pFechaLimite)
		Dim strSQL,conn,rs
 		
		strSQL =          "SELECT   r.idobra             , "
		strSQL = strSQL & "         r.fecha              , "
		strSQL = strSQL & "         r.idareaOrigen     	  idareaOrigen	, "
		strSQL = strSQL & "         r.iddetalleOrigen     iddetaOrigen	, "
		strSQL = strSQL & "         r.idareaDestino   	  idareaDestino	, "
		strSQL = strSQL & "         r.iddetalleDestino    iddetaDestino	, "
		strSQL = strSQL & "         origen.dsbudget		  detaOrigen	, "
		strSQL = strSQL & "         areaOrigen.dsbudget	  dsArea		, "
		strSQL = strSQL & "         destino.dsbudget	  detaDestino	, "
		strSQL = strSQL & "         r.importepesos       , "
		strSQL = strSQL & "         r.importedolares     , "
		strSQL = strSQL & "         r.tipocambio         , "
		strSQL = strSQL & "         r.motivo             , "
		strSQL = strSQL & "         r.cdusuario, "
		strSQL = strSQL & "         r.periodo, "
        strSQL = strSQL & "         r.idReasignacion "
		strSQL = strSQL & "FROM     TBLBUDGETREASIGNACION r "
		strSQL = strSQL & "         INNER JOIN tblbudgetObras origen "
		strSQL = strSQL & "         ON       r.idobra = origen.idobra and r.idareaorigen = origen.idarea and r.iddetalleorigen = origen.iddetalle "
		strSQL = strSQL & "         INNER JOIN tblbudgetObras destino "
		strSQL = strSQL & "         ON       r.idobra = destino.idobra and r.idareaDestino = destino.idarea and r.iddetalledestino = destino.iddetalle "
		strSQL = strSQL & "         INNER JOIN tblbudgetObras areaOrigen "
		strSQL = strSQL & "         ON       r.idobra = areaOrigen.idobra and r.idareaOrigen = areaOrigen.idarea and areaOrigen.idDetalle=0 "
		strSQL = strSQL & "WHERE    r.idobra = " & pIdObra
		if (pFechaLimite <> "") then
			strSQL = strSQL & " AND r.fecha <= " & pFechaLimite
		end if
		strSQL = strSQL & " ORDER BY r.idobra      , "
		strSQL = strSQL & "         r.momento DESC  , "
		strSQL = strSQL & "         idareaorigen, "
		strSQL = strSQL & "         iddetalleorigen"
		
		'Response.Write strSQL
		Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)

		Set obtenerReasignaciones = rs
	End Function
'-------------------------------------------------------------------------------------
'Funcion que determina el importe que se reasigno a posteriori a la fecha indicada.
'El Importe devuelto debe sumarse al budget actual para obtener así el budget a la fecha pedida.
'Explicacion:
'Esto se hace así debido a que cuando se reasigna se modifica el importe original del budget 
'y si queremos el budget a una fecha dada, debemos leer el budget actual y deshacer las reasignaciones realizadas
'a posteriori de la fecha que nos interesa.
'JAS - DEPRECATED (No se usa.)
'Function obtenerDiferenciaReasignacion(pIdObra, pIdArea, pIdDetalle, pFecha, pMoneda)
'	Dim strSQL,rsSuma,rsResta,conn,rtrn
'	Dim suma,resta
'
'	'Calculo todo el presupuesto que se paso a otra 
'	strSQL =          "SELECT	SUM(importepesos)   " & MONEDA_PESO & ", "
'	strSQL = strSQL & "			SUM(importedolares) " & MONEDA_DOLAR & " "
'	strSQL = strSQL & " FROM	TBLBUDGETREASIGNACION r "
'	strSQL = strSQL & " WHERE	idobra			= " & pIdObra
'	strSQL = strSQL & " AND		idAreaOrigen    = " & pIdArea
'	strSQL = strSQL & " AND		iddetalleorigen = " & pIdDetalle
'	strSQL = strSQL & " AND		r.fecha         > " & pFecha
'	Call GF_BD_COMPRAS(rsSuma, conn, "OPEN", strSQL)
'		
'	strSQL =		  "SELECT	SUM(importepesos)   " & MONEDA_PESO & ", "
'	strSQL = strSQL & "			SUM(importedolares) " & MONEDA_DOLAR & " "
'	strSQL = strSQL & " FROM	TBLBUDGETREASIGNACION r "
'	strSQL = strSQL & " WHERE	idobra           = " & pIdObra
'	strSQL = strSQL & " AND		idAreaDestino    = " & pIdArea
'	strSQL = strSQL & " AND		iddetalleDestino = " & pIdDetalle
'	strSQL = strSQL & " AND		r.fecha          > " & pFecha
'	Call GF_BD_COMPRAS(rsResta, conn, "OPEN", strSQL)
'	
'	suma = 0
'	if (not isnull(rsSuma(pMoneda))) then	suma = rsSuma(pMoneda)	
'	
'	resta = 0
'	if (not isnull(rsResta(pMoneda))) then 	resta = rsResta(pMoneda)
'		
'	rtrn = CDbl(suma) - CDbl(resta)
'	
'	obtenerDiferenciaReasignacion = rtrn
'	
'End Function
'-------------------------------------------------------------------------------------
Function leerBudget(pIdObra)
	Dim strSQL, conn, rs
			
	strSQL="Select * from TBLBUDGETOBRAS where IDOBRA=" & pIdObra & " Order by IDAREA, IDDETALLE"
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set leerBudget = rs
End Function
'----------------------------------------------------------------------------------------
'Funcion que determina el tipo de cambio para el presupuesto
'Primero toma el tipo de cambio que se pasa por parametro "tipoCambio"
'si no hay datos Segundo lo busca en la base de datos y si no hay info, toma el tipo de cambio del dia.
Function getTipoCambioBudget(pIdObra)
	Dim ret, rs, strSQL, conn
	
	'Obtiene de los parametros.
	ret = GF_PARAMETROS7("tipoCambio", 3, 6)
	if (ret = 0) then
		'Busca en la base de datos para el presupuesto.
		strSQL="Select MAX(TIPOCAMBIO) TC from TBLBUDGETOBRAS where IDOBRA=" & pIdObra & " and IDDETALLE > 0"
		Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if (not rs.eof) then
			if (rs("TC") <> "") then ret = CDbl(rs("TC"))
		end if
		if (ret = 0) then
			'Se asume que el importe ingresado es en dolares!
			ret = getTipoCambio(MONEDA_DOLAR, "")		
		end if
	end if
	getTipoCambioBudget = ret
End Function
'---------------------------------------------------------------------------------------------
'Funcion que calcula los gastos que fueron facturados para la partida presupeustaria indicada
'entre las fecha indicadas, si se desea el total de lo facturado, pasar los parametros de fecha en blanco.
Function calcularGastosFacturados(pIdObra,pIdArea,pIdDetalle,pFechaInicio,pFechaFin,pMoneda)
	Dim rs,strSQL,conn
	Dim rtrn
	Dim myWhere,agregado
	Dim myFechaInicio
	Dim myFechaFin
	Dim campoImporte
	rtrn = 0
		
	Call mkWhere(myWhere, "cat.tipocategoria", TIPO_CAT_IMPUESTOS	 ,"<>" , SQL_WHERE_STRING)
	Call mkWhere(myWhere, "acd7.IDObra", pIdObra	 ,"=" , SQL_WHERE_INTEGER)
		
	Call mkWhere(myWhere, "acd7.IDArticulo", "(" & ITEM_FONDO_REPARO_ARS & "," & ITEM_FONDO_REPARO_USD & ")"	 ,"NOT IN" , SQL_WHERE_INTEGER)

	if (pFechaInicio <> "") then 
		myFechaInicio = pFechaInicio
		Call mkWhere(myWhere, "acds.fecvto", myFechaInicio,">=", SQL_WHERE_INTEGER)
	end if
	if (pFechaFin <> "") then 		
		myFechaFin    = pFechaFin
		Call mkWhere(myWhere, "acds.fecvto", myFechaFin   ,"<" , SQL_WHERE_INTEGER)
	end if
	'EAB
	'SUMAR SOLO OBRA EN CURSO Y ANTICIPOS
	
	if (CInt(pIdArea) <> BUDGET_SIN_AREA) then  Call mkWhere(myWhere, "acd7.IDArea", pIdArea   ,"=" , SQL_WHERE_INTEGER)
	if (CInt(pIdDetalle) <> BUDGET_SIN_DETALLES) then Call mkWhere(myWhere, "acd7.IDDetalle", pIdDetalle,"=" , SQL_WHERE_INTEGER)	
	
	campoImporte = "ImporteDolares"
	if (pMoneda=MONEDA_PESO) then campoImporte = "ImportePesos"
	
	strSQL = strSQL & "		SELECT sum(" & campoImporte & ") AS gasto"
	strSQL = strSQL & "		FROM VWMEP001C acd7  "
	strSQL = strSQL & "		INNER JOIN VWCOMPROBANTES acds ON acd7.NroInt = acds.nroint "	
	strSQL = strSQL & "		INNER JOIN tblarticulos art on art.idarticulo=acd7.IDArticulo "
	strSQL = strSQL & "		INNER JOIN tblartcategorias cat on art.idcategoria=cat.idcategoria "
	strSQL = strSQL &		myWhere
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	
	if (not rs.eof) then
		if (not isnull(rs("gasto"))) then
				rtrn = rs("gasto")
				rtrn = round((cdbl(rtrn)*100),0)
		end if
	end if
	
	calcularGastosFacturados = rtrn
	
End Function

'---------------------------------------------------------------------------------------------
Function obtenerFechaInicioObra(pIdObra)
	Dim rs,strSQL,conn
	Dim rtrn
	
	rtrn = ""	
	strSQL = "SELECT * FROM tbldatosobras WHERE idobra = " & pIdObra
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then
		rtrn = rs("fechainicio")
	end if
	
	obtenerFechaInicioObra = rtrn
End Function
'---------------------------------------------------------------------------------------------
Function obtenerFechaFinObra(pIdObra)
	Dim rs,strSQL,conn
	Dim rtrn
	
	rtrn = ""	
	strSQL = "SELECT * FROM tbldatosobras WHERE idobra = " & pIdObra
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then
		if (cdbl(rs("fechaajustada"))<>0) then
			rtrn = rs("fechaajustada")
		else
			rtrn = rs("fechafin")
		end if
	end if
	obtenerFechaFinObra = rtrn
End Function
'--------------------------------------------------------------------------------------------
Function isMantenimiento(pIdObra)
	Dim rtrn,strSQL,rs,conn
	
	rtrn = false
	
	strSQL = "SELECT * FROM TBLDATOSOBRAS WHERE IDOBRA = " & pIdObra
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)

	if (not rs.eof) then
		if ((rs("TIPOGASTO") = OBRA_TIPO_MANT_TRIM) or (rs("TIPOGASTO") = OBRA_TIPO_MANT_ANUAL)) then rtrn = true
	end if
	
	isMantenimiento = rtrn
	
End Function
'--------------------------------------------------------------------------------------------
Function isInversion(pIdObra)
	Dim rtrn,strSQL,rs,conn
	
	rtrn = false
	
	strSQL = "SELECT * FROM tbldatosobras WHERE idobra = " & pIdObra
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)

	if (not rs.eof) then
		if (rs("ESINVERSION") = OBRA_INVERSION) then rtrn = true
	end if
	
	isInversion = rtrn
	
End Function
'--------------------------------------------------------------------------------------------------
Function obtenerTrimestreObra(pFechaInicio)
	Dim rtrn,mes
	
	mes = mid(pFechaInicio,5,2)
	
	rtrn = round((mes-2)/3,0)
	
	obtenerTrimestreObra = rtrn
End Function
'--------------------------------------------------------------------------------------------------
Function obtenerIdObrasMantenimiento(pAnio,pIdDivision)
	Dim rs,strSQL,conn,rtrn
	
	rtrn = ""
	strSQL = "select idobra from tbldatosobras where esinversion = 'N' and fechainicio between "&pAnio&"0101 and "&pAnio&"1231 and idDivision = " & pIdDivision 
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)

	while not rs.EoF
		rtrn = rtrn & rs("idobra") & ","
		rs.MoveNext
	wend
	
	rtrn = left(rtrn,len(rtrn)-1)
	
	obtenerIdObrasMantenimiento = rtrn
End Function
'----------------------------------------------------------------------------------------
Function obtenerImporteTrimestres(pidObra,pIdArea,pIdDetalle,pTrimestre,pMoneda)
	Dim strSQL,rs,conn,campoMoneda
	
	campoMoneda = "dlbudget"
	if (pMoneda = MONEDA_PESO) then campoMoneda = "psbudget"
	
	
	strSQL = "select "&campoMoneda&" from tblbudgetobrasdetalle where idobra = " &pidObra& " and idarea = " &pIdArea& " and iddetalle=" & pIdDetalle & " and periodo = "  & pTrimestre
	
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	
	rtrn = 0
	if (not rs.EoF) then rtrn = rs(campoMoneda)

	obtenerImporteTrimestres = rtrn
End Function
'---------------------------------------------------------------------------------------
'Funcion que permite obtener los codigos y descripciones de la obra, el area y el detalle indicados.
'Se pueden pasar area o detalle en sero para obtener solo las descripciones que se quieren. Si se quiere la descripcion de detalle, el area es obligatoria.
'El id de obra siempre es obligatoria.
'Autor: Javier A. Scalisi
'Fecha: 31/10/2011
Function obtenerDescripcionCompletaDetalle(idObra, idArea, idDetalle) 
	Dim strSQL, rs

	strSQL = "Select O.CDOBRA CDOBRA, O.DSOBRA DSOBRA, A.IDAREA IDAREA, A.DSBUDGET DSAREA, D.IDDETALLE IDDETALLE, D.DSBUDGET DSDETALLE  " 
	strSQL = strSQL & " from  (Select * from TBLDATOSOBRAS where IDOBRA=" & idObra & ") O "				
	strSQL = strSQL & " left join TBLBUDGETOBRAS A on O.IDOBRA=A.IDOBRA and A.IDAREA=" & idArea & " and A.IDDETALLE=0"
	strSQL = strSQL & " left join TBLBUDGETOBRAS D on O.IDOBRA=D.IDOBRA and D.IDAREA=" & idArea & " and D.IDDETALLE=" & idDetalle	
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set obtenerDescripcionCompletaDetalle = rs
	
End Function
'--------------------------------------------------------------------------------------------------------------
'Controla si un articulo es apto para una determinada obra
Function esArticuloElegibleObra(p_idArticulo, p_idObra)
    esArticuloElegibleObra = true
    if (p_idObra <> OBRA_GEID) then
        if (p_idArticulo = OBRA_ITEM_SEID) then esArticuloElegibleObra = false
    end if
End Function
'--------------------------------------------------------------------------------------------------------------
'Funcion que devuelve el tipo de formulario para una determinada partida presupuestaria
'Autor: Ajaya Nahuel
'Fecha: 15/12/2015
Function getFormTypeByIdObra(p_idObra)
    Dim strSQL
    strSQL = "SELECT TIPOFORMULARIO FROM TBLTIPOGASTOS "&_
             "WHERE ID IN (SELECT TIPOGASTO FROM TBLDATOSOBRAS WHERE IDOBRA = "& p_idObra &")"
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
    getFormTypeByIdObra = 0
    if (not rs.Eof) then getFormTypeByIdObra = CInt(rs("TIPOFORMULARIO"))
End function
'---------------------------------------------------------------------------------------------------------------
'Confirma el uso de una partida presupuestaria, de esta manera sella los valores de la misma
Function confirmarBudget(pIdObra)
	Dim strSQL, conn, rs
	if (pIdObra <> 0) then
        'Solamente se cierra la partida cuando no se la utilizo antes
        strSQL= "SELECT FECHABUDGET FROM TBLDATOSOBRAS WHERE IDOBRA=" & pIdObra & " AND FECHABUDGET = 0"
        Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
        if (not rs.Eof) then
            strSQL= "Update TBLDATOSOBRAS set FECHABUDGET=" & session("MmtoDato") & " where IDOBRA=" & pIdObra
	        Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
        end if
    end if
End Function
'---------------------------------------------------------------------------------------------------------------
Function obtenerAjustePartidaPresupuestaria(pIdObra,pFechaLimite)
		Dim strSQL,conn,rs
 		
		strSQL =          "SELECT   r.idobra             , "
		strSQL = strSQL & "         r.fecha              , "
		strSQL = strSQL & "         0   idareaOrigen	, "
		strSQL = strSQL & "         0   iddetaOrigen	, "
		strSQL = strSQL & "         r.idareaDestino   	  idareaDestino	, "
		strSQL = strSQL & "         r.iddetalleDestino    iddetaDestino	, "
		strSQL = strSQL & "         ''  detaOrigen	, "
		strSQL = strSQL & "         areaOrigen.dsbudget	  dsArea		, "
		strSQL = strSQL & "         destino.dsbudget	  detaDestino	, "
		strSQL = strSQL & "         r.importepesos       , "
		strSQL = strSQL & "         r.importedolares     , "
		strSQL = strSQL & "         r.tipocambio         , "
		strSQL = strSQL & "         r.motivo             , "
		strSQL = strSQL & "         r.cdusuario, "
		strSQL = strSQL & "         r.periodo, "
        strSQL = strSQL & "         r.idReasignacion "
		strSQL = strSQL & "FROM     TBLBUDGETREASIGNACION r "
		strSQL = strSQL & "         INNER JOIN tblbudgetObras destino "
		strSQL = strSQL & "         ON  r.idobra = destino.idobra and r.idareaDestino = destino.idarea and r.iddetalledestino = destino.iddetalle "
        strSQL = strSQL & "             AND r.idareaOrigen = 0 and r.iddetalleOrigen = 0 "
		strSQL = strSQL & "         INNER JOIN tblbudgetObras areaOrigen "
		strSQL = strSQL & "         ON       r.idobra = areaOrigen.idobra and r.idareaDestino = areaOrigen.idarea and areaOrigen.idDetalle=0 "
		strSQL = strSQL & "WHERE    r.idobra = " & pIdObra
		if (pFechaLimite <> "") then strSQL = strSQL & " AND r.fecha <= " & pFechaLimite
		strSQL = strSQL & " ORDER BY r.idobra,r.momento DESC,idareaorigen,iddetalleorigen"
		
		Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)

		Set obtenerAjustePartidaPresupuestaria = rs
End Function
'---------------------------------------------------------------------------------------------------------------
'Esta funcion se encarga de controlar si se puede usar un partida presupuestaria, esto lo hace comprobando que no tenga reasignacion o ajuste activos.
'Devuelve : verdadero en caso que la partida este activa(que no se pueda utilizar por el momento hasta que no se confirme el ajuste o la reasignacion)
'           falso en caso que la partida NO este activa (se pueda utilizar normalmente)
Function tieneReasignacionAjusteActivo(pIdObra, pIdArea, pIdDetalle)
    Dim strSQL, rtrn
    rtrn = false
    'Busco tanto de la partida origen como la destino
    strSQL = "SELECT ESTADO FROM TBLBUDGETREASIGNACION "&_
             "WHERE IDOBRA = "& pIdObra
    if ((Cdbl(pIdArea) <> 0)and(Cdbl(pIdDetalle) <> 0)) then
        strSQL = strSQL & " AND ( ( IDAREAORIGEN = "& pIdArea &" AND IDDETALLEORIGEN = "& pIdDetalle &")" &_
                          "     OR( IDAREADESTINO = "& pIdArea &" AND IDDETALLEDESTINO = "& pIdDetalle &") )"
    end if
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
    while (not rs.Eof)
        if (Cdbl(rs("ESTADO")) <> BUDGET_REASIGNACION_FINALIZADA) then rtrn = true
        rs.MoveNext()
    wend
    tieneReasignacionAjusteActivo = rtrn
End function
%>
