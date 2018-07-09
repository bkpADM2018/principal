<%
dim myTurnoDesde, myTurnoHasta, myIncluir, myOperativo, myCdAceptacion, myStickerDesde, myStickerHasta
dim myFecContableD, myFecContableM, myFecContableA, myFecContableH, myFecContableN, myFecContableDesde
dim myFecContableDH, myFecContableMH, myFecContableAH, myFecContableHH, myFecContableNH, myFecContableHasta
dim myCdCoordinador, myDsCoordinador, myCdCoordinado, myDsCoordinado, myCdProducto
dim myCdCorredor, myDsCorredor, myCdVendedor, myDsVendedor, myCdEntregador, myDsEntregador
Dim myGrado(13,6)
'---------------------------------------------------------
sub getParametros()
		myTurnoDesde = GF_PARAMETROS7("turnoDesde", "", 6) 
		call addParam("turnoDesde", myTurnoDesde, params)
		myTurnoHasta = GF_PARAMETROS7("turnoHasta", "", 6) 
		call addParam("turnoHasta", myTurnoHasta, params)
		if myTurnoDesde > myTurnoHasta then setError(TURNO_DESDEHASTA_INCORRECTO)
		
		myIncluir = GF_PARAMETROS7("Incluir", "", 6) 
		'Response.Write "(" & myIncluir & ")"
		call addParam("incluir", myIncluir, params)

		myOperativo = GF_PARAMETROS7("operativo", "", 6) 
		call addParam("operativo", myOperativo, params)

		myCdAceptacion = GF_PARAMETROS7("cdAceptacion", "", 6) 
		'Response.Write "(" & myIncluir & ")"
		call addParam("cdAceptacion", myCdAceptacion, params)

		myStickerDesde = GF_PARAMETROS7("stickerDesde", "", 6) 
		call addParam("stickerDesde", myStickerDesde, params)
		myStickerHasta = GF_PARAMETROS7("stickerHasta", "", 6) 
		call addParam("stickerHasta", myStickerHasta, params)
		if myStickerDesde > myStickerHasta then setError(STICKER_DESDEHASTA_INCORRECTO)

		myFecContableD = GF_PARAMETROS7("fecContableD", "", 6)
		if (myFecContableD = "") then myFecContableD=Day(Now())
		call addParam("fecContableD", myFecContableD, params)
		myFecContableM = GF_PARAMETROS7("fecContableM", "", 6)
		if (myFecContableM = "") then myFecContableM=Month(Now())
		Call addParam("fecContableM", myFecContableM, params)
		myFecContableA = GF_PARAMETROS7("fecContableA", "", 6)
		if (myFecContableA = "") then myFecContableA=Year(Now())
		Call addParam("fecContableA", myFecContableA, params)
		myFecContableH = GF_PARAMETROS7("fecContableH", "", 6)
		if (myFecContableH = "") then myFecContableH="00"
		call addParam("fecContableH", myFecContableH, params)
		myFecContableN = GF_PARAMETROS7("fecContableN", "", 6)
		if (myFecContableN = "") then myFecContableN="00"
		call addParam("fecContableN", myFecContableN, params)		
		
		myFecContableDesde = myFecContableA & "-" & myFecContableM & "-" & myFecContableD & " " & myFecContableH & ":" & myFecContableN & ":00.0"		
		myFecContableDesdeDate = myFecContableA & "-" & myFecContableM & "-" & myFecContableD & " " & myFecContableH & ":" & myFecContableN
		if not isdate(myFecContableDesdeDate) then setError(FECHA_INICIO_INCORRECTA)

		myFecContableDH = GF_PARAMETROS7("fecContableDH", "", 6)
		if (myFecContableDH = "") then myFecContableDH=Day(Now())
		call addParam("fecContableDH", myFecContableDH, params)
		myFecContableMH = GF_PARAMETROS7("fecContableMH", "", 6)
		if (myFecContableMH = "") then myFecContableMH=Month(Now())
		call addParam("fecContableMH", myFecContableMH, params)
		myFecContableAH = GF_PARAMETROS7("fecContableAH", "", 6)
		if (myFecContableAH = "") then myFecContableAH=Year(Now())
		call addParam("fecContableAH", myFecContableAH, params)
		myFecContableHH = GF_PARAMETROS7("fecContableHH", "", 6)
		if (myFecContableHH = "") then myFecContableHH="23"
		call addParam("fecContableHH", myFecContableHH, params)
		myFecContableNH = GF_PARAMETROS7("fecContableNH", "", 6)
		if (myFecContableNH = "") then myFecContableNH="59"
		call addParam("fecContableNH", myFecContableNH, params)
		
		myFecContableHasta = myFecContableAH & "-" & myFecContableMH & "-" &myFecContableDH & " " & myFecContableHH & ":" & myFecContableNH & ":00.0"		
		myFecContableHastaDate = myFecContableAH & "-" & myFecContableMH & "-" &myFecContableDH & " " & myFecContableHH & ":" & myFecContableNH
		if not isdate(myFecContableHastaDate) then 
			setError(FECHA_FIN_INCORRECTA)
		else
			if isdate(myFecContableDesdeDate) then
				if cdate(myFecContableDesdeDate) > cdate(myFecContableHastaDate) then setError(PERIODO_ERRONEO)
			end if	
		end if	
		
		myCdCoordinador = GF_PARAMETROS7("cdCoordinador", "", 6)
		call addParam("cdCoordinador", myCdCoordinador, params)
		myDsCoordinador = GF_PARAMETROS7("dsCoordinador", "", 6)
		call addParam("dsCoordinador", myDsCoordinador, params)

		myCdCoordinado = GF_PARAMETROS7("cdCoordinado", "", 6)
		call addParam("cdCoordinado", myCdCoordinado, params)
		myDsCoordinado = GF_PARAMETROS7("dsCoordinado", "", 6)
		call addParam("dsCoordinado", myDsCoordinado, params)

		myCdProducto = GF_PARAMETROS7("cdProducto", "", 6)
		call addParam("cdProducto", myCdProducto, params)

		myCdCorredor = GF_PARAMETROS7("cdCorredor", "", 6)
		call addParam("cdCorredor", myCdCorredor, params)
		myDsCorredor = GF_PARAMETROS7("dsCorredor", "", 6)
		call addParam("dsCorredor", myDsCorredor, params)

		myCdVendedor = GF_PARAMETROS7("cdVendedor", "", 6)
		call addParam("cdVendedor", myCdVendedor, params)
		myDsVendedor = GF_PARAMETROS7("dsVendedor", "", 6)
		call addParam("dsVendedor", myDsVendedor, params)

		myCdEntregador = GF_PARAMETROS7("cdEntregador", "", 6)
		call addParam("cdEntregador", myCdEntregador, params)
		myDsEntregador = GF_PARAMETROS7("dsEntregador", "", 6)
		call addParam("dsEntregador", myDsEntregador, params)
end sub		
'---------------------------------------------------------
function generarSQL()
dim strSQL
strSQL = ""
	'Generar consulta
	If myTurnoDesde <> "" Then strSqlFiltro = strSqlFiltro & " and o.sqTurno >= " & Trim(myTurnoDesde)
	If myTurnoHasta <> "" Then strSqlFiltro = strSqlFiltro & " and o.sqTurno <= " & Trim(myTurnoHasta)
	If myOperativo <> "" Then 
		strSqlFiltro = strSqlFiltro & " and o.cdOperativo ='" & IIf(Len(Trim(myOperativo)) = 16, Mid(Trim(myOperativo), 5, 12) & "'", Trim(myOperativo) & "'")
		strSqlFiltro = strSqlFiltro & " and o.cdOperativoSerie ='" & IIf(Len(Trim(myOperativo)) = 16, Mid(Trim(myOperativo), 1, 4) & "'", "0000'")
	end if
	If myStickerDesde <> Empty Then strSqlFiltro = strSqlFiltro & " and cld.nuBarras >='" & Trim(myStickerDesde) & "' "
	If myStickerHasta <> Empty Then strSqlFiltro = strSqlFiltro & " and cld.nuBarras <='" & Trim(myStickerHasta) & "' "
	'If myFecContableDesde <> "//" Then strSqlFiltro = strSqlFiltro & " and va.dtContableVagon >= date('" & myFecContableDesde & "') "
	'If myFecContableHasta <> "//" Then strSqlFiltro = strSqlFiltro & " and va.dtContableVagon <= date('" & myFecContableHasta & "') "
	If myFecContableDesde <> "//" Then strSqlFiltro = strSqlFiltro & " and pe.dtPesada >= '" & myFecContableDesde & "' "
	If myFecContableHasta <> "//" Then strSqlFiltro = strSqlFiltro & " and pe.dtPesada <= '" & myFecContableHasta & "' "


	If myCdAceptacion <> "" Then strSqlFiltro = strSqlFiltro & " and cld.cdAceptacion = " & myCdAceptacion

	If myCdCoordinador <> "" Then strSqlFiltro = strSqlFiltro & " and o.cdEmpresa = " & myCdCoordinador
	If myCdCoordinado <> "" Then strSqlFiltro = strSqlFiltro & " and o.cdCliente = " & myCdCoordinado
	If myCdProducto <> "" Then strSqlFiltro = strSqlFiltro & " and o.cdProducto = " & myCdProducto
	If myCdVendedor <> "" Then strSqlFiltro = strSqlFiltro & " and o.cdVendedor = " & CLng(myCdVendedor)
	If myCdCorredor <> "" Then strSqlFiltro = strSqlFiltro & " and o.cdCorredor = " & CLng(myCdCorredor)
	If myCdEntregador <> "" Then strSqlFiltro = strSqlFiltro & " and o.cdEntregador = " & CLng(myCdEntregador)

	iEstadoTara = 8          
	'Response.Write strSqlFiltro
	if len(strSqlFiltro) > 1 then
		strSQL = " Select o.SQTURNO AS Turno,"
		strSQL = strSQL & " (YEAR(va.dtContableVagon )*10000 + Month(va.dtContableVagon )*100 + DAY(va.dtContableVagon )) AS Fecha,"
		strSQL = strSQL & " (CAST((YEAR(pe.dtpesada)*10000 + Month(pe.dtpesada )*100 + DAY(pe.dtpesada)) as BIGINT)*1000000) + cast((DATEPART(HOUR,pe.dtpesada) * 10000 + DATEPART(MINUTE,pe.dtpesada) * 100 + DATEPART(SECOND,pe.dtpesada)) AS Bigint) AS dtPesada,"
        strSQL = strSQL & " CASE WHEN Em.DSEmpresa IS NULL THEN '' ELSE SUBSTRING(Em.DSEmpresa,1,18) END AS Coordinador, "
		strSQL = strSQL & " CASE WHEN cli.DSCliente IS NULL THEN '' ELSE SUBSTRING(cli.DSCliente,1,18) END AS Coordinado, "
		strSQL = strSQL & " CASE WHEN P.DSProducto IS NULL THEN '' ELSE SUBSTRING(P.DSProducto,1,20) END AS Producto, "
        strSQL = strSQL & " CASE WHEN Co.DSCorredor IS NULL THEN '' ELSE SUBSTRING(Co.DSCorredor,1,20) END AS Corredor, "
        strSQL = strSQL & " CASE WHEN En.DSEntregador IS NULL THEN '' ELSE SUBSTRING(En.DSEntregador,1,20) END AS Entregador, "
		strSQL = strSQL & " CASE WHEN Ve.DSVendedor IS NULL THEN '' ELSE SUBSTRING(Ve.DSVendedor,1,20) END AS Vendedor, "
        strSQL = strSQL & " CASE WHEN Ve.DSVendedor IS NULL THEN '' ELSE SUBSTRING(Pr.DSProcedencia,1,20) END AS Localidad, "
		strSQL = strSQL & " va.cdvagon AS NoVagon,"
		strSQL = strSQL & " CASE o.cdOperativoSERIE WHEN '0000' THEN o.cdOperativo ELSE o.cdOperativoSerie +  o.cdOperativo END AS CartaPorte,"
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
		strSQL = strSQL & " 0 AS Netos,"
		strSQL = strSQL & " cld.nuBarras AS Barras,"
		strSQL = strSQL & " cld.cdAceptacion, "
		strSQL = strSQL & " ac.dsaceptacion AS Aceptacion, "
		strSQL = strSQL & " o.cdProducto, "
		strSQL = strSQL & " 1 AS IMPRIMIR "
		strSQL = strSQL & " From dbo.HVagones va "
		strSQL = strSQL & " Left Join dbo.HCaladadeVagones cld on cld.dtContable = va.dtContable "
		strSQL = strSQL & " and cld.cdOperativo = va.cdoperativo and cld.nuCartaPorte = va.nuCartaporte and cld.cdOperativoSERIE = va.cdoperativoSERIE and cld.nuCartaPorteSERIE = va.nuCartaporteSERIE "
		strSQL = strSQL & " and cld.cdVagon = va.cdVagon "
		strSQL = strSQL & " Left Join dbo.HPesadasVagon pe on pe.dtContable = va.dtContable "
		strSQL = strSQL & " and pe.cdOperativo = va.cdoperativo and pe.nuCartaPorte = va.nuCartaporte and pe.cdOperativoSERIE = va.cdoperativoSERIE and pe.nuCartaPorteSERIE = va.nuCartaporteSERIE "
		strSQL = strSQL & " and pe.cdVagon = va.cdVagon "
		strSQL = strSQL & " Left Join dbo.AceptacionCalidad ac on cld.cdaceptacion = ac.cdaceptacion"				
		strSQL = strSQL & " , dbo.HOperativos o  "
		strSQL = strSQL & " Left Join dbo.Productos p on o.cdproducto = p.cdproducto"
		strSQL = strSQL & " Left Join dbo.Empresas em on o.cdEmpresa = em.cdEmpresa"
		strSQL = strSQL & " Left Join dbo.Clientes cli on o.cdCliente = cli.cdCliente"
		strSQL = strSQL & " Left Join dbo.Corredores co on o.cdcorredor = co.cdcorredor"
		strSQL = strSQL & " Left Join dbo.Vendedores ve on o.cdvendedor = ve.cdvendedor"
		strSQL = strSQL & " Left Join dbo.Procedencias Pr On o.CdProcedencia = Pr.CdProcedencia"
		strSQL = strSQL & " Left Join dbo.Entregadores en on o.cdentregador = en.cdentregador"
		strSQL = strSQL & " Where va.cdestado = " & iEstadoTara
		strSQL = strSQL & " and va.dtcontable = o.dtcontable and va.cdOperativo = o.cdOperativo and va.cdOperativoSERIE = o.cdOperativoSERIE "
		strSQL = strSQL & " and va.nuCartaPorte = o.nuCartaPorte and va.nuCartaPorteSERIE = o.nuCartaPorteSERIE "
		strSQL = strSQL & " and pe.sqPesada = (Select Max(SqPesada) from dbo.HPesadasVagon "
		strSQL = strSQL & " where dtcontable = pe.dtcontable and cdOperativo = pe.cdOperativo and cdOperativoSERIE = pe.cdOperativoSERIE "
		strSQL = strSQL & " and nuCartaPorte = pe.nuCartaPorte and nuCartaPorteSERIE = pe.nuCartaPorteSERIE and cdVagon = pe.cdVAgon and cdPesada=2)"
		strSQL = strSQL & " and cld.sqcalada = (Select Max(SqCalada) from dbo.HCaladadeVagones "
		strSQL = strSQL & " where dtcontable = cld.dtcontable and cdOperativo = cld.cdOperativo and cdOperativoSERIE = cld.cdOperativoSERIE "
		strSQL = strSQL & " and nuCartaPorte = cld.nuCartaPorte and nuCartaPorteSERIE = cld.nuCartaPorteSERIE and cdVagon = cld.cdVagon) "
		strSQL = strSQL & strSqlFiltro
		strSQL = strSQL & " Union Select "
		strSQL = strSQL & " o.SQTURNO AS Turno, "
		strSQL = strSQL & " (YEAR(va.dtContableVagon )*10000 + Month(va.dtContableVagon )*100 + DAY(va.dtContableVagon )) AS Fecha,"
		strSQL = strSQL & " (CAST((YEAR(pe.dtpesada)*10000 + Month(pe.dtpesada )*100 + DAY(pe.dtpesada)) as BIGINT)*1000000) + cast((DATEPART(HOUR,pe.dtpesada) * 10000 + DATEPART(MINUTE,pe.dtpesada) * 100 + DATEPART(SECOND,pe.dtpesada)) AS Bigint) AS dtPesada,"
		strSQL = strSQL & " CASE WHEN Em.DSEmpresa IS NULL THEN '' ELSE SUBSTRING(Em.DSEmpresa,1,18) END AS Coordinador, "
		strSQL = strSQL & " CASE WHEN cli.DSCliente IS NULL THEN '' ELSE SUBSTRING(cli.DSCliente,1,18) END AS Coordinado, "
		strSQL = strSQL & " CASE WHEN P.DSProducto IS NULL THEN '' ELSE SUBSTRING(P.DSProducto,1,20) END AS Producto, "
        strSQL = strSQL & " CASE WHEN Co.DSCorredor IS NULL THEN '' ELSE SUBSTRING(Co.DSCorredor,1,20) END AS Corredor, "
        strSQL = strSQL & " CASE WHEN En.DSEntregador IS NULL THEN '' ELSE SUBSTRING(En.DSEntregador,1,20) END AS Entregador, "
		strSQL = strSQL & " CASE WHEN Ve.DSVendedor IS NULL THEN '' ELSE SUBSTRING(Ve.DSVendedor,1,20) END AS Vendedor, "
        strSQL = strSQL & " CASE WHEN Ve.DSVendedor IS NULL THEN '' ELSE SUBSTRING(Pr.DSProcedencia,1,20) END AS Localidad, "
		strSQL = strSQL & " va.cdvagon AS NoVagon,"
		strSQL = strSQL & " CASE o.cdOperativoSERIE WHEN '0000' THEN o.cdOperativo ELSE o.cdOperativoSerie +  o.cdOperativo END AS CartaPorte,"
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
		strSQL = strSQL & " 0 AS Netos,"
		strSQL = strSQL & " cld.nuBarras AS Barras,"
		strSQL = strSQL & " cld.cdAceptacion, "
		strSQL = strSQL & " ac.dsaceptacion AS Aceptacion, "
		strSQL = strSQL & " o.cdProducto, "
		strSQL = strSQL & " 1 AS IMPRIMIR "
		strSQL = strSQL & " From dbo.Vagones va "
		strSQL = strSQL & " Left Join dbo.CaladadeVagones cld on "
		strSQL = strSQL & " cld.cdOperativo = va.cdoperativo and cld.nuCartaPorte = va.nuCartaporte and  cld.cdOperativoSERIE = va.cdoperativoSERIE and cld.nuCartaPorteSERIE = va.nuCartaporteSERIE "
		strSQL = strSQL & " and cld.cdVagon = va.cdVagon "
		strSQL = strSQL & " Left Join dbo.PesadasVagon pe on "
		strSQL = strSQL & " pe.cdOperativo = va.cdoperativo and pe.nuCartaPorte = va.nuCartaporte and pe.cdOperativoSERIE = va.cdoperativoSERIE and pe.nuCartaPorteSERIE = va.nuCartaporteSERIE "
		strSQL = strSQL & " and pe.cdVagon = va.cdVagon "
		strSQL = strSQL & " Left Join dbo.AceptacionCalidad ac on cld.cdaceptacion = ac.cdaceptacion"		
		strSQL = strSQL & " ,dbo.Operativos o  "
		strSQL = strSQL & " Left Join dbo.Productos p on o.cdproducto = p.cdproducto"
		strSQL = strSQL & " Left Join dbo.Empresas em on o.cdEmpresa = em.cdEmpresa"
		strSQL = strSQL & " Left Join dbo.Clientes cli on o.cdCliente = cli.cdCliente"
		strSQL = strSQL & " Left Join dbo.Corredores co on o.cdcorredor = co.cdcorredor"
		strSQL = strSQL & " Left Join dbo.Vendedores ve on o.cdvendedor = ve.cdvendedor"
		strSQL = strSQL & " Left Join dbo.Procedencias Pr On o.CdProcedencia = Pr.CdProcedencia"
		strSQL = strSQL & " Left Join dbo.Entregadores en on o.cdentregador = en.cdentregador"
		strSQL = strSQL & " Where va.cdestado = " & iEstadoTara
		strSQL = strSQL & " and va.cdOperativo = o.cdOperativo and va.cdOperativoSERIE = o.cdOperativoSERIE "
		strSQL = strSQL & " and va.nuCartaPorte = o.nuCartaPorte and va.nuCartaPorteSERIE = o.nuCartaPorteSERIE "
		strSQL = strSQL & " and pe.sqPesada = (Select Max(SqPesada) from dbo.PesadasVagon "
		strSQL = strSQL & " where cdOperativo = pe.cdOperativo and cdOperativoSERIE = pe.cdOperativoSERIE "
		strSQL = strSQL & " and nuCartaPorte = pe.nuCartaPorte and nuCartaPorteSERIE = pe.nuCartaPorteSERIE and cdVagon = pe.cdVAgon and cdPesada=2)"
		strSQL = strSQL & " and cld.sqcalada = (Select Max(SqCalada) from dbo.CaladadeVagones "
		strSQL = strSQL & " where cdOperativo = cld.cdOperativo and cdOperativoSERIE = cld.cdOperativoSERIE "
		strSQL = strSQL & " and nuCartaPorte = cld.nuCartaPorte and nuCartaPorteSERIE = cld.nuCartaPorteSERIE and cdVagon = cld.cdVagon) "
		strSQL = strSQL & strSqlFiltro
		strSQL = strSQL & " Order by 3,17"
		'Response.Write strSQL
	 	'call GF_BD_Puertos(pto, rsGeneral, "OPEN", strSQL)
	end if	
	generarSQL = strSQL
end function	
'---------------------------------------------------------------------------------------------------------------------------
Function FormatFechaSQL(ByVal pFecha)
dim myFecha, myFechaIni
	myFecha = year(pFecha) & "/" & month(pFecha) & "/" & day(pFecha)
	myFechaIni = year(pFecha) & "/01/01"
    'FormatFechaSQL = CStr(Year(myFecha) & DateDiff("d", myFechaIni, myFecha) + 1)
    FormatFechaSQL = Replace(CStr(myFecha), "/", "-")
End Function
'---------------------------------------------------------------------------------------------------------------------------
function VerGrado(pPto, nProducto, nAceptacion, sBarras, dFecha, incluir)
Dim rs, RS2, strSQL, rsensayo, rsGrado, rsresultado, rsCertificado, rsHonorarios, rsTipo, rsBonifreBaja, rsBaseoTolerancia, rsFueraStandard, auxResul, iEstadarBase
sEngra = "00000"
vCores = 1000
sGrado = "FAL"
auxFechaContable = Left(dFecha,4) &"-"& Mid(dFecha,5,2) &"-"& Right(dFecha,2)
If UCASE(pPto) = "PIEDRABUENA" Then
        'Response.Write "<br>nAceptacion" & nAceptacion
    If nAceptacion = 2 then 'Val(GetParametro("CDACEPTACIONCAMARA", "Código de Aceptacion Camara", "2")) Then
        strSQL = "Select cdEnsayo,cdGrado,cdResultado,nuCertificado,vlHonorarios,icTipo,vlBonifreBaja, "
        strSQL = strSQL & " icFueraStandard, icBaseoTolerancia from dbo.ResultadosCamara "
        strSQL = strSQL & " where nuBArras = '" & sBarras & "' "
        strSQL = strSQL & " and dtcontable = cast('" & auxFechaContable & "' as date) "
        strSQL = strSQL & " order by cdEnsayo "
        'Response.Write "<br>" & strSQL
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
        If nAceptacion = 1 then 'Val(GetParametro("CDACEPTACIONCONFORME", "Código de Aceptacion Conforme", "1")) Then
            sGrado = "CON"
        Else
            sGrado = "SAO"
        End If
    End If
    Select Case incluir 'c_orden.Text
        Case "S"
            If sGrado <> "SAC" Or sGrado <> "SAO" Then sGrado = "XXX"
        Case "C"
            If sGrado = "SAC" Or sGrado = "SAO" Then sGrado = "XXX"
    End Select
    'OTROS  /// ARROYO EL TRANSITO
Else
    If sBarras <> Empty Then
        TotEnsCond = 0
        TotEnsCondSi = 0
        TotEnsCondNo = 0
        TotEnsayos = 0
        strSQL = " Select icEstandarBase from Productos where cdproducto = " & nProducto        
        call GF_BD_Puertos(pPto, rs, "OPEN", strSQL)
        iEstandarBase = VerNull(rs(0))
        Set rs = Nothing
            
        strSQL = "Select cdEnsayo,cdGrado,cdResultado,nuCertificado,vlHonorarios,icTipo,vlBonifreBaja, "
        strSQL = strSQL & " icFueraStandard, icBaseoTolerancia from ResultadosCamara "
        strSQL = strSQL & " where nuBArras = '" & sBarras & "' "
        strSQL = strSQL & " and dtcontable = cast('" & auxFechaContable & "' as date)"
        strSQL = strSQL & " order by cdEnsayo "
        call GF_BD_Puertos(pPto, rs, "OPEN", strSQL)
        'Set rs = AdoConnection.Execute(strSQL)
        If rs.EOF Then
            sGrado = "SAC"
        Else
            rsensayo = Trim(VerNull(rs("CdEnsayo")))
            rsGrado = VerNull(rs("CdGrado"))
            rsresultado = VerNull(rs("cdResultado"))
            rsCertificado = VerNull(rs("nuCertificado"))
            rsHonorarios = VerNull(rs("vlHonorarios"))
            rsTipo = VerNull(rs("icTipo"))
            rsBonifreBaja = VerNull(rs("vlBoniFreBaja"))
            rsBaseoTolerancia = VerNull(rs("icBaseoTolerancia"))
            rsFueraStandard = VerNull(rs("icFueraStandard"))
                
            sCertificado = rsCertificado
                
            If iEstandarBase = 1 Then
                strSQL = "Select cdResultado,DsResultado from dbo.ResultadosCamara "
                strSQL = strSQL & " where nuBArras = '" & sBarras & "' "
                strSQL = strSQL & " and dtcontable = cast('" & auxFechaContable & "' as date) "
                strSQL = strSQL & " and cdEnsayo = '" & sEngra & "'"
                call GF_BD_Puertos(pPto, RS2, "OPEN", strSQL)
                'Set RS2 = AdoConnection.Execute(strSQL)
                If RS2.EOF Then
                    sGrado = "FAL"
                Else
                    auxResul = VerNull(RS2(0)) / vCores
                    If Not VerEstandares(nProducto, sBarras, rsensayo, rsresultado) Then
                        Select Case auxResul
                            Case 1
                                sGrado = "1"
                            Case 2
                                sGrado = "2"
                            Case "3"
                                sGrado = "3"
                            Case 4
                                sGrado = "FE"
                        End Select
                    Else
                        Select Case auxResul
                            Case 1
                                sGrado = "F1"
                            Case 2
                                sGrado = "F2"
                            Case "3"
                                sGrado = "F3"
                            Case 4
                                sGrado = "FE"
                        End Select
                    End If
                End If
            Else
                If Not VerEstandares(nProducto, sBarras, rsensayo, rsresultado) Then
                    sGrado = "DBT"
                Else
                    sGrado = "FBT"
                End If
            End If
        End If
    Else
        If nAceptacion = 1 then 'Val(GetParametro("CDACEPTACIONCONFORME", "Código de Aceptacion Conforme", "1")) Then
            sGrado = "CON"
        Else
            sGrado = "SAO"
        End If
    End If
    Select Case incluir 'c_orden.Text
        Case "Sin Análisis"
            If sGrado <> "SAC" And sGrado <> "SAO" Then sGrado = "XXX"
        Case "Con Análisis"
            If sGrado = "SAC" Or sGrado = "SAO" Then sGrado = "XXX"
    End Select
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


%>