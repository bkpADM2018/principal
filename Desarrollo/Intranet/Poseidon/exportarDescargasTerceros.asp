<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientosFacturacionCalidad.asp"-->
<!--#include file="../Includes/procedimientosLog.asp"-->
<!--#include file="../Includes/procedimientosMail.asp"-->
<!--#include file="../Includes/procedimientosSeguridad.asp"-->
<!--#include file="ReportesPosicionTerminal/reportePosicionTerminalPrintExcel.asp"-->

<%

Function obternerDescargasTerceros(pFechaDesde,pFechaHasta,pCliente,pPto)
	Dim strSQL,myWhere, myWhereVagon
	myWhere = " where ca.CDESTADO in (6, 8) "
	myWhereVagon = " where ca.CDESTADO in (6, 8) "
	if (pFechaDesde <> "") then  
	    Call mkWhere(myWhere, "ca.DtContable", pFechaDesde, ">=", 3)
	    Call mkWhere(myWhereVagon, "ca.DtContableVagon", pFechaDesde, ">=", 3)
	end if
	if (pFechaHasta <> "") then  
	    Call mkWhere(myWhere, "ca.DtContable", pFechaHasta, "<=", 3)
	    Call mkWhere(myWhereVagon, "ca.DtContableVagon", pFechaHasta, "<=", 3)
    end if	    
	if (pCliente <> "")then	 
	    Call mkWhere(myWhere, "cd.cdCliente", pCliente, "=", 1)
	    Call mkWhere(myWhereVagon, "op.cdCliente", pCliente, "=", 1)
	end if
	if (pIdCamion <> "") then	Call mkWhere(myWhere, "ca.idcamion", pIdCamion, "=", 3)    	
	strSQL = "Select T.*, "&_			 			 
			 "       case when cor.NUCUIT is null then '' else cor.NUCUIT end as cuitcorredor, "&_
			 "       case when v.NUDOCUMENTO is null then '' else v.NUDOCUMENTO end as cuitVendedor, "&_
			 "       case when mv.VLMERMAKILOS is null then 0 else mv.VLMERMAKILOS end as mermaVolatil, "&_
             "       do.NCCARTAPORTE as cartaPorte, " &_
             "       do.NCTIPOTRANSPORTE as tipoTransporte, " &_
             "       CASE WHEN do.NCTIPOTRANSPORTE=2 then T.IDTRANSPORTE else do.cPatente end cPatente, "&_                          
             "       do.cPatenteAcoplado, "&_             
             "       (Year(do.DCARGA)*10000 + Month(do.DCARGA)*100 + Day(do.DCARGA)) fechaCarga, "&_             
             "       do.NQPESOC pesoCarga, "&_                          
             "       do.kmRecorrer, "&_             
             "       do.ncCTG, "&_      
             "       case when do.FiTarifa is Null then 0 else do.FiTarifa end FiTarifa, " &_
             "       case when do.TarifaRef is Null then 0 else do.TarifaRef end TarifaRef , " &_       
             "       do.NCUITREPRESENTANTE cuitEntregador, " &_
             "       case when do.NCUITREMITENTE is null then '' else do.NCUITREMITENTE end as cuitRemitente, "&_             
             "       case when do.CRAZONREMITENTE is null then '' else do.CRAZONREMITENTE end as dsRemitente, "&_             
             "       case when do.NCLOCDEST is null then '' else do.NCLOCDEST end as localidadDestinoONCCA, "&_             
             "       case when do.NCLOCPROCE is null then '' else do.NCLOCPROCE end as localidadProcedenciaONCCA, "&_             
             "       case when do.NCESTABLEDEST is null then '' else do.NCESTABLEDEST end as establecimientoDestino, "&_       
             "       case when do.NCESTABLEPROCE is null then '' else do.NCESTABLEPROCE end as establecimientoProcedencia, "&_       
             "       case when do.NCPROVINO is null then '' else do.NCPROVINO end as provinciaProcedenciaONCCA, "&_                          
             "       do.NCUITCHOFER cuitChofer, " &_
             "       do.CRAZONCHOFER dsChofer, " &_
             "       do.NCUITTRANSPORTISTA cuitTransportista, " &_
             "       do.NCIVATRANSPORTISTA ivaTransportista, " &_
	     	 "       do.NCUITDESTINATARIO cuitDestinatario, " &_
             "       case when do.NCUITC1 is null then '00000000000' else do.NCUITC1 end as cuitIntermediario, "&_
             "       case when do.NCUITC1 is null then '' else do.CRAZONC1 end as dsIntermediario, "&_            
             "       case when do.NCUITC2 is null then '00000000000' else do.NCUITC2 end as cuitRteComercial, "&_
             "       case when do.NCUITC2 is null then '' else do.CRAZONC2 end as dsRteComercial, "&_
             "       case when do.NCAU is null then '' else do.NCAU end as NCAU, "&_                      
             "       (Year(do.DVENCE)*10000 + Month(do.DVENCE)*100 + Day(do.DVENCE)) vtoNCAU, "&_   
			 "		 tra.dstransportista, " &_
			 "		 tra.dsDomicilio calleTransportista, " &_
			 "		 case when (tra.altura is Null or tra.altura = 0) then 1 else tra.altura end calleNroTransportista, " &_
			 "		 case when tra.tipoDomicilio is null then 0 else tra.tipoDomicilio end tdomTransportista, " &_
			 "		 case when tra.codpos  is Null then 0 else tra.codpos end cpostalTransportista " &_
			 " from ( "&_
			 "   select (YEAR(ca.DtContable)*10000 + Month(ca.DtContable)*100 + DAY(ca.DtContable)) DtContable, "&_
			 "          (YEAR(ca.DtContable)*10000 + Month(ca.DtContable)*100 + DAY(ca.DtContable)) DtContableDescarga, "&_
			 "          '0' CDOPERATIVO," &_
			 "          ca.IDCAMION IDTRANSPORTE, " &_
			 "          ca.cdProducto, "&_		
			 "          0 IDX, " &_		
			 "          cd.nuCartaPorte, "&_
			 "          cd.cdVendedor, "&_ 			 
			 "          cd.cdCorredor, "&_
			 "          cd.cdCliente, "&_
			 "          ca.cdTransportista, "&_
			 "          cd.cdCosecha, "&_
			 "         (select case when pc.vlPesada is null then 0 else pc.vlPesada end as vlPesada from HPesadasCamion pc  "&_
			 "          where pc.dtContable = ca.dtContable and pc.Idcamion = ca.Idcamion and pc.cdPesada = 1 "&_
			 "                and pc.sqpesada =  (select max(sqPesada) from HPesadasCamion "&_ 
			 "                                    where dtcontable = pc.DtContable and pc.Idcamion = Idcamion and cdPesada = 1)) as Bruto, "&_
			 "         (select case when pc.vlPesada is null then 0 else pc.vlPesada end as vlPesada from HPesadasCamion pc "&_ 
			 "          where pc.dtContable = ca.dtContable  and pc.Idcamion = ca.Idcamion and pc.cdPesada = 2 "&_
			 "               and pc.sqpesada =  (select max(sqPesada) from HPesadasCamion where dtcontable = pc.DtContable and Idcamion = pc.Idcamion and cdPesada = 2)) as Tara,   "&_
			 "         (select case when mc.vlMermaKilos is null then 0 else mc.vlMermaKilos end as vlMermaKilos from HMermasCamiones mc "&_ 
			 "          where mc.dtContable = ca.dtContable  and mc.Idcamion = ca.Idcamion  "&_
			 "                   and mc.sqpesada =  (select max(sqPesada) from HPesadasCamion where dtcontable = mc.DtContable and Idcamion = mc.Idcamion and cdPesada = 2)) as Merma,   "&_
			 "         (select c.VLHUMEDAD from HCaladadeCamiones c "&_ 
			 "                   where c.dtContable = ca.dtContable and c.Idcamion = ca.Idcamion  "&_
			 "                   and c.sqCalada = (select max(sqcalada) from HCaladadeCamiones where dtcontable = c.DtContable and Idcamion = c.Idcamion )) as Humedad,   "&_
			 "         (select c.VLPROTEINA from HCaladadeCamiones c "&_ 
			 "                   where c.dtContable = ca.dtContable and c.Idcamion = ca.Idcamion  "&_
			 "                   and c.sqCalada = (select max(sqcalada) from HCaladadeCamiones where dtcontable = c.DtContable and Idcamion = c.Idcamion )) as Proteina   "&_
			 "   from HCamiones ca "&_
			 "       inner join HCamionesDescarga Cd on  cd.dtContable = ca.DtContable and cd.Idcamion = ca.idcamion " &_
             "       "&  myWhere
             'Si es piedrabuena agrego la consulta a los vagones
             'IMPORTANTE: Para el código de barras de los Vagones siempre se utiliza el mismo código que para los análisis de camara, el sticker se coloca a mano en la pantalla
             '            y se toma directamente de la tabla caladade vagones. Esto es así ya que en vagones no hay una impresora y el sistema no emite los sobres directamente.
             If (pPto = TERMINAL_PIEDRABUENA) then
             strSQL = strSQL & " UNION "&_
             "     SELECT 0 DtContable, "&_
             "          (YEAR(ca.DtContableVagon)*10000 + Month(ca.DtContableVagon)*100 + DAY(ca.DtContableVagon)) DtContableDescarga, "&_
             "          op.CDOPERATIVO," &_
             "          ca.CDVAGON IDTRANSPORTE, " &_
             "          ca.cdproducto, "&_
	     "		(ROW_NUMBER() over (order by ca.NUCARTAPORTE, ca.cdvagon))  idx, " &_
             "          CONCAT(op.NUCARTAPORTESERIE, SUBSTRING(op.NUCARTAPORTE, 1, 8)) nuCartaPorte, "&_             
             "          op.cdvendedor, "&_
             "          op.cdcorredor, "&_             
             "          op.cdcliente, "&_
			 "          op.cdTransportista, "&_
             "          ca.cdcosecha, "&_
             "          (SELECT CASE WHEN pc.vlpesada IS NULL THEN 0 ELSE pc.vlpesada END AS vlPesada "&_
             "           FROM   PESADASVAGON pc "&_
             "           WHERE  pc.CDOPERATIVO = ca.CDOPERATIVO AND pc.CDVAGON = ca.CDVAGON AND pc.cdpesada = 1 "&_
             "                  AND pc.sqpesada = (SELECT Max(sqpesada) "&_
             "                                     FROM   PESADASVAGON "&_
             "                                     WHERE  pc.CDOPERATIVO = CDOPERATIVO AND pc.CDVAGON = CDVAGON AND cdpesada = 1)) AS Bruto, "&_
             "         (SELECT CASE WHEN pc.vlpesada IS NULL THEN 0 ELSE pc.vlpesada END AS vlPesada "&_
             "           FROM   PESADASVAGON pc "&_
             "           WHERE pc.CDOPERATIVO = ca.CDOPERATIVO AND pc.CDVAGON = ca.CDVAGON AND pc.cdpesada = 2 "&_
             "                  AND pc.sqpesada = (SELECT Max(sqpesada) "&_
             "                                     FROM   PESADASVAGON "&_
             "                                     WHERE  pc.CDOPERATIVO = CDOPERATIVO AND pc.CDVAGON = CDVAGON AND cdpesada = 2)) AS Tara, "&_
             "          (SELECT CASE WHEN mc.vlmermakilos IS NULL THEN 0 ELSE mc.vlmermakilos END AS vlMermaKilos "&_
             "           FROM   MERMASVAGONES mc "&_
             "           WHERE  mc.CDOPERATIVO = ca.CDOPERATIVO AND mc.CDVAGON = ca.CDVAGON "&_
             "                  AND mc.sqpesada = (SELECT Max(sqpesada) "&_
             "                                     FROM   PESADASVAGON "&_
             "                                     WHERE  mc.CDOPERATIVO = CDOPERATIVO AND mc.CDVAGON = CDVAGON AND cdpesada = 2)) AS Merma, "&_
             "          (SELECT c.VLHUMEDAD "&_
             "           FROM   CALADADEVAGONES c "&_
             "           WHERE  c.CDOPERATIVO = ca.CDOPERATIVO AND c.CDVAGON = ca.CDVAGON "&_
             "                  AND c.sqcalada = (SELECT Max(sqcalada) "&_
             "                                    FROM   CALADADEVAGONES "&_
             "                                    WHERE  CDOPERATIVO = c.CDOPERATIVO AND CDVAGON = c.CDVAGON)) AS Humedad, "&_
             "          (SELECT c.VLPROTEINA "&_
             "           FROM   CALADADEVAGONES c "&_
             "           WHERE  c.CDOPERATIVO = ca.CDOPERATIVO AND c.CDVAGON = ca.CDVAGON "&_
             "                  AND c.sqcalada = (SELECT Max(sqcalada) "&_
             "                                    FROM   CALADADEVAGONES "&_
             "                                    WHERE  CDOPERATIVO = c.CDOPERATIVO AND CDVAGON = c.CDVAGON)) AS Proteina "&_
             "   FROM   VAGONES ca "&_
        	 "     INNER JOIN OPERATIVOS op "&_
             "          ON op.nucartaporte = ca.nucartaporte AND op.CDOPERATIVO = ca.CDOPERATIVO and op.CDESTADO not in (" & OPERATIVOS_ESTADO_TERMINADO & ") "&_
             "       " &  myWhereVagon
             strSQL = strSQL & " UNION "&_
             "     SELECT ( Year(ca.dtcontable) * 10000 + Month(ca.dtcontable) * 100 + Day(ca.dtcontable) ) DtContable, "&_
             "          (YEAR(ca.DtContableVagon)*10000 + Month(ca.DtContableVagon)*100 + DAY(ca.DtContableVagon)) DtContableDescarga, "&_
             "          op.CDOPERATIVO," &_
             "          ca.CDVAGON IDTRANSPORTE, " &_
             "          ca.cdproducto, "&_
	     "		(ROW_NUMBER() over (order by CA.NUCARTAPORTE, ca.cdvagon) + 500)  idx, " &_
             "          CONCAT(op.NUCARTAPORTESERIE, SUBSTRING(op.NUCARTAPORTE, 1, 8)) nuCartaPorte, "&_             
             "          op.cdvendedor, " &_
             "          op.cdcorredor, " &_
             "          op.cdcliente, " &_
			 "			op.cdTransportista, " &_
             "          ca.cdcosecha, "&_
             "          (SELECT CASE WHEN pc.vlpesada IS NULL THEN 0 ELSE pc.vlpesada END AS vlPesada "&_
             "           FROM   HPESADASVAGON pc "&_
             "           WHERE  pc.dtcontable = ca.dtcontable AND pc.CDOPERATIVO = ca.CDOPERATIVO AND pc.CDVAGON = ca.CDVAGON AND pc.cdpesada = 1 "&_
             "                  AND pc.sqpesada = (SELECT Max(sqpesada) "&_
             "                                     FROM   HPESADASVAGON "&_
             "                                     WHERE  dtcontable = pc.dtcontable AND pc.CDOPERATIVO = CDOPERATIVO AND pc.CDVAGON = CDVAGON AND cdpesada = 1)) AS Bruto, "&_
             "         (SELECT CASE WHEN pc.vlpesada IS NULL THEN 0 ELSE pc.vlpesada END AS vlPesada "&_
             "           FROM   HPESADASVAGON pc "&_
             "           WHERE  pc.dtcontable = ca.dtcontable AND pc.CDOPERATIVO = ca.CDOPERATIVO AND pc.CDVAGON = ca.CDVAGON AND pc.cdpesada = 2 "&_
             "                  AND pc.sqpesada = (SELECT Max(sqpesada) "&_
             "                                     FROM   HPESADASVAGON "&_
             "                                     WHERE  dtcontable = pc.dtcontable AND pc.CDOPERATIVO = CDOPERATIVO AND pc.CDVAGON = CDVAGON AND cdpesada = 2)) AS Tara, "&_
             "          (SELECT CASE WHEN mc.vlmermakilos IS NULL THEN 0 ELSE mc.vlmermakilos END AS vlMermaKilos "&_
             "           FROM   HMERMASVAGONES mc "&_
             "           WHERE  mc.dtcontable = ca.dtcontable AND mc.CDOPERATIVO = ca.CDOPERATIVO AND mc.CDVAGON = ca.CDVAGON "&_
             "                  AND mc.sqpesada = (SELECT Max(sqpesada) "&_
             "                                     FROM   HPESADASVAGON "&_
             "                                     WHERE  dtcontable = mc.dtcontable AND mc.CDOPERATIVO = CDOPERATIVO AND mc.CDVAGON = CDVAGON AND cdpesada = 2)) AS Merma, "&_             
             "          (SELECT c.VLHUMEDAD "&_
             "           FROM   HCALADADEVAGONES c "&_
             "           WHERE  c.dtcontable = ca.dtcontable AND c.CDOPERATIVO = ca.CDOPERATIVO AND c.CDVAGON = ca.CDVAGON "&_
             "                  AND c.sqcalada = (SELECT Max(sqcalada) "&_
             "                                    FROM   HCALADADEVAGONES "&_
             "                                    WHERE  dtcontable = c.dtcontable AND CDOPERATIVO = c.CDOPERATIVO AND CDVAGON = c.CDVAGON)) AS Humedad, "&_
             "          (SELECT c.VLPROTEINA "&_
             "           FROM   HCALADADEVAGONES c "&_
             "           WHERE  c.dtcontable = ca.dtcontable AND c.CDOPERATIVO = ca.CDOPERATIVO AND c.CDVAGON = ca.CDVAGON "&_
             "                  AND c.sqcalada = (SELECT Max(sqcalada) "&_
             "                                    FROM   HCALADADEVAGONES "&_
             "                                    WHERE  dtcontable = c.dtcontable AND CDOPERATIVO = c.CDOPERATIVO AND CDVAGON = c.CDVAGON)) AS Proteina "&_
             "   FROM   HVAGONES ca "&_
        	 "     INNER JOIN HOPERATIVOS op "&_
             "          ON op.dtcontable = ca.dtcontable AND op.nucartaporte = ca.nucartaporte AND op.CDOPERATIVO = ca.CDOPERATIVO "&_
             "       " &  myWhereVagon
             end if
             strSQL = strSQL & " )T  "&_
             "		left join DATOSONCCA do on do.NCCARTAPORTE = T.NUCARTAPORTE "&_			 		 
             "		left join MERMAVOLATIL mv on mv.NUCARTAPORTE = T.NUCARTAPORTE and mv.IDTRANSPORTE=T.IDTRANSPORTE"&_			 		 
			 "		left join vendedores v on v.cdvendedor = T.cdvendedor " &_
			 "		left join corredores cor on cor.cdcorredor = T.cdcorredor " &_			 
			 "		left join transportistas tra on tra.cdtransportista = T.cdtransportista " &_			 
             " where T.tara <> 0 " &_
			 "		and T.CDPRODUCTO in (Select CDPROPIO from TBLCONVERSIONES where TIPODATO='" & CONV_KEY_PRODUCTO & "' and NUCUITCLIENTE='" & CUIT_ADM & "')" 
	'response.write strSQL
	Call GF_BD_Puertos(pPto, rs, "OPEN", strSQL)	
	
	Set obternerDescargasTerceros = rs
End Function
'---------------------------------------------------------------------------------------------------------
Function cargarTablaAcondicionamiento(ByRef pDic, pFechaDesde, pFechaHasta)
       
    Dim rs, myClave, myValor, dtDesde, dtHasta
       
    logMig.info("cargarTablaAcondicionamiento - Inicia") 
    
    Set pDic = createObject("Scripting.Dictionary")       
    Call cargarValoresGlobalesFAC(g_strPuerto)  
    
    Call GF_BD_Puertos(g_strPuerto, rs, "OPEN", armarSQLRubrosCamiones("","", 0, pFechaDesde, pFechaHasta, "T", false, false)) 
    while (not rs.eof)
        if (InStr(1, "," & gRubrosHumedad & ",", "," & rs("RUBRO") & ",") > 0) then
            myClave = rs("FECHA") & "_" & Trim(rs("CPORTE")) & "_" & Trim(rs("IDTRANSPORTE")) & "_SECADO"
            myValor = rs("PORCMERMARUBRO")
			if (CDbl(myValor) > 0) then 
				pDic.Add myClave, myValor			
				logMig.info("Clave agregada: " & myClave & ", Valor: " & myValor)
			else
				logMig.info("Clave descartada: " & myClave & ", Valor: " & myValor)
			end if
        end if 
        if (InStr(1, "," & gRubrosZaranda & "," & RUBRO_EXCLUSIVO_ZARANDA & ",", "," & rs("RUBRO") & ",") > 0) then
            myClave = rs("FECHA") & "_" & Trim(rs("CPORTE")) & "_" & Trim(rs("IDTRANSPORTE")) & "_ZARANDA"
            myValor = rs("PORCMERMARUBRO")
            if (CDbl(myValor) > 0) then 
				pDic.Add myClave, myValor			
				logMig.info("Clave agregada: " & myClave & ", Valor: " & myValor)
			else
				logMig.info("Clave descartada: " & myClave & ", Valor: " & myValor)
			end if
        end if
        rs.MoveNext()
    wend    

	'Camiones - Tomo el rubro temperatura (Se hace meram por servicio 06-VENTILAR)
	dtDesde = GF_FN2DTCONTABLE(pFechaDesde)
	dtHasta = GF_FN2DTCONTABLE(pFechaHasta)	
	strSQL = 	"Select FORMAT(HC.DTCONTABLE, 'yyyyMMdd') FECHA, " &_
				"		HCD.NUCARTAPORTE CPORTE, " &_
				"		HC.IDCAMION IDTRANSPORTE, " &_
				"		HRVC.VLMERMA PORCMERMARUBRO " &_
				"	from HCAMIONES HC " &_
				"	inner join HCAMIONESDESCARGA HCD on HC.DTCONTABLE=HCD.DTCONTABLE and HC.IDCAMION=HCD.IDCAMION " &_
				"	inner join HRUBROSVISTEOCAMIONES HRVC on HC.DTCONTABLE=HRVC.DTCONTABLE and HC.IDCAMION=HRVC.IDCAMION and SQCALADA = (Select MAX(SQCALADA) from HRUBROSVISTEOCAMIONES A where A.DTCONTABLE=HRVC.DTCONTABLE and A.IDCAMION=HRVC.IDCAMION) " &_
				"	where 	HC.DTCONTABLE>= '" & dtDesde &"' " &_
				"		and HC.DTCONTABLE <= '" & dtHasta & "' " &_ 
				"		and HC.CDESTADO in (" & CAMIONES_ESTADO_EGRESADOOK & ", " & CAMIONES_ESTADO_PESADOTARA & ") " &_
				"		and HRVC.CDRUBRO=" & gRubroTemperatura &_ 
				" 		and HRVC.VLMERMA > 0"				
	Call GF_BD_Puertos(g_strPuerto, rs, "OPEN", strSQL) 
	while (not rs.eof)
		myClave = rs("FECHA") & "_" & Trim(rs("CPORTE")) & "_" & Trim(rs("IDTRANSPORTE")) & "_TEMPERATURA"		
		myValor = rs("PORCMERMARUBRO")
		pDic.Add myClave, myValor			
		logMig.info("Clave agregada: " & myClave & ", Valor: " & myValor)
		rs.MoveNext()
	wend
	
    Call GF_BD_Puertos(g_strPuerto, rs, "OPEN", armarSQLRubrosVagones("","", 0, pFechaDesde, pFechaHasta, "T", false)) 
    while (not rs.eof)
        if (InStr(1, "," & gRubrosHumedad & ",", "," & rs("RUBRO") & ",") > 0) then
            myClave = rs("FECHA") & "_" & Trim(rs("CPORTE")) & "_" & Trim(rs("IDTRANSPORTE")) & "_SECADO"
            myValor = rs("PORCMERMARUBRO")
            if (CDbl(myValor) > 0) then 
				pDic.Add myClave, myValor			
				logMig.info("Clave agregada: " & myClave & ", Valor: " & myValor)
			else
				logMig.info("Clave descartada: " & myClave & ", Valor: " & myValor)
			end if	    
        end if
        if (InStr(1, "," & gRubrosZaranda & "," & RUBRO_EXCLUSIVO_ZARANDA & ",", "," & rs("RUBRO") & ",") > 0) then
            myClave = rs("FECHA") & "_" & Trim(rs("CPORTE")) & "_" & Trim(rs("IDTRANSPORTE")) & "_ZARANDA"
            myValor = rs("PORCMERMARUBRO")
            if (CDbl(myValor) > 0) then 
				pDic.Add myClave, myValor			
				logMig.info("Clave agregada: " & myClave & ", Valor: " & myValor)
			else
				logMig.info("Clave descartada: " & myClave & ", Valor: " & myValor)
			end if	    
        end if
        rs.MoveNext()
    wend    
    
	'Vagones - Temperatura (Se hace meram por servicio 06-VENTILAR)
 	strSQL = 	" Select * from " &_ 
				"(Select FORMAT(HC.DTCONTABLEVAGON, 'yyyyMMdd') FECHA, " &_
				"		HC.nucartaporteserie + LEFT(HC.nucartaporte, 8) CPORTE, " &_
				"		HC.CDVAGON IDTRANSPORTE, " &_
				"		HRVC.VLMERMA PORCMERMARUBRO " &_
				"	from VAGONES HC " &_
				"	inner join RUBROSVISTEOVAGONES HRVC on HC.CDVAGON=HRVC.CDVAGON and HC.NUCARTAPORTE=HRVC.NUCARTAPORTE and SQCALADA = (Select MAX(SQCALADA) from HRUBROSVISTEOVAGONES A where A.NUCARTAPORTE=HRVC.NUCARTAPORTE and A.CDVAGON=HRVC.CDVAGON) " &_
				"	where 	HC.DTCONTABLEVAGON >= '" & dtDesde &"' " &_
				"		and HC.DTCONTABLEVAGON <= '" & dtHasta & "' " &_ 
				"		and HC.CDESTADO in (" & CAMIONES_ESTADO_EGRESADOOK & ", " & CAMIONES_ESTADO_PESADOTARA & ") " &_
				"		and HRVC.CDRUBRO=" & gRubroTemperatura &_ 
				" 		and HRVC.VLMERMA > 0" &_			
				"UNION " &_
				"Select FORMAT(HC.DTCONTABLEVAGON, 'yyyyMMdd') FECHA, " &_
				"		HC.nucartaporteserie + LEFT(HC.nucartaporte, 8) CPORTE, " &_
				"		HC.CDVAGON IDTRANSPORTE, " &_
				"		HRVC.VLMERMA PORCMERMARUBRO " &_
				"	from HVAGONES HC " &_
				"	inner join HRUBROSVISTEOVAGONES HRVC on HC.DTCONTABLE=HRVC.DTCONTABLE and HC.CDVAGON=HRVC.CDVAGON and HC.NUCARTAPORTE=HRVC.NUCARTAPORTE and SQCALADA = (Select MAX(SQCALADA) from HRUBROSVISTEOVAGONES A where A.DTCONTABLE=HRVC.DTCONTABLE and A.NUCARTAPORTE=HRVC.NUCARTAPORTE and A.CDVAGON=HRVC.CDVAGON) " &_
				"	where 	HC.DTCONTABLEVAGON >= '" & dtDesde &"' " &_
				"		and HC.DTCONTABLEVAGON <= '" & dtHasta & "' " &_ 
				"		and HC.CDESTADO in (" & CAMIONES_ESTADO_EGRESADOOK & ", " & CAMIONES_ESTADO_PESADOTARA & ") " &_
				"		and HRVC.CDRUBRO=" & gRubroTemperatura &_ 
				" 		and HRVC.VLMERMA > 0) T"				
	Call GF_BD_Puertos(g_strPuerto, rs, "OPEN", strSQL) 
	while (not rs.eof)
		myClave = rs("FECHA") & "_" & Trim(rs("CPORTE")) & "_" & Trim(rs("IDTRANSPORTE")) & "_TEMPERATURA"
		myValor = rs("PORCMERMARUBRO")
		pDic.Add myClave, myValor			
		logMig.info("Clave agregada: " & myClave & ", Valor: " & myValor)
		rs.MoveNext()
	wend
	
    logMig.info("cargarTablaAcondicionamiento - Fin")
    
End Function
'---------------------------------------------------------------------------------------------------------
Function validar(rs, pNroPuerto, pValidar, ByRef pErrMsg)
	Dim ret, myMsg, tDSTransporte
	
	logMig.info("--- Validando cartas de porte ---")
	if (pValidar) then
		ret = false
		pErrMsg = ""
		while (not rs.eof)
			'Verifico que este cartagos los datos de ONCCA			
			if (isNull(rs("cartaPorte"))) then
				pErrMsg = pErrMsg & "Carta de Porte: " & rs("nuCartaPorte") & " - " & rs("IDTRANSPORTE") & " - Faltan cargar datos ONCCA.<br>"
			else			
				tDSTransporte = "Camion"
				if (CInt(rs("TIPOTRANSPORTE")) = TIPO_TRANSPORTE_VAGON) then tDSTransporte = "Vagon"
				' Controles comunes para camiones y vagoens
				if ((Len(TRIM(rs("NCAU"))) <> 14) or (not isNumeric(rs("NCAU")))) then pErrMsg = pErrMsg & "Carta de Porte: " & rs("nuCartaPorte") & " - " & tDSTransporte & ": " & rs("IDTRANSPORTE") & " - Numero de CAU/CEE erroneo (Debe ser numerico y tener 14 caracteres).<br>"
				if (abs(GF_DTEDIFF(rs("fechaCarga"), rs("DTContableDescarga"), "D")) > 15) then pErrMsg = pErrMsg & "Carta de Porte: " & rs("nuCartaPorte") & " - " & tDSTransporte & ": " & rs("IDTRANSPORTE") & " - Error en fecha de carga, no puede ser más de 15 días previos a la descarga.<br>"				
				Call convertirDatoPuerto(CONV_KEY_PUERTO, pNroPuerto, myMsg)
				if (myMsg <> "") then pErrMsg = pErrMsg & "Carta de Porte: " & rs("nuCartaPorte") & " - " & tDSTransporte & ": " & rs("IDTRANSPORTE") & " - " & myMsg
				Call convertirDatoPuerto(CONV_KEY_PLANTA, pNroPuerto, myMsg)
				if (myMsg <> "") then pErrMsg = pErrMsg & "Carta de Porte: " & rs("nuCartaPorte") & " - " & tDSTransporte & ": " & rs("IDTRANSPORTE") & " - " & myMsg
				Call convertirDatoPuerto(CONV_KEY_PRODUCTO, rs("cdProducto"), myMsg)
				if (myMsg <> "") then pErrMsg = pErrMsg & "Carta de Porte: " & rs("nuCartaPorte") & " - " & tDSTransporte & ": " & rs("IDTRANSPORTE") & " - " & myMsg
				Call convertirDatoPuerto(CONV_KEY_SERVICIO, "SECADO", myMsg)
				if (myMsg <> "") then pErrMsg = pErrMsg & "Carta de Porte: " & rs("nuCartaPorte") & " - " & tDSTransporte & ": " & rs("IDTRANSPORTE") & " - " & myMsg
				Call convertirDatoPuerto(CONV_KEY_SERVICIO, "ZARANDEO", myMsg)
				if (myMsg <> "") then pErrMsg = pErrMsg & "Carta de Porte: " & rs("nuCartaPorte") & " - " & tDSTransporte & ": " & rs("IDTRANSPORTE") & " - " & myMsg			
				'Controles exclusivos apra cada transporte
				if (CInt(rs("TIPOTRANSPORTE")) = TIPO_TRANSPORTE_CAMION) then
					if ((Cdbl(rs("FiTarifa")) < 10) or (Cdbl(rs("FiTarifa")) > 5000)) then pErrMsg = pErrMsg & "Carta de Porte: " & rs("nuCartaPorte") & " - " & tDSTransporte & ": " & rs("IDTRANSPORTE") & " - Tarifa x tonelada fuera de rango, debe ser entre $10 y $5000.<br>"
					if ((Cdbl(rs("TarifaRef")) < 10) or (Cdbl(rs("TarifaRef")) > 5000)) then pErrMsg = pErrMsg & "Carta de Porte: " & rs("nuCartaPorte") & " - " & tDSTransporte & ": " & rs("IDTRANSPORTE") & " - Tarifa de Referencia fuera de rango, debe ser entre $10 y $5000.<br>"
					if (Cdbl(rs("KMRecorrer")) <= 0) then pErrMsg = pErrMsg & "Carta de Porte: " & rs("nuCartaPorte") & " - " & tDSTransporte & ": " & rs("IDTRANSPORTE") & " - KM a Recorer no puede ser menor o igual a cero.<br>"								
				end if
			end if
			if (not IsNumeric(Trim(rs("establecimientoProcedencia")))) then pErrMsg = pErrMsg & "Carta de Porte: " & rs("nuCartaPorte") & " - " & tDSTransporte & ": " & rs("IDTRANSPORTE") & " - El Establecimiento de Procedencia no puede quedar vacio o es incorrecto.<br>"
			rs.MoveNext()
		wend	
		rs.MoveFirst()
		if (pErrMsg = "") then 
			ret = true
			logMig.info("Todo OK!")
		else
			logMig.info(pErrMsg)
		end if		
	else
		ret = true
		logMig.info("Se fuerza la NO validacion de datos.")
	end if
	logMig.info("--- Fin de la Validacion (" & ret  & ") ---")	
	validar = ret	
End Function	
'---------------------------------------------------------------------------------------------------------
Function armarregistroDatos(myRs, pFilename, pFilename2, tipoMovimiento, pValidar, pDtDesde, pDtHasta)
    
    Dim fs, myFile, myPuerto, myGrado, myAceptacion, errMsg, auxMsg, myCartaPorte, tDSTransporte
    Dim myFile2, dicT
	
    On Error Resume Next    
    
	Set dicT  = Server.CreateObject("Scripting.Dictionary")
	myPuerto = getNumeroPuerto(g_strPuerto)    
	
	if (validar(myRs, myPuerto, pValidar, errMsg)) then
		'Se abre el archivo de datos    
		Set fs = Server.CreateObject("Scripting.FileSystemObject")
		logMig.info("Inicializando archivo de datos: " & pFilename)		
		If fs.FileExists(pFilename) Then  Call fs.deleteFile(pFilename, true)
		Set myFile = fs.OpenTextFile(myFilename, 2, true)
		logMig.info("Inicializando archivo de Transportistas: " & pFilename)		
		If fs.FileExists(pFilename2) Then  Call fs.deleteFile(pFilename2, true)
		Set myFile2 = fs.OpenTextFile(myFilename2, 2, true)
		logMig.info("Archivos listo para trabajar.")        				
		while ((not myRs.eof) and (err.number = 0) and (errMsg = ""))			
			tDSTransporte = "Camion"
			if (CInt(myRs("TIPOTRANSPORTE")) = TIPO_TRANSPORTE_VAGON) then tDSTransporte = "Vagon"
			'1.- Puerto (1-2)
			registro = GF_nDigits(convertirDatoPuerto(CONV_KEY_PUERTO, myPuerto, auxMsg), 2)
			if (auxMsg <> "") then errMsg = errMsg & "Carta de Porte: " & myRs("nuCartaPorte") & " - " & tDSTransporte & ": " & myRs("IDTRANSPORTE") & " - " & auxMsg
			'2.- Planta
			registro = registro & GF_nDigits(convertirDatoPuerto(CONV_KEY_PLANTA, myPuerto, auxMsg), 2)
			if (auxMsg <> "") then errMsg = errMsg & "Carta de Porte: " & myRs("nuCartaPorte") & " - " & tDSTransporte & ": " & myRs("IDTRANSPORTE") & " - " & auxMsg
			'3.- Carta de Porte
			'Se corrige el nro para evitar repetición en los vagones.			
			myCartaPorte = Right(Trim(myRs("cartaPorte")), 9)
			myCartaPorte = myRs("idx") & myCartaPorte	
			myCartaPorte = GF_nDigits(myCartaPorte, 12)
			'Salvo la carta de porte para futuras referencias (Solo util en el caso de vagones)
			dicCtaPte.Add Trim(myRs("cartaPorte")), myCartaPorte
			registro = registro & myCartaPorte
			'4.- Cuit Vendedor
			registro = registro & GF_nDigits(myRs("cuitVendedor"), 11)
			'5.- Cuit Corredor
			registro = registro & GF_nDigits(myRs("cuitCorredor"), 11)
			'6.- Cuit Entregador
			registro = registro & GF_nDigits(myRs("cuitEntregador"), 11)
			'7.- Producto
			registro = registro & GF_nDigits(convertirDatoPuerto(CONV_KEY_PRODUCTO, myRs("cdProducto"), auxMsg), 3)
			if (auxMsg <> "") then errMsg = errMsg & "Carta de Porte: " & myRs("nuCartaPorte") & " - " & tDSTransporte & ": " & myRs("IDTRANSPORTE") & " - " & auxMsg
			'8.- Fecha descarga
			registro = registro & GF_nDigits(myRs("DTContableDescarga"), 8)			
			'9.- Concepto
			registro = registro & "0"        
			'10.- Movimiento
			registro = registro & tipoMovimiento        
			'11.- Transporte
			registro = registro & myRs("tipoTransporte")
			'12.- Patente
			registro = registro & GF_nChars(Trim(myRs("cPatente")), 8," ",CHR_AFT)
			'13.- Filler
			registro = registro & " "
			'14.- Codigo Postal ONCCA
			registro = registro & GF_nDigits(myRs("localidadProcedenciaONCCA"), 5)
			'15.- Provincia
			registro = registro & GF_nDigits(myRs("provinciaProcedenciaONCCA"), 2)
			'16.- Kilos Brutos
			registro = registro & GF_nDigits(CLng(myRs("Bruto")) - CLng(myRs("Tara")), 9)        
			'17, 18 y 19.- Acondicionamiento 1
			if (dicAcond.Exists(myRs("DTContableDescarga") & "_" & Trim(myRs("NUCARTAPORTE")) & "_" & Trim(myRs("IDTRANSPORTE")) & "_SECADO")) then        
				aux = CDbl(dicAcond(myRs("DTContableDescarga") & "_" & Trim(myRs("NUCARTAPORTE")) & "_" & Trim(myRs("IDTRANSPORTE")) & "_SECADO"))				
				registro = registro & GF_nDigits(convertirDatoPuerto(CONV_KEY_SERVICIO, "SECADO", auxMsg), 2)
				if (auxMsg <> "") then errMsg = errMsg & "Carta de Porte: " & myRs("nuCartaPorte") & " - " & tDSTransporte & ": " & myRs("IDTRANSPORTE") & " - " & auxMsg
				'No se divide por 100 ya que se debe quitar la coma decimal (o sea multiplicarlo por 100 luego!)
				registro = registro & GF_nDigits(Round((CLng(myRs("Bruto")) - CLng(myRs("Tara"))) * aux/100, 0), 6)
				registro = registro & GF_nDigits(Round(aux * 100, 0), 5)
			'response.write myRS("cartaPorte") & " = " & GF_nDigits(Round((CLng(myRs("Bruto")) - CLng(myRs("Tara"))) * aux/100, 0), 6) & "<br>"
			else
				registro = registro & "0000000000000"
			end if                        
			'20, 21 y 22.- Acondicionamiento 2
			if (dicAcond.Exists(myRs("DTContableDescarga") & "_" & Trim(myRs("NUCARTAPORTE")) & "_" & Trim(myRs("IDTRANSPORTE")) & "_ZARANDA")) then				
				aux = CDbl(dicAcond(myRs("DTContableDescarga") & "_" & Trim(myRs("NUCARTAPORTE")) & "_" & Trim(myRs("IDTRANSPORTE")) & "_ZARANDA"))
				registro = registro & GF_nDigits(convertirDatoPuerto(CONV_KEY_SERVICIO, "ZARANDEO", auxMsg), 2)
				if (auxMsg <> "") then errMsg = errMsg & "Carta de Porte: " & myRs("nuCartaPorte") & " - " & tDSTransporte & ": " & myRs("IDTRANSPORTE") & " - " & auxMsg
				'No se divide por 100 ya que se debe quitar la coma decimal (o sea multiplicarlo por 100 luego!)
				registro = registro & GF_nDigits(Round((CLng(myRs("Bruto")) - CLng(myRs("Tara"))) * aux/100, 0), 6)
				registro = registro & GF_nDigits(round(aux * 100, 0), 5)
			else
				registro = registro & "0000000000000"
			end if                        
			'23, 24 y 25.- Acondicionamiento 3			
			if (dicAcond.Exists(myRs("DTContableDescarga") & "_" & Trim(myRs("NUCARTAPORTE")) & "_" & Trim(myRs("IDTRANSPORTE")) & "_TEMPERATURA")) then				
				aux = CDbl(dicAcond(myRs("DTContableDescarga") & "_" & Trim(myRs("NUCARTAPORTE")) & "_" & Trim(myRs("IDTRANSPORTE")) & "_TEMPERATURA"))				
				registro = registro & GF_nDigits(convertirDatoPuerto(CONV_KEY_SERVICIO, "TEMPERATURA", auxMsg), 2)
				if (auxMsg <> "") then errMsg = errMsg & "Carta de Porte: " & myRs("nuCartaPorte") & " - " & tDSTransporte & ": " & myRs("IDTRANSPORTE") & " - " & auxMsg
				'No se divide por 100 ya que se debe quitar la coma decimal (o sea multiplicarlo por 100 luego!)
				registro = registro & GF_nDigits(Round((CLng(myRs("Bruto")) - CLng(myRs("Tara"))) * aux/100, 0), 6)
				registro = registro & GF_nDigits(round(aux * 100, 0), 5)				
			else
				registro = registro & "0000000000000"
			end if                        
			'26, 27 y 28.- Acondicionamiento 4
			registro = registro & "0000000000000"
			'29, 30 y 31.- Rebaja 1
			registro = registro & "0000000000000"
			'32, 33 y 34.- Rebaja 2
			registro = registro & "0000000000000"
			'35, 36 y 37.- Rebaja 3
			registro = registro & "0000000000000"
			'38, 39 y 40.- Rebaja 4
			registro = registro & "0000000000000"        
			'Campos 41 a 56 - Analisis Solo corresponden si es rebaja convenida.
			if (CInt(myRs("tipoTransporte")) = TIPO_TRANSPORTE_CAMION) then            
				'strSQL=" Select RVC.*, CASE when (MAP.VALORGRADO = 'FUERA DE GRADO') then 3 when (MAP.VALORGRADO = 'N/A') then 0  else MAP.VALORGRADO end GRADO " &_
				'	   "     from " &_
				'	   "         (Select HRVC.*" &_
				'	   "             from HRUBROSVISTEOCAMIONES HRVC" &_
				'	   "             inner join HCALADADECAMIONES HCC on HCC.DTCONTABLE=HRVC.DTCONTABLE and HCC.IDCAMION=HRVC.IDCAMION and HCC.SQCALADA=HRVC.SQCALADA and HCC.CDACEPTACION=" & ACEPTACION_REBAJA_CONVENIDA &_		           
				'	   "             where HRVC.DTCONTABLE='" & GF_FN2DTCONTABLE(myRs("DTContableDescarga")) & "' and HRVC.IDCAMION='" & myRS("IDTRANSPORTE") & "' and HRVC.SQCALADA = (Select MAX(SQCALADA) FROM HCALADADECAMIONES CC where CC.DTCONTABLE='" & GF_FN2DTCONTABLE(myRs("DTContableDescarga")) & "' and CC.IDCAMION='" & myRS("IDTRANSPORTE") & "')" &_ 	
				'	   "         ) RVC " &_
				'	   "         left join MERMASAUTOMATICASPENALIZACION MAP on MAP.CDRUBRO=RVC.CDRUBRO and MAP.CDPRODUCTO=" & myRs("cdProducto") &_
				'	   "     where " &_
				'	   "         RVC.VLMERMA = 0 " &_
				'	   "         and VLBONREBAJA >= VALORMINIMO and VLBONREBAJA<=VALORMAXIMO" &_
				'	   "         and (MERMAVARIABLE > 0 or VALORGRADO not in ('1', '2'))"	               
				strSQL=" Select HRVC.*, CASE when HCC.CDGRADO is Null then 0 else HCC.CDGRADO end GRADO " &_					   
					   " 	from HRUBROSVISTEOCAMIONES HRVC" &_
					   "    	inner join HCALADADECAMIONES HCC on HCC.DTCONTABLE=HRVC.DTCONTABLE and HCC.IDCAMION=HRVC.IDCAMION and HCC.SQCALADA=HRVC.SQCALADA and HCC.CDACEPTACION=" & ACEPTACION_REBAJA_CONVENIDA &_		           
					   "        where HRVC.DTCONTABLE='" & GF_FN2DTCONTABLE(myRs("DTContableDescarga")) & "' and HRVC.IDCAMION='" & myRS("IDTRANSPORTE") & "' and HRVC.SQCALADA = (Select MAX(SQCALADA) FROM HCALADADECAMIONES CC where CC.DTCONTABLE='" & GF_FN2DTCONTABLE(myRs("DTContableDescarga")) & "' and CC.IDCAMION='" & myRS("IDTRANSPORTE") & "')" &_ 						   					   
					   "         and HRVC.VLMERMA = 0 and HRVC.CDRUBRO not in (" & gRubrosIgnorados & ", " & gRubroTemperatura & ", " & gRubrosZaranda & ", " & gRubrosHumedad & ", " & RUBRO_EXCLUSIVO_ZARANDA & ")"
				Call GF_BD_Puertos(g_strPuerto, rsA, "OPEN", strSQL)	               
			else
				'Los vagones pueden estar en la diaria o en la historica, dependiendo si el tren descargo completo o no.
				'strSQL=" Select RVC.*, CASE when (MAP.VALORGRADO = 'FUERA DE GRADO') then 3 when (MAP.VALORGRADO = 'N/A') then 0  else MAP.VALORGRADO end GRADO " &_
				'	   "     from " &_
				'	   "         (Select HRVC.*" &_
				'	   "             from RUBROSVISTEOVAGONES HRVC" &_
				'	   "             inner join CALADADEVAGONES HCC on HCC.CDOPERATIVO = HRVC.CDOPERATIVO and HCC.CDVAGON=HRVC.CDVAGON and HCC.SQCALADA=HRVC.SQCALADA and HCC.CDACEPTACION=" & ACEPTACION_REBAJA_CONVENIDA &_
				'	   "             inner join VAGONES HV on HV.CDOPERATIVO=HRVC.CDOPERATIVO and HV.CDVAGON=HRVC.CDVAGON " &_
				'	   "             where HV.DTCONTABLEVAGON='" & GF_FN2DTCONTABLE(myRs("DTContableDescarga")) & "' and HRVC.CDVAGON='" & myRS("IDTRANSPORTE") & "' and HRVC.SQCALADA = (Select MAX(SQCALADA) FROM HCALADADEVAGONES CC where CC.CDOPERATIVO = " & myRS("CDOPERATIVO")& " and CC.CDVAGON='" & myRS("IDTRANSPORTE") & "')" &_ 
				'	   "         ) RVC " &_
				'	   "         left join MERMASAUTOMATICASPENALIZACION MAP on MAP.CDRUBRO=RVC.CDRUBRO and MAP.CDPRODUCTO=" & myRs("cdProducto") &_
				'	   "     where " &_
				'	   "         RVC.VLMERMA = 0 " &_
				'	   "         and VLBONREBAJA >= VALORMINIMO and VLBONREBAJA<=VALORMAXIMO" &_
				'	   "         and (MERMAVARIABLE > 0 or VALORGRADO not in ('1', '2'))"	
				strSQL=" Select HRVC.*, CASE when HCC.CDGRADO is Null then 0 else HCC.CDGRADO end GRADO " &_					   
					   " 	from RUBROSVISTEOVAGONES HRVC" &_
					   "		inner join CALADADEVAGONES HCC on HCC.CDOPERATIVO = HRVC.CDOPERATIVO and HCC.CDVAGON=HRVC.CDVAGON and HCC.SQCALADA=HRVC.SQCALADA and HCC.CDACEPTACION=" & ACEPTACION_REBAJA_CONVENIDA &_
					   "        inner join VAGONES HV on HV.CDOPERATIVO=HRVC.CDOPERATIVO and HV.CDVAGON=HRVC.CDVAGON " &_
					   "    where HV.DTCONTABLEVAGON='" & GF_FN2DTCONTABLE(myRs("DTContableDescarga")) & "' and HRVC.CDVAGON='" & myRS("IDTRANSPORTE") & "' and HRVC.SQCALADA = (Select MAX(SQCALADA) FROM HCALADADEVAGONES CC where CC.CDOPERATIVO = " & myRS("CDOPERATIVO")& " and CC.CDVAGON='" & myRS("IDTRANSPORTE") & "')" &_ 	
					   "         and HRVC.VLMERMA = 0 and HRVC.CDRUBRO not in (" & gRubrosIgnorados & ", " & gRubroTemperatura & ", " & gRubrosZaranda & ", " & gRubrosHumedad & ", " & RUBRO_EXCLUSIVO_ZARANDA & ")"
				Call GF_BD_Puertos(g_strPuerto, rsA, "OPEN", strSQL)
				if ((rsA.eof) and (CLng(myRs("DtContable")) <> 0)) then	               
					'strSQL=" Select RVC.*, CASE when (MAP.VALORGRADO = 'FUERA DE GRADO') then 3 when (MAP.VALORGRADO = 'N/A') then 0  else MAP.VALORGRADO end GRADO " &_
					'	   "     from " &_
					'	   "         (Select HRVC.* " &_
					'	   "             from HRUBROSVISTEOVAGONES HRVC" &_
					'	   "             inner join HCALADADEVAGONES HCC on HCC.DTCONTABLE=HRVC.DTCONTABLE and HCC.CDOPERATIVO = HRVC.CDOPERATIVO and HCC.CDVAGON=HRVC.CDVAGON and HCC.SQCALADA=HRVC.SQCALADA and HCC.CDACEPTACION=" & ACEPTACION_REBAJA_CONVENIDA &_
					'	   "             inner join HVAGONES HV on HV.DTCONTABLE=HRVC.DTCONTABLE and HV.CDOPERATIVO=HRVC.CDOPERATIVO and HV.CDVAGON=HRVC.CDVAGON " &_
					'	   "             where HV.DTCONTABLEVAGON='" & GF_FN2DTCONTABLE(myRs("DTContableDescarga")) & "' and HRVC.CDVAGON='" & myRS("IDTRANSPORTE") & "' and HRVC.SQCALADA = (Select MAX(SQCALADA) FROM HCALADADEVAGONES CC where CC.DTCONTABLE='" & GF_FN2DTCONTABLE(myRs("DtContable")) & "' and CC.CDOPERATIVO = " & myRS("CDOPERATIVO")& " and CC.CDVAGON='" & myRS("IDTRANSPORTE") & "')" &_ 
					'	   "         ) RVC " &_
					'	   "         left join MERMASAUTOMATICASPENALIZACION MAP on MAP.CDRUBRO=RVC.CDRUBRO and MAP.CDPRODUCTO=" & myRs("cdProducto") &_
					'	   "     where " &_
					'	   "         RVC.VLMERMA = 0 " &_
					'	   "         and VLBONREBAJA >= VALORMINIMO and VLBONREBAJA<=VALORMAXIMO" &_
					'	   "         and (MERMAVARIABLE > 0 or VALORGRADO not in ('1', '2'))"	               
					strSQL=" Select HRVC.*, CASE when HCC.CDGRADO is Null then 0 else HCC.CDGRADO end GRADO " &_					   
						   " 	from HRUBROSVISTEOVAGONES HRVC" &_
						   "    	inner join HCALADADEVAGONES HCC on HCC.DTCONTABLE=HRVC.DTCONTABLE and HCC.CDOPERATIVO = HRVC.CDOPERATIVO and HCC.CDVAGON=HRVC.CDVAGON and HCC.SQCALADA=HRVC.SQCALADA and HCC.CDACEPTACION=" & ACEPTACION_REBAJA_CONVENIDA &_
						   "        inner join HVAGONES HV on HV.DTCONTABLE=HRVC.DTCONTABLE and HV.CDOPERATIVO=HRVC.CDOPERATIVO and HV.CDVAGON=HRVC.CDVAGON " &_
						   "    where HV.DTCONTABLEVAGON='" & GF_FN2DTCONTABLE(myRs("DTContableDescarga")) & "' and HRVC.CDVAGON='" & myRS("IDTRANSPORTE") & "' and HRVC.SQCALADA = (Select MAX(SQCALADA) FROM HCALADADEVAGONES CC where CC.DTCONTABLE='" & GF_FN2DTCONTABLE(myRs("DtContable")) & "' and CC.CDOPERATIVO = " & myRS("CDOPERATIVO")& " and CC.CDVAGON='" & myRS("IDTRANSPORTE") & "')" &_ 	
						   "         AND HRVC.VLMERMA = 0 and HRVC.CDRUBRO not in (" & gRubrosIgnorados & ", " & gRubroTemperatura & ", " & gRubrosZaranda & ", " & gRubrosHumedad & ", " & RUBRO_EXCLUSIVO_ZARANDA & ")"
					Call GF_BD_Puertos(g_strPuerto, rsA, "OPEN", strSQL)	                   
				end if	                   
			end if      
			myGrado = "0"
			myAceptacion = "2"
			'Campos 41 a 44
			if (not rsA.eof) then                       
				myGrado = rsA("GRADO")
				myAceptacion = "1"				
				auxAnalisis = convertirDatoPuerto("ANALISIS" & Trim(myRs("cdProducto")), rsA("CDRUBRO"), auxMsg)
				if (auxMsg <> "") then errMsg = errMsg & "Carta de Porte: " & myRs("nuCartaPorte") & " - " & tDSTransporte & ": " & myRs("IDTRANSPORTE") & " - " & auxMsg
				registro = registro & Right(auxAnalisis, 2) & Left(auxAnalisis, 2) & GF_nDigits(CDbl(rsA("VLBONREBAJA"))*100, 5) & "000000"
				rsA.MoveNext()
			else
				registro = registro & "000000000000000"
			end if            
			'Campos 45 a 48
			if (not rsA.eof) then                       
				myGrado = rsA("GRADO")
				myAceptacion = "1"
				auxAnalisis = convertirDatoPuerto("ANALISIS" & Trim(myRs("cdProducto")), rsA("CDRUBRO"), auxMsg) 
				if (auxMsg <> "") then errMsg = errMsg & "Carta de Porte: " & myRs("nuCartaPorte") & " - " & tDSTransporte & ": " & myRs("IDTRANSPORTE") & " - " & auxMsg				
				registro = registro & Right(auxAnalisis, 2) & Left(auxAnalisis, 2) & GF_nDigits(CDbl(rsA("VLBONREBAJA"))*100, 5) & "000000"            
				rsA.MoveNext()
			else
				registro = registro & "000000000000000"
			end if
			'Campos 49 a 52
			if (not rsA.eof) then                       
				myGrado = rsA("GRADO")
				myAceptacion = "1"
				auxAnalisis = convertirDatoPuerto("ANALISIS" & Trim(myRs("cdProducto")), rsA("CDRUBRO"), auxMsg)        
				if (auxMsg <> "") then errMsg = errMsg & "Carta de Porte: " & myRs("nuCartaPorte") & " - " & tDSTransporte & ": " & myRs("IDTRANSPORTE") & " - " & auxMsg
				registro = registro & Right(auxAnalisis, 2) & Left(auxAnalisis, 2) & GF_nDigits(CDbl(rsA("VLBONREBAJA"))*100, 5) & "000000"            
				rsA.MoveNext()
			else
				registro = registro & "000000000000000"
			end if
			'Campos 53 a 56
			if (not rsA.eof) then                       
				myGrado = rsA("GRADO")
				myAceptacion = "1"
				auxAnalisis = convertirDatoPuerto("ANALISIS" & Trim(myRs("cdProducto")), rsA("CDRUBRO"), auxMsg)            
				if (auxMsg <> "") then errMsg = errMsg & "Carta de Porte: " & myRs("nuCartaPorte") & " - " & tDSTransporte & ": " & myRs("IDTRANSPORTE") & " - " & auxMsg
				registro = registro & Right(auxAnalisis, 2) & Left(auxAnalisis, 2) & GF_nDigits(CDbl(rsA("VLBONREBAJA"))*100, 5) & "000000"
				rsA.MoveNext()
			else
				registro = registro & "000000000000000"
			end if
			'57.- Kilos de Merma Volatil
			registro = registro & GF_nDigits(myRs("mermaVolatil"), 6)
			'58.- Humedad
			registro = registro & GF_nDigits(round(CDbl(myRs("HUMEDAD"))*100, 0), 4)
			'59.- Proteina
			registro = registro & GF_nDigits(round(CDbl(myRs("PROTEINA"))*100, 0), 4)
			'60.- Grado
			registro = registro & myGrado        
			'61.- Conforme/Condicional
			registro = registro & myAceptacion        
			'62.- Filler
			registro = registro & space(8)
			'63.- Cuit Chofer
			registro = registro & GF_nDigits(myRs("cuitChofer"), 11)
			'64.- Filler
			registro = registro & space(3)
			'65.- Razon Social Chofer
			registro = registro & GF_nChars(Left(myRs("dsChofer"), 50), 50," ",CHR_AFT)
			'66.- Cuit Transportista.
			registro = registro & GF_nDigits(myRs("cuitTransportista"), 11)
			'67.- Filler
			registro = registro & space(50)
			'68.- IVA Transportista
			registro = registro & GF_nDigits(convertirDatoPuerto(CONV_KEY_CATEG_IVA, myRs("ivaTransportista"), auxMsg), 1)
			if (auxMsg <> "") then errMsg = errMsg & "Carta de Porte: " & myRs("nuCartaPorte") & " - " & tDSTransporte & ": " & myRs("IDTRANSPORTE") & " - " & auxMsg
			'69, 70 y 71
			registro = registro & space(32)
			'72.- Cuit Intermediario
			registro = registro & GF_nDigits(myRs("cuitIntermediario"), 11)
			'73.- Razon Social Intermediario
			registro = registro & GF_nChars(Left(myRs("dsIntermediario"), 50), 50," ",CHR_AFT)
			'74.- Cuit Rte Comercial
			registro = registro & GF_nDigits(myRs("cuitRteComercial"), 11)
			'75.- Razon Social Rte Comercial
			registro = registro & GF_nChars(Left(myRs("dsRteComercial"), 50), 50," ",CHR_AFT)
			'76.- Numero CAU
			registro = registro & GF_nDigits(myRs("NCAU"), 14)
			'77.- Vencimiento CAU
			registro = registro & myRs("vtoNCAU")
			'78.- Fecha Carga
			registro = registro & myRs("fechaCarga")
			'79.- Peso Carga
			registro = registro & GF_nDigits(myRs("pesoCarga"), 9)
			'80.- Patente Acoplado
			registro = registro & GF_nChars(Trim(myRs("cPatenteAcoplado")),8," ",CHR_AFT)
			'81.- Filler
			registro = registro & space(4)
			'82.- Km a Recorrer
			registro = registro & GF_nDigits(myRs("KMRecorrer"), 4)
			'83.- Tarifa x tonelada
			registro = registro & GF_nDigits(round(CDbl(myRs("FiTarifa"))*100, 0), 8)
			'84.- CTG
			registro = registro & GF_nChars(myRs("ncCTG"), 10," ",CHR_AFT)
			'85.- Cosecha
			registro = registro & Mid(myRs("cdcosecha"), 3, 2) & Right(myRs("cdcosecha"), 2)
			'86.- Establecimiento Procedencia
			aux = Trim(myRs("establecimientoProcedencia"))
			if (CLng(aux) = 1) then aux = ""        
			registro = registro & GF_nChars(aux, 6," ",CHR_AFT)
			'87.- Establecimiento destino
			registro = registro & GF_nChars(myRs("establecimientoDestino"), 6," ",CHR_AFT)
			'88.- Tarifa de Referencia
			registro = registro & GF_nDigits(round(CDbl(myRs("TarifaRef")) * 100, 0), 8)
			'89.- Cuit titular
			registro = registro & GF_nDigits(myRs("cuitRemitente"), 11)
			'90.- CUIT Destinatario.
			registro = registro & GF_nDigits(myRs("cuitDestinatario"), 11)
			'91 y 92.- Filler
			registro = registro & space(16)
		
			myFile.WriteLine registro        			
			
			if (not dicT.Exists(Trim(myRs("cuitTransportista")))) then
				'Se agrega el registro de transportista
				'1.- CUIT TRANSPORTISTA
				registro = GF_nDigits(myRs("cuitTransportista"), 11)
				'2.- Razon Social
				registro = registro & GF_nChars(myRs("dsTransportista"), 55," ",CHR_AFT)
				'3.- Domicilio Calle
				registro = registro & GF_nChars(myRs("calleTransportista"), 55," ",CHR_AFT)
				'4.- Domicilio Numero
				registro = registro & GF_nDigits(myRs("calleNroTransportista"), 8)
				'5.- Tipo Domicilio
				registro = registro & GF_nDigits(myRs("tdomTransportista"), 1)
				'6.- Codigo Postal Magic (MET037A)
				registro = registro & GF_nDigits(myRs("cpostalTransportista"), 5)
				myFile2.WriteLine registro 
				
				dicT.add Trim(myRs("cuitTransportista")), 1
			end if			      
			
			myRs.MoveNext()        						
	
		wend    
		
		myFile.Close()
		myFile2.Close()		
		
		Set myFile = Nothing
		Set myFile2 = Nothing
		Set fs = Nothing
					
	end if
	if ((err.number > 0) or (errMsg <> "")) then		
		mail_config_Type=MAIL_TYPE_HTML			
		if (errMsg = "") then errMsg = err.Description
		errMsg = "Se han encontrado errores armando archivos de descargas: <br>" & errMsg
		Call SendMail(TASK_POS_DESCARGA_TERCEROS, MAIL_TASK_ERROR_LIST, "Descargas de ADM - " & getNombrePuerto(g_strPuerto) & " - " & GF_FN2DTE(pDtDesde) & " a " & GF_FN2DTE(pDtHasta) & " - FALTAN DATOS EN CARTA DE PORTE", errMsg, "")
		response.write errMsg
		response.end
	end if		
End Function
'---------------------------------------------------------------------------------------------------------
Function enviarMailDescargaTercero(pPto, pFecha, pCuitDestinatario, pFileAttachment, pFileAttachment2, pFileAttachment3, pFileAttachmentXLS)
    Dim strBody, strSubject,fs,auxFileAtt, myLista, dsDestinatario
    
    myLista = cuitDestinatario	
    dsDestinatario = getDsClienteByCUIT(cuitDestinatario)
    strSubject = "Descargas de " & dsDestinatario & " en " & getNombrePuerto(pPto) & " del " & GF_FN2DTE(pFecha)  
    logMig.info(" Enviando mail de la tarea " & TASK_POS_DESCARGA_TERCEROS & " con codigo "& myLista )    
    if (pFileAttachment <> "") then
        strBody = "Se envia adjunto el archivo con las descargas realizadas y el reporte de la posicion terminal de " & dsDestinatario
        Call SendMail(TASK_POS_DESCARGA_TERCEROS, myLista, strSubject, strBody, pFileAttachment &";"& pFileAttachment2 & ";"& pFileAttachment3 & ";"& pFileAttachmentXLS)
    else
        strBody = "No se registraron descargas para " & dsDestinatario & " en la terminal este día " & GF_FN2DTE(pFecha) & "."
        Call SendMail(TASK_POS_DESCARGA_TERCEROS, myLista, strSubject, strBody, "")
		response.write strBody
    end if	
End Function
'---------------------------------------------------------------------------------------------------------
Function generarArchivoAnalisis(pPto, pFilename, pDtDesde, pDtHasta, pCliente)

	Dim strSQL, rs, fs

	Set fs = Server.CreateObject("Scripting.FileSystemObject")
	If fs.FileExists(pFilename) Then  Call fs.deleteFile(pFilename, true)
	
	logMig.info("Inicializando armado de archivo de analisis: " & pFilename)
	logMig.info("Procesando Camiones")	
	strSQL= "Select " & _
			"		FORMAT(HCD.DTCONTABLE, 'yyyyMMdd')	FECHADESCARGA, " & _
			"		HCD.NUCARTAPORTE	CARTAPORTE, " & _
			"		HCD.IDCAMION		IDTRANSPORTE, " & _
			" 		HRVC.CDRUBRO		RUBRO, " & _
			"		HRVC.VLBONREBAJA	RESULTADO, " & _
			"		HC.CDPRODUCTO		PRODUCTO, " & _
					TIPO_TRANSPORTE_CAMION & " TIPOTRANSPORTE, " & _
			"		case when HCC.CDGRADO is null then 0 else HCC.CDGRADO end	GRADO " & _
			"	from HCAMIONES HC " & _
			"		inner join HCAMIONESDESCARGA HCD on HCD.DTCONTABLE=HC.DTCONTABLE and HCD.IDCAMION=HC.IDCAMION " & _
			"		inner join HCALADADECAMIONES HCC on HCC.DTCONTABLE=HC.DTCONTABLE and HCC.IDCAMION=HC.IDCAMION and HCC.SQCALADA = (Select MAX(SQCALADA) from HCALADADECAMIONES X where X.DTCONTABLE=HCC.DTCONTABLE and X.IDCAMION=HCC.IDCAMION)	" & _
			"		inner join HRUBROSVISTEOCAMIONES HRVC on HRVC.DTCONTABLE=HCC.DTCONTABLE and HRVC.IDCAMION=HCC.IDCAMION and HRVC.SQCALADA=HCC.SQCALADA " & _
			"	where HC.DTCONTABLE >= '" & pDtDesde & "' " & _
			"		and 	HC.DTCONTABLE <= '" & pDtHasta & "' " & _
			"		and 	HC.CDESTADO in (6, 8)" & _
			"		and 	HCC.ICCAMARA='N' " & _ 
			" 		and 	HRVC.CDRUBRO not in (" & gRubrosIgnorados & ", " & gRubroTemperatura & ", " & gRubrosZaranda & ", " & gRubrosHumedad & ", " & RUBRO_EXCLUSIVO_ZARANDA & ")"			
	if (pCliente <> "") then strSQL = strSQL & " and HCD.CDCLIENTE = " &  pCliente
	Call armarRegistroAnalisis(pPto, pFilename, pDtDesde, pDtHasta, strSQL)
	
	logMig.info("Procesando Vagones")
	strSQL = " Select * from ( " & _
			"	Select " & _
			"		FORMAT(HC.DTCONTABLEVAGON, 'yyyyMMdd')					FECHADESCARGA, " & _
			"		LEFT(CONCAT(HC.NUCARTAPORTESERIE, HC.NUCARTAPORTE), 12) 	CARTAPORTE, " & _
			"		HC.CDVAGON		IDTRANSPORTE, " & _
			"		HRVC.CDRUBRO		RUBRO, " & _
			"		HRVC.VLBONREBAJA	RESULTADO, " & _
			"		HC.CDPRODUCTO		PRODUCTO, " & _
					TIPO_TRANSPORTE_VAGON & " TIPOTRANSPORTE, " & _
			"		case when HCC.CDGRADO is null then 0 else HCC.CDGRADO end	GRADO " & _
			"	from VAGONES HC 	" & _
			"		inner join OPERATIVOS OP on OP.CDOPERATIVO=HC.CDOPERATIVO" & _
			"		inner join CALADADEVAGONES HCC on HCC.NUCARTAPORTE=HC.NUCARTAPORTE and HCC.CDVAGON=HC.CDVAGON and HCC.SQCALADA = (Select MAX(SQCALADA) from CALADADEVAGONES X where X.NUCARTAPORTE=HCC.NUCARTAPORTE and X.CDVAGON=HCC.CDVAGON)	" & _
			"		inner join RUBROSVISTEOVAGONES HRVC on HRVC.NUCARTAPORTE=HCC.NUCARTAPORTE and HRVC.CDVAGON=HCC.CDVAGON and HRVC.SQCALADA=HCC.SQCALADA " & _
			"	where HC.DTCONTABLEVAGON >= '" & pDtDesde & "' " & _
			"		and 	HC.DTCONTABLEVAGON <= '" & pDtHasta & "' " & _
			"		and 	HC.CDESTADO in (6, 8)" & _
			"		and 	HCC.ICCAMARA='N' " & _ 
			" 		and 	HRVC.CDRUBRO not in (" & gRubrosIgnorados & ", " & gRubroTemperatura & ", " & gRubrosZaranda & ", " & gRubrosHumedad & ", " & RUBRO_EXCLUSIVO_ZARANDA & ")"
	if (pCliente <> "") then strSQL = strSQL & " and OP.CDCLIENTE = " &  pCliente
	strSQL = strSQL & "	UNION " & _
			"	Select " & _
			"		FORMAT(HC.DTCONTABLEVAGON, 'yyyyMMdd')					FECHADESCARGA, " & _
			"		LEFT(CONCAT(HC.NUCARTAPORTESERIE, HC.NUCARTAPORTE), 12) 	CARTAPORTE, " & _
			"		HC.CDVAGON		IDTRANSPORTE, " & _
			"		HRVC.CDRUBRO		RUBRO, " & _
			"		HRVC.VLBONREBAJA	RESULTADO, " & _
			"		HC.CDPRODUCTO		PRODUCTO, " & _
					TIPO_TRANSPORTE_VAGON & " TIPOTRANSPORTE, " & _
			"		case when HCC.CDGRADO is null then 0 else HCC.CDGRADO end	GRADO " & _
			"	from HVAGONES HC 	" & _
			"		inner join HOPERATIVOS OP on OP.CDOPERATIVO=HC.CDOPERATIVO" & _
			"		inner join HCALADADEVAGONES HCC on HCC.DTCONTABLE=HC.DTCONTABLE and HCC.NUCARTAPORTE=HC.NUCARTAPORTE and HCC.CDVAGON=HC.CDVAGON and HCC.SQCALADA = (Select MAX(SQCALADA) from HCALADADEVAGONES X where X.DTCONTABLE=HCC.DTCONTABLE and X.NUCARTAPORTE=HCC.NUCARTAPORTE and X.CDVAGON=HCC.CDVAGON)	" & _
			"		inner join HRUBROSVISTEOVAGONES HRVC on HRVC.DTCONTABLE=HCC.DTCONTABLE and HRVC.NUCARTAPORTE=HCC.NUCARTAPORTE and HRVC.CDVAGON=HCC.CDVAGON and HRVC.SQCALADA=HCC.SQCALADA " & _
			"	where HC.DTCONTABLEVAGON >= '" & pDtDesde & "' " & _
			"		and 	HC.DTCONTABLEVAGON <= '" & pDtHasta & "' " & _
			"		and 	HC.CDESTADO in (6, 8)" & _
			"		and 	HCC.ICCAMARA='N' " & _ 
			" 		and 	HRVC.CDRUBRO not in (" & gRubrosIgnorados & ", " & gRubroTemperatura & ", " & gRubrosZaranda & ", " & gRubrosHumedad & ", " & RUBRO_EXCLUSIVO_ZARANDA & ")"
	if (pCliente <> "") then strSQL = strSQL & " and OP.CDCLIENTE = " &  pCliente
	strSQL = strSQL & "	) T "
	Call armarRegistroAnalisis(pPto, pFilename, pDtDesde, pDtHasta, strSQL)
	
	Set fs = Nothing	
	
	logMig.info("Finalizando armado de archivo de analisis: " & pFilename)		
	
End Function
'---------------------------------------------------------------------------------------------------------
Function armarRegistroAnalisis(pPto, pFilename, pDtDesde, pDtHasta, pStrSQL)

	Dim myFile, registro, fs, rs, errMsg, myCtaPte
	Dim oldCtaPte, myRubroADM, auxMsg, tDSTransporte
	
	On Error Resume Next
	
	Set fs = Server.CreateObject("Scripting.FileSystemObject")	
	Set myFile = fs.OpenTextFile(pFilename, 8, true)		
	logMig.info("Obteniendo datos para analisis.")
	Call executeQueryDb(pPto, rs, "OPEN", pStrSQL)	
	
	if (not rs.eof) then						
		oldCtaPte = ""
		while (not rs.eof)
			tDSTransporte = "Camion"
			if (CInt(rs("TIPOTRANSPORTE")) = TIPO_TRANSPORTE_VAGON) then tDSTransporte = "Vagon"
			myCtaPte = Trim(rs("CARTAPORTE"))
			if (dicCtaPte.Exists(Trim(rs("CARTAPORTE")))) then myCtaPte = dicCtaPte(Trim(rs("CARTAPORTE")))
			if (myCtaPte <> oldCtaPte) then				
				'Si cambio la carta de porte agrego el rubro grado si hay.					
				logMig.info("Armando registro analisis para Cta. Pte:" & Trim(rs("CARTAPORTE")) & " - Rubro: GRADO (20)")										
				registro = rs("FECHADESCARGA")			
				registro = registro & myCtaPte
				registro = registro & GF_nDigits(rs("IDTRANSPORTE"), 10)
				registro = registro & "20"
				registro = registro & GF_nDigits(rs("GRADO"), 5)
				myFile.WriteLine registro				
				oldCtaPte = myCtaPte
			end if
			myRubroADM = Left(convertirDatoPuerto("ANALISIS" & Trim(rs("PRODUCTO")), rs("RUBRO"), auxMsg), 2)
			if (auxMsg <> "") then errMsg = errMsg & "Carta de Porte: " & Trim(rs("CARTAPORTE")) & " - " & tDSTransporte & ": " & myRs("IDTRANSPORTE") & " - " & auxMsg
			logMig.info("Armando registro analisis para Cta. Pte:" & Trim(rs("CARTAPORTE")) & " - Rubro: " & rs("RUBRO") & "(" & myRubroADM & ")")
			registro = rs("FECHADESCARGA")						
			registro = registro & myCtaPte
			registro = registro & GF_nDigits(rs("IDTRANSPORTE"), 10)
			registro = registro & myRubroADM
			registro = registro & GF_nDigits(CDbl(rs("RESULTADO"))*100, 5)
			myFile.WriteLine registro  				
			
			if ((err.number > 0) or (errMsg <> "")) then
				mail_config_Type=MAIL_TYPE_HTML							
				if (errMsg = "") then errMsg = err.Description				
				errMsg = "Se han encontrado errores armando archivos de analisis: <br>" & errMsg
				Call SendMail(TASK_POS_DESCARGA_TERCEROS, MAIL_TASK_ERROR_LIST, "Descargas de ADM - " & getNombrePuerto(pPto) & " - " & pDtDesde & " a " & pDtHasta & " - ERROR EN ANALISIS", errMsg, "")				
				response.write errMsg
				response.end
			end if	
			
			rs.MoveNext()
		wend				
	else
		logMig.info("No se encontró inforamcion de analisis.")
	end if
	
	myFile.Close()

	Set myFile = Nothing		
	Set fs = Nothing	
		
	
End Function
'---------------------------------------------------------------------------------------------------------
'                                   ***** COMIENZA PAGINA *****
'---------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------
Dim fechaDesde, fechaHasta, cliente, logMig, strSQL, rsDescarga, registro, dicAcond, myFilename, g_strPuerto
Dim flagUpdFecha, cuitDestinatario, flagValidacion, myFilename2
Dim dicCtaPte 'Diccionario para almacenar los nros de cata de porte de vagones ya que el nro debe modificarse para adaptarse al MAGIC que no maneja correctamente los vagones (patetico...)
Dim gRubroTemperatura, gRubrosIgnorados

Const PARAM_ULT_ENVIO = "DTENVIODESC_"

fechaDesde = GF_PARAMETROS7("fd", "", 6)
fechaHasta = GF_PARAMETROS7("fh", "", 6)
cliente = GF_PARAMETROS7("cl", "", 6)
g_strPuerto = GF_PARAMETROS7("pto", "", 6)

Call GP_ConfigurarMomentos()
'session("usuario") = "SYNC"

flagValidacion = true
if (GF_PARAMETROS7("v", "", 6) <> "") then flagValidacion = false

if (cliente <> "") then
	cuitDestinatario = getCUITCliente(cliente)
else
	cuitDestinatario = CUIT_TOEPFER
end if 

Set logMig = new classLog
Call startLog(HND_FILE,MSG_INF_LOG+MSG_ERR_LOG+MSG_WRN_LOG)
logMig.fileName = "EXPORTACION_DESCARGAS_"& Ucase(g_strPuerto) &"_" & cuitDestinatario & "_" & GF_nDigits(Year(Now),4) & GF_nDigits(Month(Now()),2) & GF_nDigits(Day(Now()),2)

'Si la fecha no vino por parametro, tomo la ultima migrada desde los parametros y se verifica si el archivo debe enviarse o ya se envio.
flagUpdFecha = false
if (fechaDesde = "") then 
	flagUpdFecha = true
    fechaUltimoEnvio = getValueParametro(PARAM_ULT_ENVIO & cuitDestinatario, g_strPuerto)    
    if (fechaUltimoEnvio <> "") then     
        fechaDesde = GF_DTEADD(fechaUltimoEnvio, 1, "D")
        fechaHasta = fechaDesde
        'Si la fecha a enviar es anterior a la ultima posible, se envia.
        ultimaFechaAutomatica = GF_DTEADD(Left(session("MmtoDato"), 8), -1, "D")
        if (CLng(fechaDesde) > CLng(ultimaFechaAutomatica)) then
            logMig.info("Los archivos ya fueron enviados para la fecha: " & GF_FN2DTE(fechaUltimoEnvio))
			response.write "Los archivos ya fueron enviados para la fecha: " & GF_FN2DTE(fechaUltimoEnvio)
            response.End
        end if                    
    end if        
end if

if (fechaDesde = "") then fechaDesde = GF_DTEADD(Left(session("MmtoDato"), 8), -1, "D")
if (fechaHasta = "") then fechaHasta = GF_DTEADD(Left(session("MmtoDato"), 8), -1, "D")

dtDesde = GF_FN2DTCONTABLE(fechaDesde)
dtHasta = GF_FN2DTCONTABLE(fechaHasta)

myFilename = server.MapPath(".\Temp") & "\DESCARGAS_" & g_strPuerto & "_" & cuitDestinatario & "_" & fechaDesde & ".dat"
myFilename2 = server.MapPath(".\Temp") & "\TRANSPORTISTAS_" & g_strPuerto & "_" & cuitDestinatario & "_" & fechaDesde & ".dat"
myFilename3 = server.MapPath(".\Temp") & "\ANALISIS_" & g_strPuerto & "_" & cuitDestinatario & "_" & fechaDesde & ".dat"
myFilenameXLS = "Posicion_" & g_strPuerto & "_" & cuitDestinatario & "_" & fechaDesde & ".xls"

gRubrosIgnorados = "36, 38"

logMig.info("--------------------- INCIANDO EXPORTACION ------------------------")	
logMig.info(" ---> PUERTO       : " & g_strPuerto)
logMig.info(" ---> FECHA DESDE  : " & dtDesde)
logMig.info(" ---> FECHA HASTA  : " & dtHasta)
logMig.info(" ---> CLIENTE      : " & cliente)
logMig.info(" ---> DERSTINATARIO: " & cuitDestinatario)
logMig.info(" ---> VALIDAR		: " & flagValidacion)
logMig.info("-------------------------------------------------------------------")	
response.write "--------------------- INCIANDO EXPORTACION ------------------------<br>"
response.write " ---> PUERTO       : " & g_strPuerto & "<br>"
response.write " ---> FECHA DESDE  : " & dtDesde & "<br>"
response.write " ---> FECHA HASTA  : " & dtHasta & "<br>"
response.write "-------------------------------------------------------------------<br>"

Set dicCtaPte = createObject("Scripting.Dictionary")    

gRubroTemperatura = getValueParametro("CDRUBROTEMPERATURA", g_strPuerto)	

'Se cargan los datos de acondicionamiento.
Call cargarTablaAcondicionamiento(dicAcond, fechaDesde, fechaHasta)
Set rsDescarga = obternerDescargasTerceros(dtDesde, dtHasta, cliente, g_strPuerto) 
'-->Call cargarValoresGlobalesFAC(g_strPuerto)  


'-->if (1=1) then
if (not rsDescarga.eof) then    
    'Se carga la tabla de conversiones para el cliente
    Call cargarTablaConversion(cuitDestinatario, g_strPuerto)      
    
    'Se arma el registro de datos.
    Call armarregistroDatos(rsDescarga, myFilename, myFilename2, 1, flagValidacion, fechaDesde, fechaHasta)
        
    logMig.info("Generando reporte posicion terminal...")
    'Se armar el reporte de la posicion terminal con el cliente
    myFilenameXLS = armarReporteTerminalXLS(g_strPuerto, myFilenameXLS, fechaDesde, cliente, XLS_FILE_MODE)
    
    'obtengo la ruta donde se guarda el reporte(dentro del raiz actisaintra carpeta temp)    
    logMig.info("Ruta del archivo en disco: "& myFilenameXLS)
    
	Call generarArchivoAnalisis(g_strPuerto, myFilename3, dtDesde, dtHasta, cliente)
	
    'Se envia por mail.
    Call enviarMailDescargaTercero(g_strPuerto, fechaDesde, cuitDestinatario, myFilename, myFilename2, myFilename3, myFilenameXLS)
	response.write "Envío Exitoso!"
       
else
	logMig.info("No se encontraron descargas para exportar")
	'Se envia por mail.
    Call enviarMailDescargaTercero(g_strPuerto, fechaDesde, cuitDestinatario, "", "", "", "")
end if

if (flagUpdFecha) then Call updateValueParametro(PARAM_ULT_ENVIO & cuitDestinatario,fechaDesde,g_strPuerto)

Set dicCtaPte = Nothing

logMig.info("--------------------------- FIN PROCESO ---------------------------")	
%>