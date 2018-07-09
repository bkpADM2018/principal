<!--#include file="cartadeporteEditCommon.asp"-->
<%

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function loadCaladasCamion(p_dtContable, p_idCamion,pPto)
	dim strSql, rsCalada,auxDtContable
    auxDtContable = Left(p_dtContable,4) & "-" &mid(p_dtContable,5,2) &"-"& Right(p_dtContable,2)
	strSql = "Select cc.ICCAMARA,cc.SQCALADA,cc.CDMOTIVORECHAZO,cc.vlhumedad,cc.vlproteina,cc.pcMerma,"&_
             "       (Cast(((Year(cc.dtcalada)*10000) + (Month(cc.dtcalada)*100) + Day(cc.dtcalada)) AS BIGINT)*1000000) + (Datepart(hour, cc.dtcalada) * 10000) + (Datepart(minute, cc.dtcalada) * 100) + Datepart(second, cc.dtcalada) as dtcalada , "&_
             "       cc.NUBARRAS,cc.nubalde,cc.cdaceptacion,AC.dsaceptacion,cc.CDFUERASTD,cc.DSOBSERVACIONES,cc.CDUSERNAME," &_
			 "       case when cc.ichumedimetro is null then 'N' else cc.ichumedimetro end as ichumedimetro, "&_
             "       case when cc.CDGRADO is null then '0' else RTrim(cc.CDGRADO) end as CDGRADO "&_
             " from (Select * "&_
             "       from hcaladadecamiones A "&_
             "        where A.DTCONTABLE='" & auxDtContable & "' and A.idCamion = '" & p_idCamion & "'"&_
             "              and A.SQCALADA = (SELECT MAX(SQCALADA) "&_
      		 "		                          FROM   hrubrosvisteocamiones  "&_
      		 "		                          WHERE  idcamion = A.idcamion AND dtcontable = A.dtcontable) "&_
             "      ) CC" &_
			 " inner join aceptacioncalidad AC on CC.CDACEPTACION=AC.CDACEPTACION"
    Call GF_BD_Puertos(pPto, rsCalada, "OPEN", strSql)
	if not rsCalada.Eof then
        g_IcCamara = rsCalada("ICCAMARA")
        g_IcCamaraOld = g_IcCamara
        g_UltimaCalada = rsCalada("SQCALADA")
        g_MotivoRechazo =rsCalada("CDMOTIVORECHAZO")
        g_vlHumedad = rsCalada("vlhumedad")
        g_vlProteina = rsCalada("vlproteina")
        g_PcMerma = rsCalada("pcMerma")
        'Tomo la fecha de la base de datos, como es un campo fecha puede haber problemas de formato por el idioma, lo estandarizo.
        g_DtCalada = rsCalada("dtcalada")        
        g_DtCalada = GF_FN2DTE(rsCalada("dtcalada"))
        g_IcHumedimetro = rsCalada("ichumedimetro")
        g_Sticker = Trim(rsCalada("NUBARRAS"))
        g_StickerOld = g_Sticker
        g_Aceptacion = rsCalada("cdaceptacion")
        g_AceptacionOld = g_Aceptacion
        g_Balde = rsCalada("nubalde")
        g_BaldeOld = g_Balde
        g_cdGrado = Cdbl(rsCalada("CDGRADO"))
        g_cdGradoOld = g_cdGrado
        g_FueraStandart = rsCalada("CDFUERASTD")
        g_DsObservaciones = Trim(rsCalada("DSOBSERVACIONES"))
        g_CdUsuario = rsCalada("CDUSERNAME")
    end if
End Function
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function getRubrosVisteoCamion(p_dtContable, p_idCamion, p_sqCalada,pPto)
	dim strSql, rsRub,auxDtContable
    auxDtContable = Left(p_dtContable,4) & "-" &mid(p_dtContable,5,2) &"-"& Right(p_dtContable,2)
	strSql = "Select A.*,RTRIM(B.dsrubro) as dsrubro "&_
             " from ( Select * from HRUBROSVISTEOCAMIONES "&_
             "        where dtcontable = '"& auxDtContable & "' and idcamion = '"& p_idCamion &"' and sqcalada = "& p_sqCalada &") A"&_
             "  inner join rubros B on A.cdrubro = B.cdrubro "&_
             " order by A.cdRubro"
	Call GF_BD_Puertos(pPto, rsRub, "OPEN", strSql)
	Set getRubrosVisteoCamion = rsRub
End Function
'-----------------------------------------------------------------------------------------------------------
Function controlarCambiosCalada()    
    if (CStr(g_IcCamara) <> CStr(g_IcCamaraOld)) then Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "CAMARA|"& g_IcCamaraOld &"|"& g_IcCamara)
    if (Trim(CStr(g_Sticker)) <> Trim(CStr(g_StickerOld))) then Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "STICKER|"& g_StickerOld &"|"& g_Sticker)
    if (Cdbl(g_Aceptacion) <> Cdbl(g_AceptacionOld)) then Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "ACEPTACION|"& getDsAceptacion(g_AceptacionOld) &"|"& getDsAceptacion(g_Aceptacion))
    if (Cdbl(g_Balde) <> Cdbl(g_BaldeOld)) then Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "BALDE|"& g_BaldeOld &"|"& g_Balde)
    if (Cdbl(g_cdGrado) <> Cdbl(g_cdGradoOld)) then Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "GRADO|"& g_cdGradoOld &"|"& g_cdGrado)
End Function
'-----------------------------------------------------------------------------------------------------------
Function controlarCambiosRubros()
	Dim dsRubro
    if (CInt(g_estado) = ESTADO_ACTIVO) then
        dsRubro = g_cdRubro &"-"& getDsRubro(g_cdRubro)
        if (CInt(g_estadoOld) = ESTADO_ACTIVO) then
            if (Cdbl(g_VlRebaja) <> Cdbl(g_VlRebajaOld)) then Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "VALOR (Rubro "& dsRubro&")|"& g_VlRebajaOld &"|"& g_VlRebaja)
            if (Cdbl(g_VlPeso) <> Cdbl(g_VlPesoOld)) then Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "PESO RUBRO (Rubro "&dsRubro&")|"& g_VlPesoOld &"|"& g_VlPeso)
            if (Cdbl(g_VlPCPeso) <> Cdbl(g_VlPCPesoOld)) then Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "PC PESO RUBRO (Rubro "&dsRubro&")|"& g_VlPCPesoOld &"|"& g_VlPCPeso)
            if (Cdbl(g_VlMerma) <> Cdbl(g_VlMermaOld)) then Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "MERMA (Rubro "&dsRubro&")|"& g_VlMermaOld &"|"& g_VlMerma)
        else
            Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "Se agrego el rubro "& dsRubro)
        end if
    else
        if (CInt(g_estadoOld) = ESTADO_ACTIVO) then Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "Se elimino el rubro "& g_cdRubro &"-"& getDsRubro(g_cdRubro))
    end if
End Function
'-----------------------------------------------------------------------------------------------------------
Function grabarRubro(p_dtContable,p_idCamion,p_strPuerto,p_UltimaCalada,p_CdProducto,p_DsObservaciones,p_Bruto,p_Tara)
    Dim strSQL,i,auxDtContable,myTotalMermaDetalleCalc 
    auxDtContable = GF_FN2DTCONTABLE(p_dtContable)
    
    g_listaRubrosHumedad = "," & getValueParametro(PARAM_CD_RUBRO_HUMEDAD,p_strPuerto) & ","
    g_listaRubrosZaranda = "," & getValueParametro(PARAM_CD_RUBRO_ZARANDA,p_strPuerto) & ","
    g_listaRubrosProteina= "," & getValueParametro(PARAM_CD_RUBRO_PROTEINA,p_strPuerto) & ","
    g_RowActual = 0
    myTotalMermaDetalleCalc = 0
    vlPorteina = 0
    vlHumedad = 0
    while (readNextRubroParams())
        if (Cdbl(g_estado) = ESTADO_ACTIVO) then
            strSQL = "INSERT INTO HRUBROSVISTEOCAMIONES(DTCONTABLE,IDCAMION,SQCALADA,CDRUBRO,VLBONREBAJA,VLPESORUBRO,PCPESORUBRO,DSRESULTADO,CDSUPERVISOR,VLMERMA)"&_
                     " VALUES('"&auxDtContable&"','"&p_idCamion&"',"&Cdbl(p_UltimaCalada)+1&","&g_cdRubro&","&g_VlRebaja&","&g_VlPeso&","&g_VlPCPeso&",'0','',"&g_VlMerma&")"
            Call logMig.info("Agrega Rubro: "& strSQL )
            Call GF_BD_Puertos(p_strPuerto, rs, "EXEC",strSQL)
            if (InStr(1, g_listaRubrosHumedad, "," & g_cdRubro & ",") > 0) then vlHumedad = g_VlRebaja
            if (InStr(1, g_listaRubrosProteina, "," & g_cdRubro & ",") > 0) then vlPorteina = g_VlRebaja
            'Para obtener el total porcentaje de merma debo sumar todos lo % de merma de cada rubro calculado
            myTotalMermaDetalleCalc = myTotalMermaDetalleCalc + g_VlMerma
        end if
        g_RowActual = g_RowActual + 1
    wend
    'Agrego la Calada(HCALADADECAMIONES)
    Call agregarCaladaCamiones(auxDtContable,p_idCamion,p_UltimaCalada,vlHumedad,vlPorteina,myTotalMermaDetalleCalc,p_strPuerto,p_DsObservaciones)
    
    g_Neto = Cdbl(p_Bruto) - Cdbl(p_Tara)
    'Para enganchar las tablas de mermas necesito crear una nueva pesada con los mimsmo valores de la anterior    
    Call agregarPesadaCamiones(auxDtContable,p_idCamion,p_strPuerto,p_Bruto,p_Tara)
    'Calculo la merma en kilos multiplicando los kilos netos con el porcentaje nuevo de merma
    g_Merma = Round((Cdbl(g_Neto)*Cdbl(myTotalMermaDetalleCalc))/100)
    'Actualizo la merma 
    Call actualizarMerma(auxDtContable,p_idCamion,Cdbl(g_Merma),p_strPuerto)
    Call actualizarHAcondicionamientoProductos(auxDtContable,p_idCamion,p_UltimaCalada,myTotalMermaDetalleCalc,Cdbl(g_Merma),p_strPuerto)
End Function
'----------------------------------------------------------------------------------------------------------
'Agrego la nueva pesada respitiendo los valores solo cambio la sqPesada
Function agregarPesadaCamiones(p_DtContable,p_idCamion,p_strPuerto,p_KiloBruto,p_KiloTara)
    
    Call agregarPesada(p_idCamion, p_DtContable, p_KiloBruto, PESADA_BRUTO, p_strPuerto)
    
    Call agregarPesada(p_idCamion, p_DtContable, p_KiloTara, PESADA_TARA, p_strPuerto)
    
End Function
'***************************************************************************************************************************
'************************************************ TRANSACCIONES DE AUDITORIA ***********************************************
'***************************************************************************************************************************
Function agregarRegistrosAuditoria(p_DtContable,p_idCamion,p_UltimaCalada,p_strPuerto)
    Dim rsAud,rs,auxDtContable    
    'Formateo la fecha contable
    auxDtContable = GF_FN2DTCONTABLE(p_DtContable)    

    'Agrega Auditoria de Calada
    Call agregarHAudCaladaCamiones(p_idCamion, auxDtContable, p_UltimaCalada, p_strPuerto)
    'Agrega Auditoria de Rubros
    Call agregarHAudRubrosVisteoCamiones(p_idCamion, auxDtContable, p_UltimaCalada, p_strPuerto)
    'Agrega Auditoria de Pesada Bruto
    Call agregarHAudPesadasCamion(p_idCamion, auxDtContable, PESADA_BRUTO, p_strPuerto)
    'Agrega Auditoria de Pesada Tara
    Call agregarHAudPesadasCamion(p_idCamion, auxDtContable, PESADA_TARA, p_strPuerto)
    'Agrega Auditoria de Merma
    Call agregarHAudMermas(p_idCamion, auxDtContable, p_strPuerto)
End Function
'----------------------------------------------------------------------------------------------------------
Function agregarHAudCaladaCamiones(p_idCamion,p_DtContable,p_UltimaCalada,p_strPuerto)
    Dim strSQL
    strSQL = "Insert into dbo.HAudCaladaDeCamiones "&_
             " Select A.dtContable,(SELECT MAX(SQAUDITORIA)+1 AS SQAUDITORIA FROM HAudCaladaDeCamiones WHERE DTCONTABLE='"&p_DtContable&"' AND IDCAMION='"& p_idCamion&"'), A.idCamion, A.sqCalada, "&_
             " A.icTipoCalada, A.vlHumedad, A.vlProteina, A.cdAceptacion, "&_
             " A.cdRubroPpal, A.cdgrado, A.icCondicionFabrica, A.cdUserName, A.dtCalada,"&_
             " A.DsObservaciones, A.nuBarras, A.cdMotivoRechazo,A.PcMerma, A.icCamara, A.icHumedimetro, A.dsObsHumedimetro,A.cdSupervisor, A.NuBalde, A.cdFueraSTD, A.IdContrato "&_
             " FROM dbo.HCaladadeCamiones A "&_
             " WHERE A.DtContable='" & p_DtContable & "' and A.idCamion='"& p_idCamion &"' AND A.sqcalada="& Cdbl(p_UltimaCalada)+1 &" AND A.IcTipoCalada= 'V'"
    Call logMig.info("Agrega HAudCaladaDeCamiones: "& strSQL )
    Call GF_BD_Puertos(p_strPuerto, rsAud, "EXEC",strSQL)
End Function
'----------------------------------------------------------------------------------------------------------
Function agregarHAudRubrosVisteoCamiones(p_idCamion,p_DtContable,p_UltimaCalada,p_strPuerto)
    Dim strSQL
    strSQL = "Insert into dbo.HAudRubrosVisteoCamiones "&_
             " select A.dtcontable,(SELECT MAX(SQAUDITORIA)+1 AS SQAUDITORIA FROM HAudRubrosVisteoCamiones WHERE DTCONTABLE='"&p_DtContable&"' AND IDCAMION='"& p_idCamion&"'),A.idcamion,A.sqcalada,A.cdrubro,A.vlbonrebaja,A.vlpesorubro,A.pcpesorubro,A.dsresultado,A.cdsupervisor,A.vlmerma "&_
             " from HRUBROSVISTEOCAMIONES A where A.DtContable='" & p_DtContable & "' and A.idCamion='"& p_idCamion &"' AND A.sqcalada="& Cdbl(p_UltimaCalada)+1
    Call logMig.info("Agrega HAudRubrosVisteoCamiones: "& strSQL )
    Call GF_BD_Puertos(p_strPuerto, rs, "EXEC",strSQL)
End function
'----------------------------------------------------------------------------------------------------------
Function agregarHAudPesadasCamion(p_idCamion,p_DtContable,p_TipoPesada,p_strPuerto)
    Dim strSQL
    strSQL = "Insert into dbo.HAudPesadasCamion "&_
             " select A.DTCONTABLE,(SELECT MAX(SQAUDITORIA)+1 AS SQAUDITORIA FROM HAUDPESADASCAMION WHERE DTCONTABLE='"&p_DtContable&"' AND IDCAMION='"& p_idCamion&"'),A.IDCAMION,A.SQPESADA,A.CDPESADA,A.VLPESADA,A.ICMETODO,A.CDPUESTO,A.CDUSERNAME,A.DTPESADA,A.PCKILOS "&_
             " from HPESADASCAMION A where A.DtContable='" & p_DtContable & "' and A.idCamion='"& p_idCamion &"'"&_
             "   AND A.sqpesada=(SELECT MAX(SQPESADA) FROM HPESADAsCAMION WHERE IDCAMION = A.IDCAMION AND DTCONTABLE = A.DTCONTABLE  AND CDPESADA = "& p_TipoPesada &")"
    Call logMig.info("Agrega HAudPesadasCamion Bruto: "& strSQL )
    Call GF_BD_Puertos(p_strPuerto, rs, "EXEC",strSQL)
End Function
'***************************************************************************************************************************
'************************************************ TRANSACCIONES OTRAS TABLAS ***********************************************
'***************************************************************************************************************************
Function agregarGrupoEnsayos(p_DtContable,p_idCamion,p_UltimaCalada,p_strPuerto)
    Dim strSQL,rsEns,rs,auxDtContable 
    auxDtContable = GF_FN2DTCONTABLE(p_DtContable)
    strSQL = "select g.cdgrupo as  Grupo, '' as Ensayo from dbo.hgruposensayoscamiones g "&_
             " where g.IdCamion ='"& p_idCamion &"' and g.dtContable = '" & auxDtContable  & "'"&_
             " and g.sqCalada= " & p_UltimaCalada &_
             " union all select '' as Grupo, e.cdensayo as Ensayo from dbo.hensayoscamiones e "&_
             " where e.IdCamion = '"& p_idCamion &"' and e.dtContable = '" & auxDtContable  & "'"&_
             " and e.sqCalada=" & p_UltimaCalada
    Call GF_BD_Puertos(p_strPuerto, rsEns, "OPEN", strSQL)
    while not rsEns.Eof 
        if (rsEns("GRUPO") <> "") then
            strSQL="INSERT INTO hgruposensayoscamiones(DTCONTABLE,IDCAMION,SQCALADA,CDGRUPO)VALUES('"&auxDtContable &"','"&p_idCamion&"',"&Cdbl(p_UltimaCalada)+1&",'"&rsEns("GRUPO")&"')"
        end if
        if (rsEns("ENSAYO") <> "") then
            strSQL="INSERT INTO hensayoscamiones(DTCONTABLE,IDCAMION,SQCALADA,CDENSAYO)VALUES('"&auxDtContable &"','"&p_idCamion&"',"&Cdbl(p_UltimaCalada)+1&",'"&rsEns("ENSAYO")&"')"
        end if
        Call logMig.info("Agrega Grupo/Ensayos: "& strSQL )
        Call GF_BD_Puertos(p_strPuerto, rs, "EXEC",strSQL)
        rsEns.MoveNext()
    wend 
End Function
'----------------------------------------------------------------------------------------------------------
Function agregarMuestraHumedad(p_DtContable,p_idCamion,p_UltimaCalada,p_strPuerto)
    Dim strSQL,auxDtContable,rsMuesta
    auxDtContable = GF_FN2DTCONTABLE(p_DtContable)
    strSQL = "SELECT * FROM HMuestrasHumedCamiones "&_
             "where dtcontable = '"& auxDtContable &"' and idcamion= '"&p_idCamion&"' and sqcalada ="&p_UltimaCalada
    Call GF_BD_Puertos(p_strPuerto, rsMuestra, "OPEN",strSQL)
    while (not rsMuestra.Eof)
        strSQL ="INSERT INTO HMuestrasHumedCamiones values('"& auxDtContable &"','"&p_idCamion&"',"& Cdbl(p_UltimaCalada)+1 &","& rsMuestra("sqmuestra") &","& rsMuestra("vlhumedad") &","& rsMuestra("vltemperatura") &","& rsMuestra("vlpeso") &")"
        Call logMig.info("Agrega HMuestrasHumedCamiones: "& strSQL )
        Call GF_BD_Puertos(p_strPuerto, rs, "EXEC",strSQL)
        rsMuestra.MoveNext()
    wend
End Function
'----------------------------------------------------------------------------------------------------------
Function agregarCaladaCamiones(p_DtContable,p_idCamion,p_UltimaCalada,p_vlHumedad,p_vlPorteina,p_vlMerma,p_strPuerto,p_DsObservaciones)
   auxDtCalada = Year(Now())&"-"&GF_nDigits(Month(Now()),2)&"-"&GF_nDigits(Day(Now()),2)&" "&GF_nDigits(Hour(Now()),2)&":"&GF_nDigits(Minute(Now()),2)&":"&GF_nDigits(Second(Now()),2)
   strSQL = "INSERT INTO HCALADADECAMIONES (DTCONTABLE,IDCAMION,SQCALADA,ICTIPOCALADA,VLHUMEDAD,VLPROTEINA,CDACEPTACION,CDRUBROPPAL,CDGRADO,ICCONDICIONFABRICA,CDUSERNAME,NUBALDE,CDFUERASTD,IDCONTRATO,DTCALADA,DSOBSERVACIONES,NUBARRAS,CDMOTIVORECHAZO,PCMERMA,ICCAMARA,ICHUMEDIMETRO,DSOBSHUMEDIMETRO,CDSUPERVISOR)"&_
            "VALUES('"& p_DtContable &"','"& p_idCamion &"',"& Cdbl(p_UltimaCalada)+1 &",'V',"& p_vlHumedad &","& p_vlPorteina &","& g_Aceptacion &",0,'"& g_cdGrado &"','','"& Session("Usuario") &"',"& g_Balde &","& g_FueraStandart &",0,'"& auxDtCalada &"','"& Trim(p_DsObservaciones) &"','"& g_Sticker &"',"& g_MotivoRechazo &","& p_vlMerma &",'"& g_IcCamara &"','"&g_IcHumedimetro&"','','')"   
   Call logMig.info("Agrega nueva calada: "& strSQL )
   Call GF_BD_Puertos(p_strPuerto, rs, "EXEC",strSQL)
End Function
'---------------------------------------------------------------------------------------------------------
Function actualizarHAcondicionamientoProductos(p_DtContable,p_idCamion,p_UltimaCalada,p_MermaPorc,p_MermaKilo,p_strPuerto)
    Dim strSQL
    strSQL = "INSERT INTO HAcondicProductoCamiones VALUES('"&p_DtContable&"','"&p_idCamion&"',"&Cdbl(p_UltimaCalada)+1&",'FSTDR',"&p_MermaPorc&","&p_MermaKilo&",0)"
    Call logMig.info("Agrega HAcondicProductoCamiones: "& strSQL )
    Call GF_BD_Puertos(p_strPuerto, rs, "EXEC",strSQL)
End function
'---------------------------------------------------------------------------------------------------------
Function controlarRubros(ByRef p_Msj)
    Dim ret,auxCdRubo,auxEstado
    flagSeguir = true
    g_RowActual = 0
    while ((readNextRubroParams())and(flagSeguir))
        if (Cdbl(g_estado) = ESTADO_ACTIVO) then
            if (Trim(g_dsRubro) = "") then
                if (Cdbl(g_cdRubro) <> 0) then
                    if (hayRubroDuplicado(g_cdRubro,g_RowCount, g_RowActual)) then 
                        flagSeguir = false
                        p_Msj = "Un Rubro se encuentra duplicado"
                    end if
                else
                    p_Msj = "Debe seleccionar un Rubro"
                    flagSeguir = false
                end if
            end if
        end if
        Call controlarCambiosRubros()
        g_RowActual = g_RowActual + 1
    wend
    Call controlarCambiosCalada()
    if (flagSeguir) then
        if (oDiccModificaciones.Count = 0) then 
            flagSeguir = false
            p_Msj = "No se encontraron cambios"
        end if
    end if
    controlarRubros = flagSeguir
End Function
'---------------------------------------------------------------------------------------------------------
Function hayRubroDuplicado(p_cdRubro, p_IndexMax, p_Index)
    Dim ret,auxCdRubo,auxEstado
    hayRubroDuplicado = false
    for i = 0 to p_IndexMax
        auxCdRubo = GF_Parametros7("cdRubro_" & i,0,6)
        auxEstado = GF_Parametros7("estado_" & i,0,6)
        if (Cdbl(auxEstado) = ESTADO_ACTIVO) then
            if (Cdbl(p_cdRubro) = Cdbl(auxCdRubo))and(p_Index <> i) then 
                hayRubroDuplicado= true
                
            end if
        end if
    next
End Function
'----------------------------------------------------------------------------------------------------------
Function readNextRubroParams()
    Dim index
	index = PM_DetalleActual
	g_RowCount = GF_PARAMETROS7("rowCount",0,6)
	readNextRubroParams = false
	if Cdbl(g_RowActual) < Cdbl(g_RowCount) then
		g_dsRubro = Trim(GF_PARAMETROS7("dsRubro_" & g_RowActual,"",6))
        'Si tiene descripcion (div) viene ya cargado
        if (g_dsRubro <> "") then
            g_cdRubro = GF_PARAMETROS7("cdRubro_" & g_RowActual,0,6)
        else
            g_cdRubro = GF_PARAMETROS7("cmbRubros_" & g_RowActual,0,6)
        end if
        g_VlRebaja = GF_PARAMETROS7("rebaja_" & g_RowActual,2,6)
        g_VlPeso = GF_PARAMETROS7("vlPeso_" & g_RowActual,2,6)
		g_VlPCPeso = GF_PARAMETROS7("pcPeso_" & g_RowActual,2,6)
        g_VlMerma = GF_PARAMETROS7("merma_" & g_RowActual,2,6)
        g_estado = GF_PARAMETROS7("estado_" & g_RowActual,0,6)
        g_VlRebajaOld = GF_PARAMETROS7("rebajaOld_" & g_RowActual,2,6)
        g_VlPesoOld = GF_PARAMETROS7("vlPesoOld_" & g_RowActual,2,6)
		g_VlPCPesoOld = GF_PARAMETROS7("pcPesoOld_" & g_RowActual,2,6)
        g_VlMermaOld = GF_PARAMETROS7("mermaOld_" & g_RowActual,2,6)
        g_estadoOld = GF_PARAMETROS7("estadoOld_" & g_RowActual,0,6)
		readNextRubroParams = true		
	end if
End Function
'----------------------------------------------------------------------------------------------------------
Function getParameterCalada()
    g_IcCamara = GF_PARAMETROS7("radioCamara","",6)
    g_UltimaCalada = GF_PARAMETROS7("ultimaCalada",0,6)
    g_MotivoRechazo = GF_PARAMETROS7("motivoRechazo","",6)
    g_vlHumedad = GF_PARAMETROS7("vlHumedad","",6)
    g_vlProteina = GF_PARAMETROS7("vlProteina","",6)
    g_PcMerma = GF_PARAMETROS7("vlMerma","",6)
    g_DtCalada = GF_PARAMETROS7("dtCalada","",6)
    g_IcHumedimetro = GF_PARAMETROS7("icHumedimetro","",6)
    g_Sticker = GF_PARAMETROS7("sticker","",6)
    if (g_IcCamara = "S") and (g_Sticker = "") then g_Sticker = getProximoSticker(g_strPuerto, GF_FN2DTCONTABLE(g_dtContable), g_ctaPte, TIPO_TRANSPORTE_CAMION,g_idCamion)    
    g_Aceptacion = GF_PARAMETROS7("cdAceptacion",0,6)
    g_Balde = GF_PARAMETROS7("balde",0,6)
    g_cdGrado = GF_PARAMETROS7("cdGrado",0,6)
    g_FueraStandart = GF_PARAMETROS7("fueraStd",0,6)
    g_DsObservaciones = GF_PARAMETROS7("observaciones","",6)
    g_CdUsuario = GF_PARAMETROS7("cdUsuario","",6)
    g_IcCamaraOld = GF_PARAMETROS7("radioCamaraOld","",6) 
    g_StickerOld = GF_PARAMETROS7("stickerOld","",6)
    g_AceptacionOld = GF_PARAMETROS7("cdAceptacionOld",0,6)
    g_BaldeOld = GF_PARAMETROS7("baldeOld",0,6)
    g_cdGradoOld = GF_PARAMETROS7("cdGradoOld",0,6)
End Function
'----------------------------------------------------------------------------------------------------------
Function loadMailCtaPte(p_ctaPte, p_dtContable, p_strPuerto)
    Dim strSQL,rs
    'Obtengo algunos datos del camion y la descarga para poder enviar el mail
    strSQL = "SELECT HC.NUCUITREM, HCD.CDCORREDOR, HCD.CDVENDEDOR "&_
             "FROM HCAMIONES HC "&_
             "INNER JOIN HCAMIONESDESCARGA HCD "&_
             "      ON HCD.IDCAMION = HC.IDCAMION AND HCD.DTCONTABLE = HC.DTCONTABLE "&_ 
             "WHERE HCD.NUCARTAPORTE = '"& p_ctaPte &"' AND HC.DTCONTABLE = '"& GF_FN2DTCONTABLE(p_dtContable) &"'"
    Call GF_BD_Puertos(p_strPuerto, rs, "OPEN",strSQL)
    if (not rs.Eof) then
        Call cargarIntermediariosCtaPte(rs("NUCUITREM"),g_idCamion,rs("CDVENDEDOR"))
        auxCorredorOld = rs("CDCORREDOR")
        auxDestinatarioOld = auxDestinatario
        g_ctaPteOld = g_ctaPte
        Call enviarMailCtaPte()
    end if
End function
'----------------------------------------------------------------------------------------------------------
Dim g_listaRubrosHumedad,g_listaRubrosZaranda,g_listaRubrosProteina,countResultados,g_UltimaCalada,g_RowCount,g_IcCamara,g_MotivoRechazo,g_vlHumedad,g_vlProteina,g_DtCalada,g_Neto,g_PcMerma
Dim g_DsObservaciones,g_NroAnalisis,g_strSector,g_RowActual,g_cdRubro,g_dsRubro,g_VlRebaja,g_VlPeso,g_VlPCPeso,g_VlMerma,g_estado,g_IcHumedimetro,g_Sticker,g_Aceptacion,g_Balde,g_cdGrado,g_FueraStandart,g_CdUsuario
Dim g_IcCamaraOld,g_StickerOld,g_AceptacionOld,g_BaldeOld,g_cdGradoOld,g_estadoOld,g_VlRebajaOld,g_VlPesoOld,g_VlPCPesoOld,g_VlMermaOld
Dim myPermiso

g_strPuerto = GF_Parametros7("Pto","",6)
g_dtContable = GF_Parametros7("dtContable","",6)
g_ctaPte = GF_Parametros7("ctaPte","",6)
g_ctaPteOld = g_ctaPte 
g_idCamion = GF_Parametros7("idCamion","",6)
auxGranoOld = GF_Parametros7("cdProducto",0,6)
g_UltimaCalada = GF_Parametros7("sqCalada",0,6)
g_MermaOld = GF_Parametros7("mermaOld",0,6)
auxDestinatario = GF_Parametros7("destinatario",0,6)
accion = GF_Parametros7("accion","",6)


Call initTaskAccessInfo(TASK_POS_MODIFICACION_HISTORICA, session("DIVISION_PUERTO"))

if (accion = ACCION_SUBMITIR) then    
    Call getParameterCalada()
    flagControl = controlarRubros(msjError)
    if (flagControl) then
        Call inicializarLogCartaPorte()
        'Obtengo los kilos bruto Y tara(no se modifican en esta pagina)
        auxPesadaBruto = getKilosPesadaCamion(g_idCamion, g_dtContable, PESADA_BRUTO, g_strPuerto)
        auxPesadaTara  = getKilosPesadaCamion(g_idCamion, g_dtContable, PESADA_TARA, g_strPuerto)
        Call grabarRubro(g_dtContable,g_idCamion,g_strPuerto,g_UltimaCalada,auxGranoOld,g_DsObservaciones,auxPesadaBruto,auxPesadaTara)
        Call agregarRegistrosAuditoria(g_dtContable,g_idCamion,g_UltimaCalada,g_strPuerto)
        Call agregarGrupoEnsayos(g_dtContable,g_idCamion,g_UltimaCalada,g_strPuerto)
        Call agregarMuestraHumedad(g_dtContable,g_idCamion,g_UltimaCalada,g_strPuerto)
        auxNroAjuste = actualizarAjusteCamion(g_idCamion,g_dtContable,auxGranoOld,auxGranoOld,auxDestinatario,auxDestinatario,g_Merma,g_MermaOld,auxPesadaBruto,auxPesadaBruto,auxPesadaTara,auxPesadaTara)
        Call rearmarCartaPorte(g_idCamion,g_dtContable,auxGranoOld,auxGranoOld,auxDestinatario,auxDestinatario,g_Merma,g_MermaOld,auxPesadaBruto,auxPesadaBruto,auxPesadaTara,auxPesadaTara)
        Call rearmarStockFisico(g_idCamion,g_dtContable,auxGranoOld,auxGranoOld,auxDestinatario,auxDestinatario,g_Merma,g_MermaOld,auxPesadaBruto,auxPesadaBruto,auxPesadaTara,auxPesadaTara)

        'Preparo lo datos para envio de mails.
        Call loadMailCtaPte(g_ctaPte, g_dtContable, g_strPuerto)
        g_MermaOld = g_Merma
    end if
    oDiccModificaciones.RemoveAll
end if

if (accion = "" or flagControl) then Call loadCaladasCamion(g_dtContable, g_idCamion,g_strPuerto)
g_RowActual = 0

%>
<HTML>
<HEAD>
	<TITLE>Poseidon - Informacion de C�lidad de Camion </TITLE>
    <link rel="stylesheet" type="text/css" href="../css/ActiSAIntra-1.css">	
    <link rel="stylesheet" type="text/css" href="../css/main.css">
    <link rel="stylesheet" type="text/css" href="../css/toolbar.css">
    <link rel="stylesheet" type="text/css" href="../css/Header.css">
    <link rel="stylesheet" href="../css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css">

	<script type="text/javascript" src="../scripts/channel.js"></script>
    <script type="text/javascript" src="../scripts/formato.js"></script>
    <script type="text/javascript" src="../scripts/controles.js"></script>
    <script type="text/javascript" src="../scripts/jQueryPopUp.js"></script>
    <script type="text/javascript" src="../scripts/jquery/jquery-1.3.2.min.js"></script>
    <script type="text/javascript" src="../scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
    <script type="text/javascript" src="../scripts/Toolbar.js"></script>
    <script type="text/javascript" src="../scripts/date.js"></script>
    <script type="text/javascript" src="../scripts/calendar.js"></script>
    <script type="text/javascript" src="../scripts/calendar-1.js"></script>
	<SCRIPT LANGUAGE="JavaScript">
	    var isFirefox = !(navigator.appName == "Microsoft Internet Explorer");		
        var ch= new channel();
		function bodyOnload() {
		    parent.dimensionarIframe(750, 650);
		    var tb = new Toolbar('toolbar');
		    tb.addButton("toolbar-save", "Guardar", "saveCtaPte()");
		    tb.draw();
            <% if (accion = ACCION_SUBMITIR) then %>
                <% if (not flagControl) then %>
                    document.getElementById("errCtaPte").innerHTML = "<%=msjError %>";
                    document.getElementById("errCtaPte").className = "reg_Header_Error";
	                document.getElementById("errCtaPte").style.width = "100%";
                 <% while(readNextRubroParams()) %>
                      <% if (g_dsRubro = "") then %>
                           AddLineRubro("<%= g_cdRubro%>");   
                      <% end if %>
                      fetchRubros("<%= g_cdRubro%>","<%= g_dsRubro%>","<%=g_VlRebaja %>","<%=g_VlPeso %>","<%=g_VlPCPeso %>","<%=g_VlMerma %>","<%=g_estado %>","<%=g_RowActual %>");
                <%    g_RowActual = g_RowActual + 1
                    wend %>
                <% else %>
                    document.getElementById("errCtaPte").innerHTML = "Se ha guardado correctamente";
                    document.getElementById("errCtaPte").className = "reg_Header_success";
	                document.getElementById("errCtaPte").style.width = "100%";
                <% end if %>
            <% end if %>
		}
		function saveCtaPte() {
		    p_Indice = document.getElementById("rowCount").value;
            if (p_Indice != 0) {                
                    document.getElementById("accion").value = "<%=ACCION_SUBMITIR %>";
                    document.getElementById("frmSel").submit();
            }
            else {
                alert("No se puede generar una nueva calada sin Rubros");
            }
        }
        function fetchRubros(p_CdRubro, p_DsRubro, p_VlRebaja, p_vlPeso, p_vlPcPeso, p_VlMerma, p_Estado, p_Index){
            document.getElementById("rebaja_" + p_Index).value = p_VlRebaja;
            document.getElementById("vlPeso_" + p_Index).value = p_vlPeso;
            document.getElementById("pcPeso_" + p_Index).value = p_vlPcPeso;
            document.getElementById("merma_" + p_Index).value = p_VlMerma;
            document.getElementById("divMerma_" + p_Index).innerHTML = p_VlMerma;
            if (document.getElementById("dsRubro_" + p_Index)) {
                document.getElementById("dsRubro_" + p_Index).value = p_DsRubro;
                document.getElementById("cdRubro_" + p_Index).value = p_CdRubro;
            }
            document.getElementById("estado_" + p_Index).value = p_Estado;
        }
		
		function removeAllChilds(a){			
			while(a.hasChildNodes()){
				a.removeChild(a.firstChild);
			}	
		}
		function lightOn(tr) {
			tr.className = "reg_Header_navdosHL";
		}
		function lightOff(tr) {
			tr.className = "reg_Header_navdos";
		}
		function AddLineRubro(p_CdRubro) {
		    var index = document.getElementById("rowCount").value;		    
		    var tr_1 = document.createElement('tr');
            tr_1.id = "tr_" + index;
		    var td_Cab_1 = document.createElement('td');		    
            var div_Cab_1 = document.createElement('div');
		    div_Cab_1.id = "divCmb_" + index;
		    div_Cab_1.name = "divCmb_" + index;
		    ch.bind("cartadeporteRubrosAjax.asp?pto=<%=g_strPuerto%>&idCamion=<%=g_idCamion %>&dtContable=<%=g_dtContable %>&ctaPte=<%=g_ctaPte%>&cdRubro="+p_CdRubro, "CallBack_verCalada(" + index + ")");
		    ch.send();
		    td_Cab_1.appendChild(div_Cab_1);
            tr_1.appendChild(td_Cab_1);
            var hid_Cab_1 = document.createElement('input');
		    hid_Cab_1.type = "hidden";
		    hid_Cab_1.id = "cdRubroOld_" + index;
		    hid_Cab_1.name = "cdRubroOld_" + index;
		    hid_Cab_1.value = 0;
            td_Cab_1.appendChild(hid_Cab_1);
            tr_1.appendChild(td_Cab_1);
		    var td_Cab_2 = document.createElement('td');
		    td_Cab_2.align = "center";
		    var txt_Cab_1 = document.createElement('input');
		    txt_Cab_1.type = "text";
		    txt_Cab_1.id = "rebaja_" + index;
		    txt_Cab_1.name = "rebaja_" + index;
		    txt_Cab_1.value = 0;
		    txt_Cab_1.style.textAlign = "right";
		    txt_Cab_1.size = 6;
		    txt_Cab_1.maxlength = 7;
		    if (isFirefox) {
		        txt_Cab_1.setAttribute('onkeypress', "return controlIngreso(this, event, 'N')");
		    } else {
		        txt_Cab_1['onkeypress'] = new Function("return controlIngreso(this, event, 'N')");
		    }
		    if (isFirefox) {
		        txt_Cab_1.setAttribute('onblur', "validarRubroMerma("+ index +")");
		    } else {
		        txt_Cab_1['onblur'] = new Function("validarRubroMerma("+ index +")");
		    }
		    td_Cab_2.appendChild(txt_Cab_1);
		    tr_1.appendChild(td_Cab_2);
		    var td_Cab_3 = document.createElement('td');
		    td_Cab_3.align = "center";
		    var txt_Cab_2 = document.createElement('input');
		    txt_Cab_2.type = "text";
		    txt_Cab_2.id = "vlPeso_" + index;
		    txt_Cab_2.name = "vlPeso_" + index;
		    txt_Cab_2.value = 0
		    txt_Cab_2.style.textAlign = "right";
		    txt_Cab_2.size = 6;
		    txt_Cab_1.maxlength = 9;
		    if (isFirefox) {
		        txt_Cab_2.setAttribute('onkeypress', "return controlIngreso(this, event, 'N')");
		    } else {
		        txt_Cab_2['onkeypress'] = new Function("return controlIngreso(this, event, 'N')");
		    }
		    td_Cab_3.appendChild(txt_Cab_2);
		    tr_1.appendChild(td_Cab_3);
		    var td_Cab_4 = document.createElement('td');
		    td_Cab_4.align = "center";
		    var spa_Cab_1 = document.createElement('input');
		    spa_Cab_1.id = "pcPeso_" + index;
		    spa_Cab_1.name = "pcPeso_" + index;
		    spa_Cab_1.value = 0;
		    spa_Cab_1.style.textAlign = "right";
		    spa_Cab_1.type = "text";
		    spa_Cab_1.size = 6;
		    txt_Cab_1.maxlength = 7;
		    if (isFirefox) {
		        spa_Cab_1.setAttribute('onkeypress', "return controlIngreso(this, event, 'N')");
		    } else {
		        spa_Cab_1['onkeypress'] = new Function("return controlIngreso(this, event, 'N')");
		    }
		    td_Cab_4.appendChild(spa_Cab_1);
		    tr_1.appendChild(td_Cab_4);
		    var td_Cab_5 = document.createElement('td');
		    td_Cab_5.align = "center";
		    var txt_Cab_3 = document.createElement('input');
		    txt_Cab_3.type = "hidden";
		    txt_Cab_3.id = "merma_" + index;
		    txt_Cab_3.name = "merma_" + index;
		    txt_Cab_3.value = 0;
		    txt_Cab_3.style.textAlign = "right";
		    txt_Cab_1.maxlength = 7;
		    if (isFirefox) {
		        txt_Cab_3.setAttribute('onkeypress', "return controlIngreso(this, event, 'N')");
		    } else {
		        txt_Cab_3['onkeypress'] = new Function("return controlIngreso(this, event, 'N')");
		    }
		    td_Cab_5.appendChild(txt_Cab_3);
		    var div_Cab_2 = document.createElement('div');
		    div_Cab_2.id = "divMerma_" + index;
		    div_Cab_2.name = "divMerma_" + index;
		    div_Cab_2.innerHTML = "0";
		    div_Cab_2.style.textAlign = "right";
		    td_Cab_5.appendChild(div_Cab_2);
		    tr_1.appendChild(td_Cab_5);
		    var td_Cab_6 = document.createElement('td');
		    td_Cab_6.align = "center";
		    var img_3 = document.createElement('img');
		    img_3.id = "delete_" + index;
		    img_3.name = "delete_" + index;
		    img_3.src = "../images/icon_del.gif";
		    img_3.title = "Borrar";
		    img_3.style.cursor = "pointer";
		    if (isFirefox) {
		        img_3.setAttribute('onclick', "limpiarFila(" + index + ")");
		    } else {
		        img_3['onclick'] = new Function("limpiarFila(" + index  + ")");
		    }
		    td_Cab_6.appendChild(img_3);
		    var hiddenEstado = document.createElement('input');
		    hiddenEstado.type = "hidden";
            hiddenEstado.id = "estado_" + index;
		    hiddenEstado.name = "estado_" + index;
		    hiddenEstado.value = "<%= ESTADO_ACTIVO %>";		    
		    tr_1.appendChild(td_Cab_6);
            var hiddenEstadoOld = document.createElement('input');
		    hiddenEstadoOld.type = "hidden";
            hiddenEstadoOld.id = "estadoOld_" + index;
		    hiddenEstadoOld.name = "estadoOld_" + index;
		    hiddenEstadoOld.value = "<%= ESTADO_BAJA %>";		    		    
		    $("#tblRubros tbody").append(tr_1);
		    $("#tblRubros tbody").append(hiddenEstado);
            $("#tblRubros tbody").append(hiddenEstadoOld);
		    //document.getElementById("addRubro").style.display = "none";
		    index = parseInt(index) + 1
		    document.getElementById("rowCount").value = index;
		}
		function CallBack_verCalada(p_Index) {
		    var respuesta = ch.response();
		    var elem = document.getElementById("divCmb_" + p_Index).innerHTML = respuesta;
		    $("#divCmb_" + p_Index).each(function () {
		        var id = "#" + this.id;
		        $(id + " select").attr("id", "cmbRubros_" + p_Index);
		        $(id + " select").attr("name", "cmbRubros_" + p_Index);
		        var selectRubro = document.getElementById("cmbRubros_" + p_Index);
		        if (isFirefox) {
		            selectRubro.setAttribute('onchange', "validarRubroMerma("+ p_Index +")");
		        } else {
		            selectRubro['onchange'] = new Function("validarRubroMerma("+ p_Index +")");
		        }
		    })

		}
		function limpiarFila(pIndex) {
		    if (hayRubroActivos(pIndex) == true) {
		        document.getElementById("tr_" + pIndex).style.display = "none";
                document.getElementById("estado_" + pIndex).value = "<%= ESTADO_BAJA %>";
            }
            else
            {
                alert("La nueva calada debe tener rubros")
            }
        }
        //Permite saber si hay rubros activos (no eliminados) debido a que no puede ser 0 la cantidad de Rubors en la secuencia de calada
        function hayRubroActivos(pIndex) {
            var maxIndice = document.getElementById("rowCount").value;
            var ret = false;
            for (var i = 0; i < maxIndice; i++) {
                if ((document.getElementById("estado_" + i).value == "<%=ESTADO_ACTIVO %>")&&(pIndex != i)) ret = true;
            }
            return ret;
        }        
        function guardarCambios(p_Indice) {
            if (confirm("Si elimina el Rubro se generara una nueva calada, desea continuar?")) {
                document.getElementById("accion").value = "<%=ACCION_SUBMITIR %>";
                document.getElementById("selRubro").value = p_Indice;
                document.getElementById("frmSel").submit();
            }            
        }
        function obtenerMerma(p_Index, p_CdRubro){
            var vlRebaja = document.getElementById("rebaja_" + p_Index).value;
            ch.bind("cartadeporteEditAjax.asp?pto=<%=g_pto %>&cdProducto=<%=auxGranoOld%>&accion=<%=ACCION_CALCULAR%>&valorRebaja="+vlRebaja+"&cdRubro="+p_CdRubro, "obtenerMerma_callBack("+ p_Index +")");
            ch.send();
        }
        function obtenerMerma_callBack(p_Index){
            var rtrn = ch.response()
            document.getElementById("merma_"+p_Index).value = rtrn;
            document.getElementById("divMerma_"+p_Index).innerHTML = rtrn;
        }
        function validarRubroMerma(p_Index){
            var cdRubro = document.getElementById("cmbRubros_" + p_Index).value;
            if (cdRubro != 0) obtenerMerma(p_Index,cdRubro)
        }
    </script>
</HEAD>	
<BODY onload="bodyOnload();">
<div id="toolbar"></div>
<form id="frmSel" name="frmSel" method=post action="cartadeportePopUpRubros.asp">
	<div class="col66"></div>
    <table width="100%" align="center"><tr><td><div id="errCtaPte" ></div></td></tr></table>
    <div class="tableasidecontent">
        
        <div class="col26 reg_header_navdos"> Fecha Descarga </div>
        <div class="col26"> <% =GF_FN2DTE(g_dtContable) %>  </div>
        
        <div class="col26 reg_header_navdos"> ID Camion </div>
        <div class="col26">  <% =g_idCamion %> </div>
        
        <div class="col26 reg_header_navdos"> Carta Porte </div>
        <div class="col26"> <% =GF_EDIT_CTAPTE(g_ctaPte) %>  </div>
        
    </div>
	<div class="col66"></div>
	
    <div class="tableaside size100">
		<h3>Datos Calada</h3>
			<table class="datagrid datagridlv1" width="98%">
			    <thead>
        	        <tr> 
        		        <th>&nbsp;</th>       	    
            	        <th colspan="2"><%=GF_Traducir("Datos generales")%></th>
            	        <th>&nbsp;</th>
                    </tr>
                </thead>
    	        <tbody>
                    <tr>
                        <td width="25%"><b><%=GF_Traducir("Ultima Calada")%>:</b></td>
                        <td width="25%"><input id="ultimaCalada" name="ultimaCalada" type="hidden" value="<%=g_UltimaCalada%>" /><%=g_UltimaCalada%></td>
                        <td width="25%"><b><%=GF_Traducir("Tipo Calada")%>.:</b></td>
                        <td width="25%"><input id="motivoRechazo" name="motivoRechazo" type="hidden" value="<%=g_MotivoRechazo %>" /><%=g_MotivoRechazo %></td>
                    </tr>
                    <tr>
                        <td width="25%"><b><%=GF_Traducir("Humedad")%>:</b></td>
                        <td width="25%"><input id="vlHumedad" name="vlHumedad" type="hidden" value="<%=g_vlHumedad %>" /><%=g_vlHumedad %></td>
                        <td width="25%"><b><%=GF_Traducir("Proteina")%>:</b></td>
                        <td width="25%"><input id="vlProteina" name="vlProteina" type="hidden" value="<%=g_vlProteina %>" /><%=g_vlProteina %></td>
                    </tr>
                    <tr>                
                        <td><b><%=GF_Traducir("Merma")%>.:</b></td>
                        <td><input id="vlMerma" name="vlMerma" type="hidden" value="<%=g_PcMerma %>" /><%=g_PcMerma %></td>
                        <td><b><%=GF_Traducir("Dt calada")%>.:</b></td>
                        <td><input id="dtCalada" name="dtCalada" type="hidden" value="<%=g_DtCalada %>" /><%=g_DtCalada %></td>
                    </tr>
                    <tr>                
                        <td><b><%=GF_Traducir("Camara")%>.:</b></td>
                        <td>
                            <input type="radio" name="radioCamara" id="radioCamara" value="S" <%if (g_IcCamara = "S") then %>checked="checked"<%end if%> /><% = GF_TRADUCIR("Si")%>&nbsp&nbsp&nbsp
                            <input type="radio" name="radioCamara" id="radioCamara" value="N" <%if (g_IcCamara = "N") then %>checked="checked"<%end if%> /><% = GF_TRADUCIR("No")%>
                            <input type="hidden" id="radioCamaraOld" name="radioCamaraOld" value=<%=g_IcCamaraOld %>>
                        </td>
                        <td><b><%=GF_Traducir("Humedimetro")%>.:</b></td>
                        <td><input type="hidden" name="icHumedimetro" id="icHumedimetro" value="<%=g_IcHumedimetro %>" /><%=g_IcHumedimetro %></td>
                    </tr>
                    <tr>                
                        <td><b><%=GF_Traducir("Sticker")%>.:</b></td>
                        <td><input id="sticker" name="sticker" type="text" value="<%=g_Sticker %>" size="14" disabled />
                            <input type="hidden" id="stickerOld" name="stickerOld" value="<%=g_StickerOld %>"/>
                        </td>                        
                        <td><b><%=GF_Traducir("Aceptacion")%>.:</b></td>
                        <td>
                            <select id="cdAceptacion" name="cdAceptacion">
                        <%  Call GF_BD_Puertos(g_strPuerto, rsAce, "OPEN","SELECT * FROM ACEPTACIONCALIDAD") 
                            while (not rsAce.Eof) %>
                                <option value="<%=rsAce("CDACEPTACION") %>" <% if(Cdbl(rsAce("CDACEPTACION")) = Cdbl(g_Aceptacion))then %>selected<% end if %>><%=rsAce("CDACEPTACION")&"-"&rsAce("DSACEPTACION") %></option>
                        <%      rsAce.MoveNext()
                            wend %>                            
                            </select>
                            <input type="hidden" id="cdAceptacionOld" name="cdAceptacionOld" value="<%=g_AceptacionOld %>"/>
                        </td>
                    </tr>                    
				    <tr>                
                        <td><b><%=GF_Traducir("Balde")%>.:</b></td>
                        <td><input id="balde" name="balde" type="text" value="<%=g_Balde %>" size="14"/>
                            <input type="hidden" id="baldeOld" name="baldeOld" value="<%=g_BaldeOld %>"/>
                        </td>
                        <td><b><%=GF_Traducir("Grado")%>.:</b></td>
                        <td>
                           <select id="cdGrado" name="cdGrado">
                           <option value=0>Seleccione..</option>  
                        <%  Call GF_BD_Puertos(g_strPuerto, rsGra, "OPEN","SELECT * FROM GRADOS")
                            while (not rsGra.Eof) %>
                                <option value="<%=Cdbl(rsGra("CDGRADO")) %>" <% if(Cdbl(rsGra("CDGRADO")) = Cdbl(g_cdGrado))then %>selected<% end if %>><%=rsGra("CDGRADO")&"-"&rsGra("DSGRADO") %></option>
                        <%      rsGra.MoveNext()
                            wend %>
                           </select>
                           <input type="hidden" id="cdGradoOld" name="cdGradoOld" value="<%=g_cdGrado %>"/>
                        </td>

                    </tr>
                     <tr>
                        <td><b><%=GF_Traducir("Fuera Std")%>.:</b></td>
                        <td><input id="fueraStd" name="fueraStd" type="hidden" value="<%=g_FueraStandart %>" /><%=g_FueraStandart %></td>
                        <td><b><%=GF_Traducir("Usuario")%>.:</b></td>
                        <td><input id="cdUsuario" name="cdUsuario" type="hidden" value="<%=g_CdUsuario %>" /><%=g_CdUsuario %></td>
                    </tr>
                    <tr>
                        <td><b><%=GF_TRADUCIR("Observaciones")%>.:</b></td>
                        <td colspan="3"><input id="observaciones" name="observaciones" type="text" size="60" maxlength="100" value="<%=g_DsObservaciones%>" /></td>
                    </tr>
                </tbody>
                <input type="hidden" id="sqCalada" name="sqCalada" value="<%=g_UltimaCalada %>"/>
			</table>	  
		<br>
	</div>
    <!-- Si no tiene ninguna Calada, no tendria que tener Rubros (debido que en la tabla de RubrosVisteosCamiones se utiliza la SQCALADA) -->
    <% if (Cdbl(g_UltimaCalada) <> 0)then %>
    <div class="tableaside size100">
		<h3>Datos Rubros</h3>
			<table class="datagrid datagridlv1" width="98%" id="tblRubros">
                <thead>
        	        <tr> 
        		        <th width="35%"><%=GF_Traducir("Rubro")%></th>
            	        <th width="15%"><%=GF_Traducir("Rebaja")%></th>
            	        <th width="15%"><%=GF_Traducir("Peso rubro")%></th>
                        <th width="15%"><%=GF_Traducir("PC peso rubro")%></th>
                        <th width="15%"><%=GF_Traducir("Merma")%></th>
                        <th width="5%">.</th>
                    </tr>
                </thead>
    	        <tbody>
                <%Set g_rsRubros = getRubrosVisteoCamion(g_dtContable, g_idCamion, g_UltimaCalada, g_strPuerto) %>
		        <%index = 0
                if not g_rsRubros.eof then
                while not g_rsRubros.Eof %>
                    <tr id="tr_<%=index %>">
                        <td align="left"><%=g_rsRubros("CDRUBRO")&"-"&g_rsRubros("DSRUBRO")%>
                            <input type="hidden" id="dsRubro_<%=index %>" name="dsRubro_<%=index %>" value="<%=g_rsRubros("CDRUBRO")&"-"&g_rsRubros("DSRUBRO")%>" />
                            <input type="hidden" id="cdRubro_<%=index%>" name="cdRubro_<%=index%>" value="<%=g_rsRubros("CDRUBRO") %>" />
                            <input type="hidden" id="cdRubroOld_<%=index%>" name="cdRubroOld_<%=index%>" value="<%=g_rsRubros("CDRUBRO") %>" />
                        </td>
                        <td align="center">
                            <input type="text" id="rebaja_<%=index%>" name="rebaja_<%=index%>" value="<%=g_rsRubros("VLBONREBAJA")%>" onKeyPress="return controlIngreso (this, event, 'N');" onblur="obtenerMerma(<%=index%>,document.getElementById('cdRubro_<%=index %>').value)" size="6" style="text-align:right;" maxlength="7" />
                            <input type="hidden" id="rebajaOld_<%=index%>" name="rebajaOld_<%=index%>" value="<%=g_rsRubros("VLBONREBAJA")%>"/>
                        </td>                        
                        <td align="center">
                            <input type="text" id="vlPeso_<%=index%>" name="vlPeso_<%=index%>" value="<%=g_rsRubros("VLPESORUBRO")%>" onKeyPress="return controlIngreso (this, event, 'N');" size="6" style="text-align:right;" maxlength="9"/>
                            <input type="hidden" id="vlPesoOld_<%=index%>" name="vlPesoOld_<%=index%>" value="<%=g_rsRubros("VLPESORUBRO")%>"/>
                        </td>
                        <td align="center">
                            <input type="text" id="pcPeso_<%=index%>" name="pcPeso_<%=index%>" value="<%=g_rsRubros("PCPESORUBRO")%>" onKeyPress="return controlIngreso (this, event, 'N');" size="6" style="text-align:right;" maxlength="7"/>
                            <input type="hidden" id="pcPesoOld_<%=index%>" name="pcPesoOld_<%=index%>" value="<%=g_rsRubros("PCPESORUBRO")%>"/>
                        </td>
                        <td align="center">
                            <div id="divMerma_<%=index%>" style="float:right;"><%= g_rsRubros("VLMERMA")%></div>
                            <input type="hidden" id="merma_<%=index%>" name="merma_<%=index%>" value="<%=g_rsRubros("VLMERMA")%>" size="6" style="text-align:right;" maxlength="7" />
                            <input type="hidden" id="mermaOld_<%=index%>" name="mermaOld_<%=index%>" value="<%=g_rsRubros("VLMERMA")%>"/>
                        </td>
                        <td align="center">
                            <img src="../images/icon_del.gif" style="cursor:pointer;" id="delete_<%=index %>" onclick="limpiarFila(<%=index %>)" title="Eliminar"/>
                        </td>
                    </tr>
                    <input type="hidden" id="estado_<%=index %>" name="estado_<%=index %>" value="<%=ESTADO_ACTIVO %>"/>
                    <input type="hidden" id="estadoOld_<%=index %>" name="estadoOld_<%=index %>" value="<%=ESTADO_ACTIVO %>"/>
                 <%  index = index + 1
                     g_rsRubros.MoveNext()
                  wend%>
                <% else %>	
                    <tr><td colspan="6" align="center">No se encontraron Rubros</td></tr>
                <%end if%>
                </tbody>
                <input type="hidden" id="rowCount" name="rowCount" value="<%=index%>"/>
                <tfoot>
                    <tr>
                        <td colspan="5" align="right"></td>
                        <td align="center"><img id="addRubro" src="../images/add.gif" style="cursor:pointer;" onclick="AddLineRubro(0);" title="Agregar" /></td>
                    </tr>
                </tfoot>
			</table>
		<br>
	</div>	
    <% end if %>
    <input type="hidden" id="Pto" name="Pto" value="<%=g_strPuerto %>"/>
    <input type="hidden" id="dtContable" name="dtContable" value="<%=g_dtContable %>"/>
    <input type="hidden" id="ctaPte" name="ctaPte" value="<%=g_ctaPte %>"/>
    <input type="hidden" id="idCamion" name="idCamion" value="<%=g_idCamion %>"/>
    <input type="hidden" id="accion" name="accion" value="<%=accion%>"/>
    <input type="hidden" id="cdProducto" name="cdProducto" value="<%=auxGranoOld%>"/>
    <input type="hidden" id="selRubro" name="selRubro"/>
    <input type="hidden" id="destinatario" name="destinatario" value="<%=auxDestinatario%>"/>
    <input type="hidden" id="mermaOld" name="mermaOld" value="<%=g_MermaOld%>"/>
    </form>
</BODY>
</HTML>
<%

%>