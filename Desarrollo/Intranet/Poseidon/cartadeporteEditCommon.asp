<!--#include file="../Includes/procedimientostraducir.asp"-->
<!--#include file="../Includes/procedimientosFechas.asp"-->
<!--#include file="../Includes/procedimientosSQL.asp"-->
<!--#include file="../Includes/procedimientosCompras.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosPuertos.asp"-->
<!--#include file="../Includes/procedimientosUser.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->
<!--#include file="../Includes/procedimientosMail.asp"-->
<!--#include file="../Includes/procedimientosLaboratorio.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientosValidacion.asp"-->
<!--#include file="../Includes/procedimientosLog.asp"-->
<!--#include file="../Includes/procedimientosSeguridad.asp"-->
<%
Const INTERVINIENTE_TITULAR = 1
Const INTERVINIENTE_INTERMEDIARIO = 2
Const INTERVINIENTE_REMITENTE = 3
Const INTERVINIENTE_CORREDOR = 4
Const INTERVINIENTE_ENTREGADOR = 5
Const INTERVINIENTE_DESTINATARIO = 6
Const INTERVINIENTE_DESTINO = 7
Const INTERVINIENTE_TRANSPORTISTA = 8
Const INTERVINIENTE_CHOFER = 9

Const CUENTAYORDEN_INTERMEDIARIO = 1
Const CUENTAYORDEN_REMITENTE = 2

Const TRANSACCION_MODIFICACION_HISTORICA = 12

Const PARAM_MAX_TARA  = "VLMAXTARA"
Const PARAM_MAX_BRUTO = "VLMAXBRUTO"
Const PARAM_MIN_TARA  = "VLMINTARA"
Const PARAM_MIN_BRUTO = "VLMINBRUTO"

CONST PARAM_CD_RUBRO_HUMEDAD  = "CDRUBROHUMEDAD"
CONST PARAM_CD_RUBRO_ZARANDA  = "CDRUBROZARANDA"

Const PARAM_ESTADO_BAJA      = "CDESTADOBAJA"
Const PARAM_ESTADO_EGRERECH  = "CDESTADOEGRRECH"
Const PARAM_ESTADO_RECHAZADO = "CDESTADORECHAZO"

Const OPERATOR_SUMA = "+"
Const OPERATOR_RESTA = "-"

'-----------------------------------------------------------------------------------------------------------------------------------
Function addParam(p_strKey,p_strValue,ByRef p_strParam)
       if (not isEmpty(p_strValue)) then
          if (isEmpty(p_strParam)) then
             p_strParam = "?"
          else
             p_strParam = p_strParam & "&"
          end if
          p_strParam = p_strParam & p_strKey & "=" & p_strValue
       end if
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function controlarCabeceraCartaPorte()    
    Dim msj
    msj = ""
    if (Trim(auxCartaPorte1) <> "")and(Trim(auxCartaPorte2) <> "") then
        if (Len(Trim(auxCartaPorte1)) = 4)and(Len(Trim(auxCartaPorte2)) = 8) then
            if (controlarDuplicidadCartaPorte(Trim(auxCartaPorte1) & Trim(auxCartaPorte2),g_ctaPte)) then
                if (Trim(auxCTG) <> "") then
                    if (auxDtCartaPorte <> "") then
                        if (auxDtVencimiento <> "") then
                            if (GF_CONTROL_PERIODO(Right(auxDtCartaPorte,2),Right(auxDtVencimiento,2),Mid(auxDtCartaPorte,5,2),Mid(auxDtVencimiento,5,2),Left(auxDtCartaPorte,4),Left(auxDtVencimiento,4)) <> 0) then msj = "Error en el periodo de fecha de Carga y de Vencimiento"
                        else
                            msj = "Se deben completar la fecha de Vencimiento"
                        end if
                    else
                        msj = "Se deben completar la fecha de Carga"
                    end if
                else
                    msj = "El CTG se encuentra vacio"
                end if
            else
                msj = "Se encontro duplicada la carta de porte"    
            end if
        else
            msj = "La carta de porte no cumple con el formato adecuado"
        end if
    else
        msj = "La carta de porte se encuentra vacia"
    end if
    controlarCabeceraCartaPorte = msj
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function controlarIntervinientesCartaPorte()
    Dim msj
    msj = ""
    'Controlo que por lo menos tenga 1 interviniente de procedencia (Titular o Intermediario o Remitente)
    if (controlarIntervinienteObligatorio(auxTitularCd,auxTitularCuit1,auxTitularCuit2,auxTitularCuit3)) then
        if (controlarInterviniente(auxRemitenteCd,auxRemitenteCuit1, auxRemitenteCuit2, auxRemitenteCuit3)) then
            if (controlarInterviniente(auxIntermediarioCd,auxIntermediarioCuit1, auxIntermediarioCuit2, auxIntermediarioCuit3)) then
                if (controlarIntervinienteObligatorio(auxCorredor, auxCorredorCuit1, auxCorredorCuit2, auxCorredorCuit3)) then
                    if (controlarIntervinienteObligatorio(auxEntregador, auxEntregadorCuit1, auxEntregadorCuit2, auxEntregadorCuit3)) then
                        if (controlarIntervinienteObligatorio(auxDestinatario, auxDestinatarioCuit1, auxDestinatarioCuit2, auxDestinatarioCuit3)) then
                            if (controlarIntervinienteObligatorio(auxTransportista, auxTransportistaCuit1, auxTransportistaCuit2, auxTransportistaCuit3)) then
                               if (Not controlarChofer(auxChoferNumDoc1 & auxChoferNumDoc2 & auxChoferNumDoc3)) then  msj = "No se encontro el Chofer o presenta un error en el formato del CUIT"
                            else
                                msj = "No se encontro el Transportista"
                            end if
                        else
                            msj = "No se encontro el Destinatario"
                        end if
                    else
                        msj = "No se encontro el Entregador"
                    end if
                else
                    msj = "No se encontro el Corredor"
                end if
            else
                msj = "Error en el CUIT del Itermediario"
            end if        
        else
            msj = "Error en el CUIT del Remitente"
        end if
    else
        msj = "No se encontro el Titular("& auxTitularCd&" "&auxTitularCuit1&" "&auxTitularCuit2&" "&auxTitularCuit3&")<BR>"
    end if
    controlarIntervinientesCartaPorte = msj
End Function
'---------------------------------------------------------------------------------------------------------------------------------
Function controlarChofer(p_Cuit)
        if ((NOT GF_CONTROL_CUIT(p_Cuit)) OR (auxChoferTipoDoc = "0" )) then
            controlarChofer = false
        else
            controlarChofer = true
        end if
End Function 
'-----------------------------------------------------------------------------------------------------------------------------------
'Controla que el cuit se encuentre obligatoriamente y que su formato sea correcto, ademas controla que tenga un codigo asociado
'(salvo que el Titular se encuentre en la tabla de Buenos Aires y no tenga codigo de vendedor, en ese caso se asigna -1 para que pase el control)
Function controlarIntervinienteObligatorio(p_CdInterviniente, p_Cuit_1, p_Cuit_2, p_Cuit_3)
    controlarIntervinienteObligatorio = false
    if (GF_CONTROL_CUIT(p_Cuit_1 & p_Cuit_2 & p_Cuit_3)) then 
          if (Cdbl(p_CdInterviniente) <> 0)then controlarIntervinienteObligatorio = true
    end if
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function controlarInterviniente(p_CdInterviniente,p_Cuit_1, p_Cuit_2, p_Cuit_3)
    controlarInterviniente = true
    if ((Trim(p_Cuit_1) <> "")or(Trim(p_Cuit_2) <> "")or(Trim(p_Cuit_3) <> "")) then
        if (not GF_CONTROL_CUIT(p_Cuit_1 & p_Cuit_2 & p_Cuit_3)) then 
            controlarInterviniente= false
        else
            if (Cdbl(p_CdInterviniente) = 0) then controlarInterviniente= false
        end if
    end if
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function controlarProductoCartaPorte()
    Dim msj
    msj = ""
    if (auxCosecha <> "") then
        if (Len(auxCosecha) = 8) then
            if (auxGrano <> 0) then
                if (auxPesoBruto <> 0)and(Trim(auxPesoBruto) <> "") then
                    if (auxPesoTara <> 0)and(Trim(auxPesoTara) <> "") then
                        if (auxPesoNeto <> 0)and(Trim(auxPesoNeto) <> "") then
                            'if (Trim(auxCupo) <> "") then    
                                if (Cdbl(auxProcedenciaProv) <> 0) then
                                    if (Cdbl(auxProcedenciaCd) = 0) then msj = "Debe seleccionar una localidad de Procedencia"
                                else
                                    msj = "Debe seleccionar una provincia de Procedencia"
                                end if
                            'else
                            '    msj = "Debe ingresar el Cupo"
                            'end if
                        else
                            msj = "Debe ingresar el Peso Neto"
                        end if
                    else
                        msj = "Debe ingresar el Peso Tara"
                    end if
                else
                    msj = "Debe ingresar el Peso Bruto"
                end if
            else
                msj = "Debe seleccionar un Producto"
            end if
        else
            msj = "La cosecha presenta un error en el formato" 
        end if
    else
        msj = "Debe ingresar la cosecha" 
    end if
    controlarProductoCartaPorte = msj
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function controlarTransporteCartaPorte()
    Dim msj
    msj = ""
    if (Trim(auxChapa) <> "") then
        if (Trim(auxAcoplado) = "") then msj = "Debe ingresar el acoplado del camion"
    else
        msj = "Debe ingresar la chapa del camion"
    end if
    controlarTransporteCartaPorte = msj
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function controlarDescargaCartaPorte()
    Dim msj
    msj = ""
    if (Cdbl(auxTurno) > 0) then
        if (controlarMaximoPesoCamion(PARAM_MAX_TARA, auxPesadaTara)) then
            if (controlarMinimoPesoCamion(PARAM_MIN_TARA, auxPesadaTara)) then
                if (controlarMaximoPesoCamion(PARAM_MAX_BRUTO, auxPesadaBruto)) then
                    if (not controlarMinimoPesoCamion(PARAM_MIN_BRUTO, auxPesadaBruto)) then
                        msj = "Los Kgs Bruto no deben ser inferiores al Minimo Tara."
                    end if
                else
                    msj = "Los Kgs Bruto no deben ser superiores al Maximo Tara."
                end if
            else
                msj = "Los Kgs Tara no deben ser inferiores al Minimo Tara."
            end if
        else
            msj = "Los Kgs Tara no deben ser superiores al Maximo Tara."
        end if
    else     
        msj = "Debe ingresar el turno "
    end if
    controlarDescargaCartaPorte = msj
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function controlarMaximoPesoCamion(p_Parametro, p_Kilo)
    Dim vlParametro 
    controlarMaximoPesoCamion = true
    vlParametro = getValueParametro(p_Parametro, g_strPuerto)
    if (Cdbl(p_Kilo) > Cdbl(vlParametro)) then controlarMaximoPesoCamion = false
End function
'-----------------------------------------------------------------------------------------------------------------------------------
Function controlarMinimoPesoCamion(p_Parametro, p_Kilo)
    Dim vlParametro 
    controlarMinimoPesoCamion = true
    vlParametro = getValueParametro(p_Parametro, g_strPuerto)
    if (Cdbl(p_Kilo) < Cdbl(vlParametro)) then controlarMinimoPesoCamion = false
End function
'-----------------------------------------------------------------------------------------------------------------------------------
'En caso de modificar la carta de porte controla que no exista otra con igual codigo 
Function controlarDuplicidadCartaPorte(p_CtaPteNew,p_CtaPteOld)
    Dim strSQL,rs,myEstadoBaja,myEstadoEgreRech,myEstadoRechazo
    controlarDuplicidadCartaPorte = true
    if (Trim(p_CtaPteNew) <> Trim(p_CtaPteOld)) then
        'Obtengo los estados que no se tendran en cuenta a la hora de validar la carta de porte
        myEstadoBaja     = getValueParametro(PARAM_ESTADO_BAJA,g_strPuerto)
        myEstadoEgreRech = getValueParametro(PARAM_ESTADO_EGRERECH,g_strPuerto)
        myEstadoRechazo  = getValueParametro(PARAM_ESTADO_RECHAZADO,g_strPuerto)
        'Verifico tanto en la descarga como en la carga por la carta de porte
        strSQL = "SELECT A.* "&_
                 "FROM ( SELECT NUCARTAPORTE,IDCAMION,DTCONTABLE FROM HCAMIONESCARGA WHERE NUCARTAPORTE = '"& p_CtaPteNew &"'"&_
	             "       UNION "&_
	             "       SELECT NUCARTAPORTE,IDCAMION,DTCONTABLE FROM HCAMIONESDESCARGA WHERE NUCARTAPORTE = '"& p_CtaPteNew &"') A "&_
                 "INNER JOIN HCAMIONES B ON A.DTCONTABLE = B.DTCONTABLE AND A.IDCAMION = B.IDCAMION "&_
                 "WHERE B.CDESTADO NOT IN ("& myEstadoBaja &","& myEstadoEgreRech &","& myEstadoRechazo &") "
        Call GF_BD_Puertos(g_strPuerto, rs, "OPEN",strSQL)
        if (not rs.Eof) then controlarDuplicidadCartaPorte = false
    end if
End function
'-----------------------------------------------------------------------------------------------------------------------------------
Function loadDataCartaPorte(p_CtaPte, p_idCamion, p_DtContable, p_Pto)
    Dim strSQL, auxDtContable
    auxDtContable = Left(p_DtContable,4) &"-"& Mid(p_DtContable,5,2) &"-"& Right(p_DtContable,2)
    strSQL =" SELECT ((YEAR(T.DTCONTABLE)*10000) + (MONTH(T.DTCONTABLE)*100) + DAY(T.DTCONTABLE)) as DTCONTABLE,"&_
            "        T.IDCAMION,T.NUCARTAPORTE,T.CTG,"&_
            "        ((YEAR(T.DTCARTAPORTE)*10000) + (MONTH(T.DTCARTAPORTE)*100) + DAY(T.DTCARTAPORTE)) as DTCARTAPORTE, "&_
            "        H.DSCLIENTE,I.DSVENDEDOR,T.CDTRANSPORTISTA, "&_
		    "        T.CDCOSECHA,T.CDPROCEDENCIA,T.CDCORREDOR,T.CDENTREGADOR,T.CDCLIENTE,T.CDVENDEDOR,"&_
            "        ((YEAR(T.DTCPVENCIMIENTO)*10000) + (MONTH(T.DTCPVENCIMIENTO)*100) + DAY(T.DTCPVENCIMIENTO)) as DTCPVENCIMIENTO, "&_
            "        T.CDTIPODOC,T.NUDOCUMENTO,T.DSAPELLIDOCONDUCTOR,T.DSNOMBRECONDUCTOR,T.CDPRODUCTO,T.NUCUPO,G.DSENTREGADOR,"&_
            "        CASE WHEN K.CDPROV IS NULL THEN 0 ELSE K.CDPROV END AS CDPROV,"&_
            "        CASE WHEN K.DSPROCEDENCIA IS NULL THEN '' ELSE RTRIM(K.DSPROCEDENCIA) END AS DSPROCEDENCIA, "&_
	        "        T.CDCHAPACAMION,T.CDCHAPAACOPLADO,"&_
            "        CAST(INGRESOFECHA as BIGINT)*1000000 + INGRESOHORA AS DTINGRESO, "&_
            "        CAST(EGRESOFECHA as BIGINT)*1000000 + EGRESOHORA AS DTEGRESO, "&_
            "        T.SQTURNO,T.NUCUITREM,D.DSBIOTECNOLOGIA,F.DSCORREDOR,J.DSTRANSPORTISTA, "&_
            "        CASE WHEN T.VLBRUTOORIGEN IS NULL THEN 0 ELSE T.VLBRUTOORIGEN END AS VLBRUTOORIGEN, "&_
            "        CASE WHEN T.VLTARAORIGEN IS NULL THEN 0 ELSE T.VLTARAORIGEN END AS VLTARAORIGEN, "&_
            "        CASE WHEN F.NUCUIT IS NULL THEN '' ELSE RTRIM(F.NUCUIT) END AS CUITCORREDOR, "&_
            "        CASE WHEN G.NUCUIT IS NULL THEN '' ELSE RTRIM(G.NUCUIT) END AS CUITENTREGADOR, "&_
            "        CASE WHEN H.NUCUIT IS NULL THEN '' ElSE RTRIM(H.NUCUIT) END AS CUITCLIENTE, "&_
            "        CASE WHEN I.NUDOCUMENTO IS NULL THEN '' ELSE RTRIM(I.NUDOCUMENTO) END AS CUITVENDEDOR, "&_
            "        CASE WHEN J.CDTRANSPORTISTA IS NULL THEN '' ELSE RTRIM(J.NUDOCUMENTO) END AS CUITTRANSPORTISTA, "&_
            "        CASE WHEN C.IDBIOTECNOLOGIA IS NULL THEN 0 ELSE C.IDBIOTECNOLOGIA END AS IDBIOTECNOLOGIA, "&_
            "        CASE WHEN E.DSOBSERVACIONES IS NULL THEN '' ELSE RTRIM(E.DSOBSERVACIONES) END AS DSOBSERVACIONES, "&_
            "        CASE WHEN T.OBSERVACIONESCALADA IS NULL THEN '' ELSE RTRIM(T.OBSERVACIONESCALADA) END AS OBSERVACIONESCALADA, "&_
            "        CASE WHEN T.BRUTO IS NULL THEN 0 ELSE T.BRUTO END AS BRUTO, "&_
            "        CASE WHEN T.TARA IS NULL THEN 0 ELSE T.TARA END AS TARA, "&_
            "        CASE WHEN T.MERMA IS NULL THEN 0 ELSE T.MERMA END AS MERMA, "&_
            "        CASE WHEN T.MERMAPORCENTAJE IS NULL THEN 0 ELSE T.MERMAPORCENTAJE END AS MERMAPORCENTAJE, "&_
            "        CASE WHEN L.NCUIT IS NULL THEN '' ELSE L.NCUIT END AS CUIT_CHOFER, "&_
            "        CASE WHEN L.DSNOMBRE IS NULL THEN '' ELSE L.DSNOMBRE END AS NOM_CHOFER, "&_
            "        CASE WHEN L.DSAPELLIDO IS NULL THEN '' ELSE L.DSAPELLIDO END AS AP_CHOFER "&_
            " FROM ( SELECT A.*, "&_
	        "               B.CDTIPODOC,B.NUDOCUMENTO,B.DSAPELLIDOCONDUCTOR,B.DSNOMBRECONDUCTOR,B.CDPRODUCTO,B.NUCUPO, "&_
	        "               B.CDCHAPACAMION,B.CDCHAPAACOPLADO,"&_
            "               ((Year(b.dtingreso) * 10000) + (Month(b.dtingreso) * 100) + Day(b.dtingreso))AS INGRESOFECHA, "&_
            "               ((DATEPART(HOUR, b.dtingreso) * 10000) + (DATEPART(MINUTE, b.dtingreso) * 100) + DATEPART(SECOND, b.dtingreso)) AS INGRESOHORA, "&_
            "               ((Year(b.DTEGRESO) * 10000) + (Month(b.DTEGRESO) * 100) + Day(b.DTEGRESO))AS EGRESOFECHA, "&_
            "               ((DATEPART(HOUR, b.dtegreso) * 10000) + (DATEPART(MINUTE, b.dtegreso) * 100) + DATEPART(SECOND, b.dtegreso)) AS EGRESOHORA, "&_
            "               B.DTEGRESO,B.SQTURNO,B.NUCUITREM,B.CDTRANSPORTISTA, "&_
	        "               (SELECT CASE WHEN PC.VLPESADA IS NULL THEN 0 ELSE PC.VLPESADA END AS VLPESADA "&_
	        "               FROM HPESADASCAMION PC "&_
	        "               WHERE PC.DTCONTABLE = A.DTCONTABLE AND PC.IDCAMION = A.IDCAMION AND PC.CDPESADA = 1 "&_
		    "                   AND PC.SQPESADA = (SELECT MAX(SQPESADA) "&_
			"		                                FROM HPESADASCAMION "&_
			"		                                WHERE PC.DTCONTABLE = DTCONTABLE AND PC.IDCAMION = IDCAMION AND CDPESADA = 1)) AS BRUTO, "&_
	        "               (SELECT CASE WHEN PC.VLPESADA IS NULL THEN 0 ELSE PC.VLPESADA END AS VLPESADA "&_
	        "               FROM HPESADASCAMION PC "&_
	        "               WHERE PC.DTCONTABLE = A.DTCONTABLE AND PC.IDCAMION = A.IDCAMION AND PC.CDPESADA = 2 "&_
	   	    "                   AND PC.SQPESADA = (SELECT MAX(SQPESADA) "&_
	   		"		                                FROM HPESADASCAMION "&_
	   		"		                                WHERE PC.DTCONTABLE = DTCONTABLE AND PC.IDCAMION = IDCAMION AND CDPESADA = 2)) AS TARA, "&_
            "		        (SELECT CASE WHEN HMC.VLMERMAKILOS IS NULL THEN 0 ELSE HMC.VLMERMAKILOS END "&_
            "                FROM HMERMASCAMIONES HMC "&_
            "                WHERE HMC.DTCONTABLE=A.DTCONTABLE AND HMC.IDCAMION = A.IDCAMION "&_
            "                   AND HMC.SQPESADA= (SELECT MAX(SQPESADA) "&_
            "                                      FROM HMERMASCAMIONES "&_
            "                                      WHERE DTCONTABLE=HMC.DTCONTABLE AND IDCAMION = HMC.IDCAMION)) AS MERMA, "&_
	        "               (SELECT CASE WHEN CD.DSOBSERVACIONES IS NULL THEN '' ELSE CD.DSOBSERVACIONES END AS DSOBSERVACIONES "&_
	        "               FROM HCALADADECAMIONES CD  "&_
	        "               WHERE CD.DTCONTABLE = A.DTCONTABLE AND CD.IDCAMION = A.IDCAMION  "&_
	   	    "                   AND CD.SQCALADA = (SELECT MAX(SQCALADA) "&_
	   		"		                                FROM HCALADADECAMIONES "&_
	   		"		                                WHERE CD.DTCONTABLE = DTCONTABLE AND CD.IDCAMION = IDCAMION)) AS OBSERVACIONESCALADA, "&_
            "               (SELECT CASE WHEN CD.PCMERMA IS NULL THEN 0 ELSE CD.PCMERMA END AS PCMERMA "&_
            "                FROM HCALADADECAMIONES CD "&_
            "                WHERE CD.DTCONTABLE = A.DTCONTABLE AND CD.IDCAMION = A.IDCAMION "&_
            "                   AND CD.SQCALADA = (SELECT MAX(SQCALADA) "&_
            "					                   FROM HCALADADECAMIONES  "&_
			"                               	   WHERE CD.DTCONTABLE = DTCONTABLE AND CD.IDCAMION = IDCAMION)) AS MERMAPORCENTAJE "&_
            "       FROM ( SELECT DTCONTABLE,IDCAMION,NUCARTAPORTE,CTG,DTCARTAPORTE,DTCPVENCIMIENTO,CDCORREDOR,CDENTREGADOR, "&_
		    "                     CDCLIENTE,CDCOSECHA,VLBRUTOORIGEN,VLTARAORIGEN,CDPROCEDENCIA,CDVENDEDOR "&_
	        "               FROM HCAMIONESDESCARGA  "&_
	        "               WHERE IDCAMION='" & p_idCamion & "' and NUCARTAPORTE = '"& p_CtaPte &"' AND DTCONTABLE = '"& auxDtContable &"') A "&_
            "       INNER JOIN HCAMIONES B ON A.DTCONTABLE = B.DTCONTABLE AND A.IDCAMION = B.IDCAMION ) T "&_
            "  LEFT JOIN TBLBIOTECNOLOGIASDECLARADAS C ON C.NUCARTAPORTE = T.NUCARTAPORTE and C.TIPOTRANSPORTE = 1 "&_
            "  LEFT JOIN TBLBIOTECNOLOGIAS D ON D.IDBIOTECNOLOGIA = C.IDBIOTECNOLOGIA "&_
            "  LEFT JOIN OBSERVACIONESCAMION E ON E.IDCAMION = T.IDCAMION "&_
            "  LEFT JOIN CORREDORES F ON F.CDCORREDOR = T.CDCORREDOR "&_
            "  LEFT JOIN ENTREGADORES G ON G.CDENTREGADOR = T.CDENTREGADOR "&_
            "  LEFT JOIN CLIENTES H ON H.CDCLIENTE = T.CDCLIENTE "&_
            "  LEFT JOIN VENDEDORES I ON I.CDVENDEDOR = T.CDVENDEDOR "&_
            "  LEFT JOIN TRANSPORTISTAS J ON J.CDTRANSPORTISTA = T.CDTRANSPORTISTA "&_
            "  LEFT JOIN PROCEDENCIAS K ON K.CDPROCEDENCIA = T.CDPROCEDENCIA " &_
            "  LEFT JOIN CONDUCTOR L ON SUBSTRING( L.NCUIT ,3 , 8 ) = T.NUDOCUMENTO"            
            Call GF_BD_Puertos(p_Pto, rs, "OPEN",strSQL)
    Set loadDataCartaPorte = rs
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function getProvinciaProcedencia(p_Pto)
    Dim strSQL, rsProc
    strSQL = "SELECT CDPROVINCIA,RTRIM(DSPROVINCIA) AS DSPROVINCIA FROM PROVINCIAS ORDER BY CDPROVINCIA"
    Call GF_BD_Puertos(p_Pto, rsProc, "OPEN",strSQL)
    Set getProvinciaProcedencia = rsProc
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function getProductosPto(p_Pto)
    Dim strSQL, rsProd
    strSQL = "SELECT CDPRODUCTO,RTRIM(DSPRODUCTO) AS DSPRODUCTO FROM PRODUCTOS ORDER BY CDPRODUCTO"
    Call GF_BD_Puertos(p_Pto, rsProd, "OPEN",strSQL)
    Set getProductosPto = rsProd
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function getBiotecnologiaByProducto(p_Pto, p_CdProducto)
    Dim strSQL, rsBio
    strSQL = "SELECT IDBIOTECNOLOGIA,RTRIM(DSBIOTECNOLOGIA) AS DSBIOTECNOLOGIA FROM TBLBIOTECNOLOGIAS WHERE IDPRODUCTO ="&p_CdProducto&" ORDER BY IDBIOTECNOLOGIA"
    Call GF_BD_Puertos(p_Pto, rsBio, "OPEN",strSQL)
    Set getBiotecnologiaByProducto = rsBio
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function getDsVendedorPto(pCuit, pCdVendedor)
    Dim strSQL,myWhere,rsVen
    getDsVendedorPto = ""
    if (pCuit <> "") then Call mkWhere(myWhere, "NUDOCUMENTO", pCuit, "=", 3)
    if (pCdVendedor <> "") then Call mkWhere(myWhere, "CDVENDEDOR", pCdVendedor, "=", 1)
    strSQL = "SELECT CASE WHEN DSVENDEDOR IS NULL THEN '' ELSE RTRIM(DSVENDEDOR) END AS DSVENDEDOR  FROM VENDEDORES "& myWhere
    Call GF_BD_Puertos(g_pto, rsVen, "OPEN",strSQL)
    if (not rsVen.Eof) then getDsVendedorPto = Trim(rsVen("DSVENDEDOR"))    
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function getCuitVendedorPto(pCdVendedor)
    Dim strSQL,myWhere,rsVen
    getCuitVendedorPto = ""
    if (pCdVendedor <> "") then
        strSQL = "SELECT CASE WHEN NUDOCUMENTO IS NULL THEN '' ELSE RTRIM(NUDOCUMENTO) END AS NUDOCUMENTO FROM VENDEDORES WHERE CDVENDEDOR = "& pCdVendedor
        Call GF_BD_Puertos(g_pto, rsVen, "OPEN",strSQL)
        if (not rsVen.Eof) then getCuitVendedorPto = Trim(rsVen("NUDOCUMENTO"))
    end if
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function getCdVendedorPto(pCuit)
    Dim strSQL,myWhere,rsVen
    getCdVendedorPto = -1
    if (pCuit <> "") then
        strSQL = "SELECT CASE WHEN CDVENDEDOR IS NULL THEN 0 ELSE CDVENDEDOR END AS CDVENDEDOR FROM VENDEDORES WHERE RTRIM(NUDOCUMENTO) = '"& Trim(pCuit) &"'"
        Call GF_BD_Puertos(g_pto, rsVen, "OPEN",strSQL)
        if (not rsVen.Eof) then getCdVendedorPto = rsVen("CDVENDEDOR")
    end if
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
'Carga los datos de la cabecera a variables, dependiendo de la accion lo hace por parametros o por la DB
Function fetchCabeceraCartaPorte()
    if (not isFormSubmit) then
        auxCartaPorte1 = Left(g_ctaPte,4)
        auxCartaPorte2 = Right(g_ctaPte,8)
        g_ctaPteOld = auxCartaPorte1 & auxCartaPorte2
        auxCTG = g_rs("CTG")
        auxCTGOld = auxCTG
        auxDtCartaPorte = g_rs("DTCARTAPORTE")
        auxDtCartaPorteOld = auxDtCartaPorte
        auxDtVencimiento = g_rs("DTCPVENCIMIENTO")
        auxDtVencimientoOld = auxDtVencimiento
        g_IdCamion = Cstr(g_rs("IDCAMION"))
    else
        auxCartaPorte1 = GF_PARAMETROS7("txtCartaPorte1", "" ,6)
        auxCartaPorte2 = GF_PARAMETROS7("txtCartaPorte2", "" ,6)
        g_ctaPteOld = GF_PARAMETROS7("cartaPorteOld", "" ,6)
        auxCTG = GF_PARAMETROS7("txtCTG", "" ,6)
        auxCTGOld = GF_PARAMETROS7("txtCTGOld", "" ,6)
        auxDtCartaPorte = GF_PARAMETROS7("issuedateCarga", "" ,6)
        auxDtCartaPorteOld = GF_PARAMETROS7("issuedateCargaOld", "" ,6)
        auxDtVencimiento = GF_PARAMETROS7("issuedateVencimiento", "" ,6)
        auxDtVencimientoOld = GF_PARAMETROS7("issuedateVencimientoOld", "" ,6)
        g_IdCamion = GF_PARAMETROS7("idCamion", "" ,6)
    end if
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function fetchIntervinientesCartaPorte()
    Dim auxCuit1,auxCuit2,auxCuit3
    if (not isFormSubmit) then
        auxCorredor = g_rs("CDCORREDOR")
        auxCorredorOld = auxCorredor
        auxCorredorCuit1 = Left(g_rs("CUITCORREDOR"),2)
        auxCorredorCuit2 = Mid(g_rs("CUITCORREDOR"),3,8)
        auxCorredorCuit3 = Right(g_rs("CUITCORREDOR"),1)
        auxCorredorCuitOld = auxCorredorCuit1 & auxCorredorCuit2 & auxCorredorCuit3
        auxCorredorDs = g_rs("DSCORREDOR")
        auxEntregador = g_rs("CDENTREGADOR")
        auxEntregadorOld = auxEntregador
        auxEntregadorCuit1 = Left(g_rs("CUITENTREGADOR"),2)
        auxEntregadorCuit2 = Mid(g_rs("CUITENTREGADOR"),3,8)
        auxEntregadorCuit3 = Right(g_rs("CUITENTREGADOR"),1)
        auxEntregadorCuitOld = auxEntregadorCuit1 & auxEntregadorCuit2 & auxEntregadorCuit3
        auxEntregadorDs = g_rs("DSENTREGADOR")
        auxDestinatario = g_rs("CDCLIENTE")
        auxDestinatarioOld = auxDestinatario
        auxDestinatarioCuit1 = Left(g_rs("CUITCLIENTE"),2)
        auxDestinatarioCuit2 = Mid(g_rs("CUITCLIENTE"),3,8)
        auxDestinatarioCuit3 = Right(g_rs("CUITCLIENTE"),1)
        auxDestinatarioCuitOld = auxDestinatarioCuit1 & auxDestinatarioCuit2 & auxDestinatarioCuit3
        auxDestinatarioDs = g_rs("DSCLIENTE")
        auxTransportista = g_rs("CDTRANSPORTISTA")
        auxTransportistaOld = auxTransportista 
        auxTransportistaDs = g_rs("DSTRANSPORTISTA")
        auxTransportistaCuit1 = Left(g_rs("CUITTRANSPORTISTA"),2)
        auxTransportistaCuit2 = Mid(g_rs("CUITTRANSPORTISTA"),3,8)
        auxTransportistaCuit3 = Right(g_rs("CUITTRANSPORTISTA"),1)
        auxTransportistaCuitOld = auxTransportistaCuit1 & auxTransportistaCuit2 & auxTransportistaCuit3
        if (Trim(g_rs("CUIT_CHOFER")) <> "" ) then            
            auxChoferNumDoc1 = Left(g_rs("CUIT_CHOFER"),2) 
            auxChoferNumDoc2 = Mid(g_rs("CUIT_CHOFER"),3,8)
            auxChoferNumDoc3 = Right(g_rs("CUIT_CHOFER"),1)
            auxChoferDs = g_rs("AP_CHOFER") &", "& g_rs("NOM_CHOFER")
        ELSE
            auxChoferNumDoc1 = ""
            auxChoferNumDoc2 = g_rs("NUDOCUMENTO") 'Se trata de un DNI
            auxChoferNumDoc3 = ""
            auxChoferDs = g_rs("DSAPELLIDOCONDUCTOR") &", "& g_rs("DSNOMBRECONDUCTOR")
        end if
        
        auxChoferCuitOld = auxChoferNumDoc1 & auxChoferNumDoc2 & auxChoferNumDoc3
        auxChoferTipoDoc = auxChoferCuitOld
        auxChoferTipoDocOld = auxChoferTipoDoc
        
        Call cargarIntermediariosCtaPte(Cstr(g_rs("NUCUITREM")),Cstr(g_rs("IDCAMION")),Cstr(g_rs("CDVENDEDOR")))
    else
        auxTitularDs = GF_PARAMETROS7("valDs_" & INTERVINIENTE_TITULAR, "" ,6)
        auxTitularCd = GF_PARAMETROS7("valCd_" & INTERVINIENTE_TITULAR, 0 ,6)
        auxTitularCuit1 = GF_PARAMETROS7("Cuit_" & INTERVINIENTE_TITULAR &"_1", "" ,6)
        auxTitularCuit2 = GF_PARAMETROS7("Cuit_" & INTERVINIENTE_TITULAR &"_2", "" ,6)
        auxTitularCuit3 = GF_PARAMETROS7("Cuit_" & INTERVINIENTE_TITULAR &"_3", "" ,6)
        auxTitularCdOld = GF_PARAMETROS7("valCdOld_" & INTERVINIENTE_TITULAR, 0 ,6)
        auxTitularCuitOld = GF_PARAMETROS7("valCuitOld_" & INTERVINIENTE_TITULAR, "" ,6)
        auxRemitenteDs = GF_PARAMETROS7("valDs_" & INTERVINIENTE_REMITENTE, "" ,6)
        auxRemitenteCd = GF_PARAMETROS7("valCd_" & INTERVINIENTE_REMITENTE, 0 ,6)
        auxRemitenteCdOld = GF_PARAMETROS7("valCdOld_" & INTERVINIENTE_REMITENTE, 0 ,6)
        auxRemitenteCuitOld = GF_PARAMETROS7("valCuitOld_" & INTERVINIENTE_REMITENTE, "" ,6)
        auxRemitenteCuit1 = GF_PARAMETROS7("Cuit_" & INTERVINIENTE_REMITENTE &"_1", "" ,6)
        auxRemitenteCuit2 = GF_PARAMETROS7("Cuit_" & INTERVINIENTE_REMITENTE &"_2", "" ,6)
        auxRemitenteCuit3 = GF_PARAMETROS7("Cuit_" & INTERVINIENTE_REMITENTE &"_3", "" ,6)
        auxIntermediarioDs = GF_PARAMETROS7("valDs_" & INTERVINIENTE_INTERMEDIARIO, "" ,6)
        auxIntermediarioCd = GF_PARAMETROS7("valCd_" & INTERVINIENTE_INTERMEDIARIO, 0 ,6)
        auxIntermediarioCuitOld = GF_PARAMETROS7("valCuitOld_" & INTERVINIENTE_INTERMEDIARIO, "" ,6)
        auxIntermediarioCdOld = GF_PARAMETROS7("valCdOld_" & INTERVINIENTE_INTERMEDIARIO, 0 ,6)
        auxIntermediarioCuit1 = GF_PARAMETROS7("Cuit_" & INTERVINIENTE_INTERMEDIARIO &"_1", "" ,6)
        auxIntermediarioCuit2 = GF_PARAMETROS7("Cuit_" & INTERVINIENTE_INTERMEDIARIO &"_2", "" ,6)
        auxIntermediarioCuit3 = GF_PARAMETROS7("Cuit_" & INTERVINIENTE_INTERMEDIARIO &"_3", "" ,6)
        auxCorredor = GF_PARAMETROS7("valCd_" & INTERVINIENTE_CORREDOR, 0 ,6)
        auxCorredorOld = GF_PARAMETROS7("valCdOld_" & INTERVINIENTE_CORREDOR, 0 ,6)
        auxCorredorCuit1 = GF_PARAMETROS7("Cuit_" & INTERVINIENTE_CORREDOR &"_1", "" ,6)
        auxCorredorCuit2 = GF_PARAMETROS7("Cuit_" & INTERVINIENTE_CORREDOR &"_2", "" ,6)
        auxCorredorCuit3 = GF_PARAMETROS7("Cuit_" & INTERVINIENTE_CORREDOR &"_3", "" ,6)
        auxCorredorDs = GF_PARAMETROS7("valDs_" & INTERVINIENTE_CORREDOR, "" ,6)
        auxCorredorCuitOld = GF_PARAMETROS7("valCuitOld_" & INTERVINIENTE_CORREDOR, "" ,6)
        auxEntregador = GF_PARAMETROS7("valCd_" & INTERVINIENTE_ENTREGADOR, 0 ,6)
        auxEntregadorOld = GF_PARAMETROS7("valCdOld_" & INTERVINIENTE_ENTREGADOR, 0 ,6)
        auxEntregadorCuit1 = GF_PARAMETROS7("Cuit_" & INTERVINIENTE_ENTREGADOR &"_1", "" ,6)
        auxEntregadorCuit2 = GF_PARAMETROS7("Cuit_" & INTERVINIENTE_ENTREGADOR &"_2", "" ,6)
        auxEntregadorCuit3 = GF_PARAMETROS7("Cuit_" & INTERVINIENTE_ENTREGADOR &"_3", "" ,6)
        auxEntregadorDs = GF_PARAMETROS7("valDs_" & INTERVINIENTE_ENTREGADOR, "" ,6)
        auxEntregadorCuitOld = GF_PARAMETROS7("valCuitOld_" & INTERVINIENTE_ENTREGADOR, "" ,6)
        auxDestinatario = GF_PARAMETROS7("valCd_" & INTERVINIENTE_DESTINATARIO, 0 ,6)
        auxDestinatarioOld = GF_PARAMETROS7("valCdOld_" & INTERVINIENTE_DESTINATARIO, 0 ,6) 
        auxDestinatarioCuit1 = GF_PARAMETROS7("Cuit_" & INTERVINIENTE_DESTINATARIO &"_1", "" ,6)
        auxDestinatarioCuit2 = GF_PARAMETROS7("Cuit_" & INTERVINIENTE_DESTINATARIO &"_2", "" ,6)
        auxDestinatarioCuit3 = GF_PARAMETROS7("Cuit_" & INTERVINIENTE_DESTINATARIO &"_3", "" ,6)
        auxDestinatarioDs = GF_PARAMETROS7("valDs_" & INTERVINIENTE_DESTINATARIO, "" ,6)
        auxDestinatarioCuitOld = GF_PARAMETROS7("valCuitOld_" & INTERVINIENTE_DESTINATARIO, "" ,6)
        auxTransportista = GF_PARAMETROS7("valCd_" & INTERVINIENTE_TRANSPORTISTA, 0 ,6)
        auxTransportistaOld = GF_PARAMETROS7("valCdOld_" & INTERVINIENTE_TRANSPORTISTA, 0 ,6)
        auxTransportistaCuit1 = GF_PARAMETROS7("Cuit_" & INTERVINIENTE_TRANSPORTISTA &"_1", "" ,6)
        auxTransportistaCuit2 = GF_PARAMETROS7("Cuit_" & INTERVINIENTE_TRANSPORTISTA &"_2", "" ,6)
        auxTransportistaCuit3 = GF_PARAMETROS7("Cuit_" & INTERVINIENTE_TRANSPORTISTA &"_3", "" ,6)
        auxTransportistaDs = GF_PARAMETROS7("valDs_" & INTERVINIENTE_TRANSPORTISTA, "" ,6)
        auxTransportistaCuitOld = GF_PARAMETROS7("valCuitOld_" & INTERVINIENTE_TRANSPORTISTA, "" ,6)
        auxChoferTipoDoc = GF_PARAMETROS7("valCd_" & INTERVINIENTE_CHOFER, "" ,6)
        auxChoferCuitOld = GF_PARAMETROS7("valCuitOld_" & INTERVINIENTE_CHOFER, "" ,6)
        auxChoferNumDoc1 = GF_PARAMETROS7("Cuit_" & INTERVINIENTE_CHOFER & "_1", "" ,6)
        auxChoferNumDoc2 = GF_PARAMETROS7("Cuit_" & INTERVINIENTE_CHOFER & "_2", "" ,6)
        auxChoferNumDoc3 = GF_PARAMETROS7("Cuit_" & INTERVINIENTE_CHOFER & "_3", "" ,6)
        auxChoferDs = GF_PARAMETROS7("valDs_" & INTERVINIENTE_CHOFER, "" ,6)
        auxChoferTipoDocOld = GF_PARAMETROS7("valCdOld_" & INTERVINIENTE_CHOFER, "" ,6)
        auxChoferNewNombre = GF_PARAMETROS7("Nom_Chofer_New","",6) 
        auxchoferNewApellido = GF_PARAMETROS7("Ape_Chofer_New","",6)
    end if
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function fecthProductoCartaPorte()
    if (not isFormSubmit) then
        auxCosecha = g_rs("CDCOSECHA")
        auxCosechaOld = auxCosecha
        auxGrano = g_rs("CDPRODUCTO")
        auxGranoOld = auxGrano
        auxPesoBruto =CDbl(g_rs("VLBRUTOORIGEN"))
        auxPesoBrutoOld = auxPesoBruto
        auxPesoTara = CDbl(g_rs("VLTARAORIGEN"))
        auxPesoTaraOld = auxPesoTara
        auxPesoNeto = Cdbl(g_rs("VLBRUTOORIGEN")) - Cdbl(g_rs("VLTARAORIGEN"))
        auxPesoNetoOld = auxPesoNeto
        auxCupo = Trim(g_rs("NUCUPO"))
        auxCupoOld = auxCupo
        auxBiotecnologia = g_rs("IDBIOTECNOLOGIA")
        auxBiotecnologiaOld = auxBiotecnologia
        auxObservaciones = editText4Input(Trim(g_rs("DSOBSERVACIONES")))
        auxObservacionesOld = auxObservaciones
        auxProcedenciaProv = g_rs("CDPROV")        
        auxProcedenciaCd = g_rs("CDPROCEDENCIA")
        auxProcedenciaCdOld = auxProcedenciaCd
        auxProcedenciaDs = g_rs("DSPROCEDENCIA")
    else
        auxCosecha = GF_PARAMETROS7("cosecha", "" ,6)
        auxCosechaOld = GF_PARAMETROS7("cosechaOld", "" ,6)
        auxGrano = GF_PARAMETROS7("cmbProducto", 0 ,6)
        auxGranoOld = GF_PARAMETROS7("cdProductoOld", 0 ,6)
        auxPesoBruto = GF_PARAMETROS7("pesoBruto", "" ,6)
        auxPesoBrutoOld = GF_PARAMETROS7("pesoBrutoOld", "" ,6)
        auxPesoTara = GF_PARAMETROS7("pesoTara", "" ,6)
        auxPesoTaraOld = GF_PARAMETROS7("pesoTaraOld", "" ,6)
        auxPesoNeto = GF_PARAMETROS7("pesoNeto", "" ,6)
        auxPesoNetoOld = GF_PARAMETROS7("pesoNetoOld", "",6)
        auxCupo = GF_PARAMETROS7("cupo", "" ,6)
        auxCupoOld = GF_PARAMETROS7("cupoOld", "" ,6)
        auxBiotecnologia = GF_PARAMETROS7("cmbBiotecnologia", 0 ,6)
        auxBiotecnologiaOld = GF_PARAMETROS7("idBiotecnologiaOld", 0,6)
        auxObservaciones = GF_PARAMETROS7("observaciones", "" ,6)
        auxObservacionesOld = GF_PARAMETROS7("observacionesOld", "",6)
        auxProcedenciaProv = GF_PARAMETROS7("cmbProvincia", 0 ,6)
        auxProcedenciaProvOld = GF_PARAMETROS7("cdProvinciaOld", 0 ,6)
        auxProcedenciaCd = GF_PARAMETROS7("procedenciaCd", 0 ,6)
        auxProcedenciaCdOld = GF_PARAMETROS7("procedenciaCdOld", 0, 6)
        auxProcedenciaDs = GF_PARAMETROS7("procedenciaDs", "" ,6)
    end if
end Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function fecthTransporteCartaPorte()
    if (not isFormSubmit) then
        auxChapa = Trim(Ucase(g_rs("CDCHAPACAMION")))
        auxChapaOld = auxChapa
        auxAcoplado = Trim(Ucase(g_rs("CDCHAPAACOPLADO")))
        auxAcopladoOld = auxAcoplado
    else
        auxChapa = GF_PARAMETROS7("chapa", "" ,6)
        auxChapaOld = GF_PARAMETROS7("chapaOld", "", 6)
        auxAcoplado = GF_PARAMETROS7("acoplado", "" ,6)
        auxAcopladoOld = GF_PARAMETROS7("acopladoOld", "", 6)
    end if
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function fecthDescargaCartaPorte()
    if (not isFormSubmit) then
        auxFechaArribo = Left(GF_FN2DTE(g_rs("DTINGRESO")),10)
        auxHoraArribo = Right(GF_FN2DTE(g_rs("DTINGRESO")),8)
        auxFechaEgreso = Left(GF_FN2DTE(g_rs("DTEGRESO")),10)
        auxHoraEgreso = Right(GF_FN2DTE(g_rs("DTEGRESO")),8)
        auxTurno = Cdbl(g_rs("SQTURNO"))
        auxTurnoOld = auxTurno 
        auxPesadaBruto = Cdbl(g_rs("BRUTO"))
        auxPesadaBrutoOld = auxPesadaBruto
        auxPesadaTara = Cdbl(g_rs("TARA"))
        auxPesadaTaraOld = auxPesadaTara
        auxMerma = Cdbl(g_rs("MERMA"))
        auxMermaOld = auxMerma
        auxMermaPorcentaje = g_rs("MERMAPORCENTAJE")
        auxMermaPorcentajeOld = auxMermaPorcentaje
        auxNetoCMerma = Cdbl(auxPesadaBruto) - Cdbl(auxPesadaTara) - Cdbl(auxMerma)
        auxNetoSMerma = Cdbl(auxPesadaBruto) - Cdbl(auxPesadaTara)
        auxObservacionesDescarga = editText4Input(Trim(g_rs("OBSERVACIONESCALADA")))
    else
        auxFechaArribo = GF_PARAMETROS7("fechaArribo", "" ,6)
        auxHoraArribo = GF_PARAMETROS7("horaArribo", "" ,6)
        auxFechaEgreso = GF_PARAMETROS7("fechaEgreso", "" ,6)
        auxHoraEgreso = GF_PARAMETROS7("horaDescarga", "" ,6)
        auxTurno = GF_PARAMETROS7("turno", 0 ,6)
        auxTurnoOld = GF_PARAMETROS7("turnoOld", 0, 6)
        auxPesadaBruto = GF_PARAMETROS7("pesadaBruto", 0 ,6)
        auxPesadaBrutoOld = GF_PARAMETROS7("pesadaBrutoOld", 0 ,6)
        auxPesadaTara = GF_PARAMETROS7("pesadaTara", 0 ,6)
        auxPesadaTaraOld = GF_PARAMETROS7("pesadaTaraOld", 0 ,6)
        auxMerma = GF_PARAMETROS7("mermaKg", 0 ,6)
        auxMermaOld = GF_PARAMETROS7("mermaKgOld", 0 ,6)
        auxMermaPorcentaje = GF_PARAMETROS7("mermaPorcentaje", "" ,6)
        auxMermaPorcentajeOld = GF_PARAMETROS7("mermaPorcentajeOld", "" ,6)
        auxObservacionesDescarga = GF_PARAMETROS7("observacionesDescarga", "" ,6)
        auxNetoCMerma = Cdbl(auxPesadaBruto) - Cdbl(auxPesadaTara) - Cdbl(auxMerma)
        auxNetoSMerma = Cdbl(auxPesadaBruto) - Cdbl(auxPesadaTara)
    end if
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function clearIntermediarios()
    auxRemitenteCuit1 = ""
    auxRemitenteCuit2 = ""
    auxRemitenteCuit3 = ""
    auxRemitenteCd = 0
    auxRemitenteCdOld = 0
    auxRemitenteCuitOld = ""
    auxRemitenteDs = ""
    auxTitularDs = ""
    auxTitularCd = 0
    auxTitularCdOld = 0
    auxTitularCuitOld = ""
    auxTitularCuit1 = ""
    auxTitularCuit2 = ""
    auxTitularCuit3 = ""
    auxIntermediarioCd = 0
    auxIntermediarioDs = ""
    auxIntermediarioCdOld = 0
    auxIntermediarioCuitOld =""
    auxIntermediarioCuit1 = ""
    auxIntermediarioCuit2 = ""
    auxIntermediarioCuit3 = ""
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
'Carga el Titular, Intermediario y el Remitente comercial
Function cargarIntermediariosCtaPte(pCuitRem,pIdCamion,pCdVendedor)        
    Dim strSQL, myCuit,rsIn
    Call clearIntermediarios()
    strSQL = "SELECT CDVENDEDOR,SQORDEN FROM HCUENTAYORDENESCAMIONES WHERE IDCAMION = '"& pIdCamion & "' AND DTCONTABLE ='"& Left(g_dtContable ,4) &"-"& Mid(g_dtContable ,5,2) &"-"& Right(g_dtContable ,2)&"' ORDER BY SQORDEN "
    Call GF_BD_Puertos(g_pto, rsIn, "OPEN",strSQL)
    if (not rsIn.Eof) then
        auxTitularDs = getDsVendedorPto(pCuitRem,"") 'CON CAMIONES.NUCUITREM OBTENGO LA DESCRIPCION (PUERTOS O BSAS)
        auxTitularCd = getCdVendedorPto(pCuitRem) 'CON CAMIONES.NUCUITREM OBTENGO EL CODIGO (SOLO PUERTOS)
        auxTitularCdOld = auxTitularCd
        auxTitularCuit1 = Left(pCuitRem,2) 'CAMIONES.NUCUITREM
        auxTitularCuit2 = mid(pCuitRem,3,8) 
        auxTitularCuit3 = Right(pCuitRem,1) 
        auxTitularCuitOld = auxTitularCuit1 & auxTitularCuit2 & auxTitularCuit3
        if (Cdbl(rsIn("SQORDEN")) = CUENTAYORDEN_INTERMEDIARIO) then
            ' ORDEN 1
            auxIntermediarioCd = Cdbl(rsIn("CDVENDEDOR")) 'CUENTAYORDENESCAMIONES.CDVENDEDOR (ORDEN 1)
            auxIntermediarioCdOld = auxIntermediarioCd
            auxIntermediarioDs = getDsVendedorPto("",auxIntermediarioCd) 'CON CUENTAYORDENESCAMIONES.CDVENDEDOR (ORDEN 1) OBTENGO LA DESCRIPCION
            myCuit = getCuitVendedorPto(auxIntermediarioCd) 'CON CUENTAYORDENESCAMIONES.CDVENDEDOR (ORDEN 1) OBTENGO EL CUIT
            if (myCuit <> "") then
                auxIntermediarioCuit1 = Left(myCuit,2)
                auxIntermediarioCuit2 = mid(myCuit,3,8)
                auxIntermediarioCuit3 = Right(myCuit,1)
                auxIntermediarioCuitOld = auxIntermediarioCuit1 & auxIntermediarioCuit2 & auxIntermediarioCuit3
            end if
            rsIn.MoveNext()
        end if
        if (not rsIn.Eof) then
            if (Cdbl(rsIn("SQORDEN")) = CUENTAYORDEN_REMITENTE) then
                auxRemitenteCd = Cdbl(rsIn("CDVENDEDOR")) 'CUENTAYORDENESCAMIONES .CDVENDEDOR (ORDEN 2)
                auxRemitenteCdOld = auxRemitenteCd
                auxRemitenteDs = getDsVendedorPto("",auxRemitenteCd) 'CON CUENTAYORDENESCAMIONES.CDVENDEDOR (ORDEN 2) OBTENGO LA DESCRIPCION        
                myCuit = getCuitVendedorPto(auxRemitenteCd) 'CON CUENTAYORDENESCAMIONES.CDVENDEDOR (ORDEN 1) OBTENGO EL CUIT
                if (myCuit <> "") then
                    auxRemitenteCuit1 = Left(myCuit,2)
                    auxRemitenteCuit2 = mid(myCuit,3,8)
                    auxRemitenteCuit3 = Right(myCuit,1)
                    auxRemitenteCuitOld = auxRemitenteCuit1 & auxRemitenteCuit2 & auxRemitenteCuit3
                end if
            end if
        end if
    else
        auxTitularDs = getDsVendedorPto(pCuitRem,pCdVendedor) 'CAMIONES.NUCUITREM Y CAMIONESDESCARGA.CDVENDEDOR
        auxTitularCuit1 = Left(pCuitRem,2) 'CAMIONES.NUCUITREM
        auxTitularCuit2 = mid(pCuitRem,3,8) 'CAMIONES.NUCUITREM
        auxTitularCuit3 = Right(pCuitRem,1) 'CAMIONES.NUCUITREM
        auxTitularCuitOld = auxTitularCuit1 & auxTitularCuit2 & auxTitularCuit3
        auxTitularCd = pCdVendedor 'CAMIONESDESCARGA.CDVENDEDOR
        auxTitularCdOld = auxTitularCd
    end if
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function hayErroresCartaPorte(ByRef p_Error)
    hayErroresCartaPorte = true
    if (errorCabecera = "") then 
        if (errorInterviniente = "") then
            if (errorProducto = "") then
                if (errorTransporte = "") then
                    if (errorDescarga = "") then
                        hayErroresCartaPorte = false
                    else
                        p_Error = errorDescarga
                    end if
                else
                    p_Error = errorTransporte
                end if
            else
                p_Error = errorProducto
            end if
        else
            p_Error = errorInterviniente
        end if
    else
        p_Error = errorCabecera
    end if
End Function 
'-----------------------------------------------------------------------------------------------------------------------------------
'Graba los cambios en caso de que se hallan modificado los datos, la cabecera solo utiliza la tabla HCamionesDescarga
Function grabarCabeceraCtaPte()
    Dim mySet,strSQL
    mySet = ""
    if (Trim(g_ctaPte) <> Trim(auxCartaPorte1) & Trim(auxCartaPorte2)) then 
        Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "Carta de Porte|"&Trim(g_ctaPte)&"|"&Trim(auxCartaPorte1) & Trim(auxCartaPorte2))
        mySet = "NUCARTAPORTE = '" & Trim(auxCartaPorte1) & Trim(auxCartaPorte2) &"',"
        'Actualizo las tablas para que la carta de porte sea la nueva
        Call actualizarCartaPorteGlobales(Trim(auxCartaPorte1) & Trim(auxCartaPorte2),g_ctaPte)
    end if
    if (Trim(auxCTG) <> Trim(auxCTGOld)) then
        Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "CTG|"&Trim(auxCTGOld)&"|"&Trim(auxCTG) )
        mySet = mySet  & "CTG = '" & Trim(auxCTG) &"',"
        auxCTGOld = Trim(auxCTG)
    end if
    if (auxDtCartaPorte <> auxDtCartaPorteOld) then
        Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "FECHA DE CARGA|"&GF_FN2DTE(auxDtCartaPorteOld)&"|"&GF_FN2DTE(auxDtCartaPorte))
        mySet = mySet  & "DTCARTAPORTE = CAST('" & Left(auxDtCartaPorte,4) &"-"& Mid(auxDtCartaPorte,5,2) &"-"& Right(auxDtCartaPorte,2) &"' AS DATE),"
        auxDtCartaPorteOld = auxDtCartaPorte
    end if
    if (auxDtVencimiento <> auxDtVencimientoOld) then
        Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "FECHA DE VENCIMIENTO|"&GF_FN2DTE(auxDtVencimientoOld)&"|"&GF_FN2DTE(auxDtVencimiento))
        mySet = mySet  & "DTCPVENCIMIENTO = CAST('" & Left(auxDtVencimiento,4) &"-"& Mid(auxDtVencimiento,5,2) &"-"& Right(auxDtVencimiento,2) &"' AS DATE),"
        auxDtVencimientoOld = auxDtVencimiento
    end if
    if (oDiccModificaciones.Count <> 0) then
        mySet = left(mySet,len(mySet)-1)
        strSQL = "UPDATE HCAMIONESDESCARGA SET "& mySet & " WHERE NUCARTAPORTE = '"& g_ctaPte &"' AND DTCONTABLE = '"& Left(g_dtContable ,4) &"-"& Mid(g_dtContable ,5,2) &"-"& Right(g_dtContable ,2)&"'"
        Call logMig.info("Modifico la Cabecera: " & strSQL)
        Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
        g_ctaPte = Trim(auxCartaPorte1) & Trim(auxCartaPorte2)
    end if
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function actualizarCartaPorteGlobales(p_CartaPorteNew, p_CartaPorteOld)
    Dim strSQL
    Call logMig.info("Modifico la Carta de Porte: " & p_CartaPorteOld)
    strSQL = "UPDATE DATOSONCCA SET NCCARTAPORTE ='"& p_CartaPorteNew &"' WHERE NCCARTAPORTE='" & p_CartaPorteOld & "'"
	Call logMig.info("-----> 1) " & strSQL)
    Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
    

    strSQL = "UPDATE TBLBIOTECNOLOGIASDECLARADAS SET NUCARTAPORTE = '"& p_CartaPorteNew &"' WHERE NUCARTAPORTE = '"& p_CartaPorteOld &"' AND TIPOTRANSPORTE = "& TIPO_TRANSPORTE_CAMION
	Call logMig.info("-----> 2) " & strSQL)
    Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
    

    strSQL = "UPDATE HAUDCAMIONESDESCARGA SET NUCARTAPORTE = '"& p_CartaPorteNew &"' WHERE NUCARTAPORTE = '"& p_CartaPorteOld &"'"
	Call logMig.info("-----> 3) " & strSQL)
    Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
    
	
	strSQL = "Update STICKERSCAMARA set NUCARTAPORTE='" & p_CartaPorteNew & "' where NUCARTAPORTE='" & p_CartaPorteOld & "'  AND TIPOTRANSPORTE = "& TIPO_TRANSPORTE_CAMION	
	Call logMig.info("-----> 4) " & strSQL)
    Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
    
	Call logMig.info(" ---- Fin actualizarCartaPorteGlobales ----" & strSQL)
	
End function
'-----------------------------------------------------------------------------------------------------------------------------------
Function grabarIntervinienteCtaPte()
    Dim mySet,strSQL, auxCdInterviniente,dsChofer
    mySet = ""
    'Tiene que ser distinto los dos(codigo y cuit) por que puede cambiar el cdvendedor pero puede tener el mismo cuit que el anterior
    if ((Cdbl(auxTitularCd) <> Cdbl(auxTitularCdOld))or(Trim(auxTitularCuit1)&Trim(auxTitularCuit2)&Trim(auxTitularCuit3) <> Trim(auxTitularCuitOld))) then
        'Si hubo un cambio en Titular, verifico tambien el Interviniente y el Remitente para ver en donde guardo el cambio
        strSQL = "UPDATE HCAMIONES SET NUCUITREM = '"& auxTitularCuit1 & auxTitularCuit2 & auxTitularCuit3 &"' WHERE IDCAMION = '"& g_IdCamion &"' AND DTCONTABLE = '"& Left(g_dtContable ,4) &"-"& Mid(g_dtContable ,5,2) &"-"& Right(g_dtContable ,2)&"'"
        Call logMig.info("1 - Modifico el Titular: " & strSQL)
        Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
        Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "TITULAR|"& GF_STR2CUIT(auxTitularCuitOld) &" ("& auxTitularCdOld &")|"& auxTitularCuit1 &"-"& auxTitularCuit2 &"-"& auxTitularCuit3 &" ("&auxTitularCd&") ")
        if ((Cdbl(auxRemitenteCd) = 0)and(Cdbl(auxIntermediarioCd) = 0)) then
            'Si no tiene Remitente y Intermediario guardo en HCAMIONESDESCARGA el CdVendedor. En caso que venga de BS.AS no tiene codigo (-1), por lo tanto grabo 0
            auxCdInterviniente = auxTitularCd
            if (Cdbl(auxCdInterviniente) < 0) then auxCdInterviniente = 0 
            strSQL = "UPDATE HCAMIONESDESCARGA SET CDVENDEDOR = "& auxCdInterviniente &" WHERE NUCARTAPORTE = '"& g_ctaPte &"' AND DTCONTABLE = '"& Left(g_dtContable ,4) &"-"& Mid(g_dtContable ,5,2) &"-"& Right(g_dtContable ,2)&"'"
            Call logMig.info("2 - Modifico el Titular: " & strSQL)
            Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
        end if
        auxTitularCdOld = auxTitularCd
        auxTitularCuitOld = Trim(auxTitularCuit1)&Trim(auxTitularCuit2)&Trim(auxTitularCuit3) 
    end if
    if ((Cdbl(auxIntermediarioCd) <> Cdbl(auxIntermediarioCdOld))or(Trim(auxIntermediarioCuit1)&Trim(auxIntermediarioCuit2)&Trim(auxIntermediarioCuit3) <> Trim(auxIntermediarioCuitOld))) then
        'Controlo si agrega un intermediario o lo borra
        if (Cdbl(auxIntermediarioCd) <> 0) then
            'Verifico si el camion tiene un registro en la tabla CUENTAYORDENESCAMIONES
            Call GF_BD_Puertos(g_pto, rsCOI, "OPEN", "SELECT * FROM HCUENTAYORDENESCAMIONES WHERE IDCAMION = '"& g_IdCamion &"' AND DTCONTABLE = '"& Left(g_dtContable ,4) &"-"& Mid(g_dtContable ,5,2) &"-"& Right(g_dtContable ,2)&"' AND SQORDEN = "&CUENTAYORDEN_INTERMEDIARIO)
            if not rsCOI.Eof then
                strSQL = "UPDATE HCUENTAYORDENESCAMIONES SET CDVENDEDOR = "& auxIntermediarioCd &" WHERE IDCAMION ='"& g_IdCamion &"' AND DTCONTABLE = '"& Left(g_dtContable ,4) &"-"& Mid(g_dtContable ,5,2) &"-"& Right(g_dtContable ,2)&"' AND SQORDEN ="&CUENTAYORDEN_INTERMEDIARIO
                Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "INTERMEDIRARIO|"& GF_STR2CUIT(auxIntermediarioCuitOld) &" ("& auxIntermediarioCdOld & " - " & getDsCorredor(auxIntermediarioCdOld) & ")|"& auxIntermediarioCuit1 &"-"& auxIntermediarioCuit2 &"-"& auxIntermediarioCuit3 &" ("&auxIntermediarioCd & " - " & getDsCorredor(auxIntermediarioCd) & ")")
            else
                strSQL = "INSERT INTO HCUENTAYORDENESCAMIONES (DTCONTABLE,IDCAMION,SQORDEN,CDVENDEDOR) VALUES ('"& Left(g_dtContable ,4) &"-"& Mid(g_dtContable ,5,2) &"-"& Right(g_dtContable ,2)&"','"& g_IdCamion &"',"&CUENTAYORDEN_INTERMEDIARIO&","& auxIntermediarioCd & ")"
                Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "INTERMEDIRARIO|Ninguno|"& auxIntermediarioCuit1 &"-"& auxIntermediarioCuit2 &"-"& auxIntermediarioCuit3 &" ("&auxIntermediarioCd& ")")
            end if
            Call logMig.info("1 - Modifico el Intermediario: " & strSQL)
            Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
            'Si no encuentro Remitente guardo como vendedor al Intermediario
            if (Cdbl(auxRemitenteCd) = 0) then
                strSQL = "UPDATE HCAMIONESDESCARGA SET CDVENDEDOR = "& auxIntermediarioCd &" WHERE NUCARTAPORTE = '"& g_ctaPte &"' AND DTCONTABLE = '"& Left(g_dtContable ,4) &"-"& Mid(g_dtContable ,5,2) &"-"& Right(g_dtContable ,2)&"'"
                Call logMig.info("2 - Modifico el Intermediario: " & strSQL)
                Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
            end if
            auxIntermediarioCuitOld = Trim(auxIntermediarioCuit1) & Trim(auxIntermediarioCuit2) & Trim(auxIntermediarioCuit3)
        else
            strSQL = "DELETE FROM HCUENTAYORDENESCAMIONES WHERE DTCONTABLE = '"& Left(g_dtContable ,4) &"-"& Mid(g_dtContable ,5,2) &"-"& Right(g_dtContable ,2)&"' AND IDCAMION ='"& g_IdCamion &"' AND SQORDEN ="&CUENTAYORDEN_INTERMEDIARIO
            Call logMig.info("3 - Modifico el Intermediario: " & strSQL)
            Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
            Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "INTERMEDIRARIO|"& GF_STR2CUIT(auxIntermediarioCuitOld) &" ("& auxIntermediarioCdOld & " - " & getDsCorredor(auxIntermediarioCdOld) & ")|Ninguno")
            auxIntermediarioCuitOld = ""
        end if
        auxIntermediarioCdOld = auxIntermediarioCd
    end if
    if ((Cdbl(auxRemitenteCd) <> Cdbl(auxRemitenteCdOld))or(Trim(auxRemitenteCuit1)&Trim(auxRemitenteCuit2)&Trim(auxRemitenteCuit3) <> Trim(auxRemitenteCuitOld))) then
        if (Cdbl(auxRemitenteCd) <> 0) then
            'Verifico si el camion tiene un registro en la tabla CUENTAYORDENESCAMIONES
            Call GF_BD_Puertos(g_pto, rsCOR, "OPEN", "SELECT * FROM HCUENTAYORDENESCAMIONES WHERE DTCONTABLE = '"& Left(g_dtContable ,4) &"-"& Mid(g_dtContable ,5,2) &"-"& Right(g_dtContable ,2)&"' AND IDCAMION = '"& g_IdCamion &"' AND SQORDEN = "& CUENTAYORDEN_REMITENTE)
            if not rsCOR.Eof then
                strSQL = "UPDATE HCUENTAYORDENESCAMIONES SET CDVENDEDOR = "& auxRemitenteCd &" WHERE DTCONTABLE = '"& Left(g_dtContable ,4) &"-"& Mid(g_dtContable ,5,2) &"-"& Right(g_dtContable ,2)&"' AND IDCAMION ='"& g_IdCamion &"' AND SQORDEN ="& CUENTAYORDEN_REMITENTE
                Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "REMITENTE|"& GF_STR2CUIT(auxRemitenteCuitOld) &" ("& auxRemitenteCdOld & " - " & getDsCorredor(auxRemitenteCdOld) & ")|"& auxRemitenteCuit1 &"-"& auxRemitenteCuit2 &"-"& auxRemitenteCuit3 &" ("& auxRemitenteCd & " - " & getDsCorredor(auxRemitenteCd) & ")")
            else
                strSQL = "INSERT INTO HCUENTAYORDENESCAMIONES (DTCONTABLE,IDCAMION,SQORDEN,CDVENDEDOR) VALUES ('"& Left(g_dtContable ,4) &"-"& Mid(g_dtContable ,5,2) &"-"& Right(g_dtContable ,2)&"','"& g_IdCamion &"',"& CUENTAYORDEN_REMITENTE &","& auxRemitenteCd &")"
                Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "REMITENTE|Ninguno|"& auxRemitenteCuit1 &"-"& auxRemitenteCuit2 &"-"& auxRemitenteCuit3 &" ("& auxRemitenteCd & " - " & getDsCorredor(auxRemitenteCd) & ")")
            end if
            Call logMig.info("1 - Modifico el Remitente: " & strSQL)
            Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
            'Si tiene un Remitente guardo a este ultimo como vendedor
            strSQL = "UPDATE HCAMIONESDESCARGA SET CDVENDEDOR = "& auxRemitenteCd &" WHERE NUCARTAPORTE = '"& g_ctaPte &"' AND DTCONTABLE = '"& Left(g_dtContable ,4) &"-"& Mid(g_dtContable ,5,2) &"-"& Right(g_dtContable ,2)&"'"
            Call logMig.info("2 - Modifico el Remitente: " & strSQL)
            Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
            auxRemitenteCuitOld = Trim(auxRemitenteCuit1) & Trim(auxRemitenteCuit2) & Trim(auxRemitenteCuit3)
        else
            strSQL = "DELETE FROM HCUENTAYORDENESCAMIONES WHERE DTCONTABLE = '"& Left(g_dtContable ,4) &"-"& Mid(g_dtContable ,5,2) &"-"& Right(g_dtContable ,2)&"' AND IDCAMION ='"& g_IdCamion &"' AND SQORDEN ="& CUENTAYORDEN_REMITENTE
            Call logMig.info("3 - Modifico el Remitente: " & strSQL)
            Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
            Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "REMITENTE|"& GF_STR2CUIT(auxRemitenteCuitOld) &" ("& auxRemitenteCdOld & " - " & getDsCorredor(auxRemitenteCdOld) & ")|Ninguno ")
            auxRemitenteCuitOld = ""
            if (Cdbl(auxIntermediarioCd) <> 0) then
                'Si borro el remitente y tiene Intermediario debo delegar el cdvendedor de HcamionesDescarga a este.
                strSQL = "UPDATE HCAMIONESDESCARGA SET CDVENDEDOR = "& auxIntermediarioCd &" WHERE NUCARTAPORTE = '"& g_ctaPte &"' AND DTCONTABLE = '"& Left(g_dtContable ,4) &"-"& Mid(g_dtContable ,5,2) &"-"& Right(g_dtContable ,2)&"'"
            else
                'Si borro el remitente pero no tingo Intermediario debo delegar el cdvendedor de HcamionesDescarga a el Titular.
                auxCdInterviniente = auxTitularCd
                if (Cdbl(auxCdInterviniente) < 0) then auxCdInterviniente = 0 
                strSQL = "UPDATE HCAMIONESDESCARGA SET CDVENDEDOR = "& auxCdInterviniente &" WHERE NUCARTAPORTE = '"& g_ctaPte &"' AND DTCONTABLE = '"& Left(g_dtContable ,4) &"-"& Mid(g_dtContable ,5,2) &"-"& Right(g_dtContable ,2)&"'"
            end if
            Call logMig.info("4 - Modifico el Remitente: " & strSQL)
            Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
        end if
        auxRemitenteCdOld = auxRemitenteCd
    end if
    if ((Cdbl(auxCorredor) <> Cdbl(auxCorredorOld))or(Trim(auxCorredorCuit1)&Trim(auxCorredorCuit2)&Trim(auxCorredorCuit3) <> Trim(auxCorredorCuitOld))) then
        strSQL = "UPDATE HCAMIONESDESCARGA SET CDCORREDOR = "& auxCorredor &" WHERE NUCARTAPORTE = '"& g_ctaPte &"' AND DTCONTABLE = '"& Left(g_dtContable ,4) &"-"& Mid(g_dtContable ,5,2) &"-"& Right(g_dtContable ,2)&"'"
        Call logMig.info("1 - Modifico el Corredor: " & strSQL)
        Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
        Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "CORREDOR|"& GF_STR2CUIT(auxCorredorCuitOld) &" ("& auxCorredorOld & " - " & getDsCorredor(auxCorredorOld) & ")|"& Trim(auxCorredorCuit1)&"-"&Trim(auxCorredorCuit2)&"-"&Trim(auxCorredorCuit3) &" ("& auxCorredor & " - " & getDsCorredor(auxCorredor) & ")")
        auxCorredorOld = auxCorredor
        auxCorredorCuitOld = Trim(auxCorredorCuit1) & Trim(auxCorredorCuit2) & Trim(auxCorredorCuit3)
    end if
    if ((Cdbl(auxEntregador) <> Cdbl(auxEntregadorOld))or(Trim(auxEntregadorCuit1)&Trim(auxEntregadorCuit2)&Trim(auxEntregadorCuit3) <> Trim(auxEntregadorCuitOld))) then
        strSQL = "UPDATE HCAMIONESDESCARGA SET CDENTREGADOR = "& auxEntregador &" WHERE NUCARTAPORTE = '"& g_ctaPte &"' AND DTCONTABLE = '"& Left(g_dtContable ,4) &"-"& Mid(g_dtContable ,5,2) &"-"& Right(g_dtContable ,2)&"'"
        Call logMig.info("1 - Modifico el Entregador: " & strSQL)
        Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
        Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "ENTREGADOR|"& GF_STR2CUIT(auxEntregadorCuitOld) &" ("& auxEntregadorOld & " - " & getDsEntregador(auxEntregadorOld) & ")|"& Trim(auxEntregadorCuit1)&"-"&Trim(auxEntregadorCuit2)&"-"&Trim(auxEntregadorCuit3) &" ("& auxEntregador & " - " & getDsEntregador(auxEntregador) & ")")
        auxEntregadorOld = auxEntregador
        auxEntregadorCuitOld = Trim(auxEntregadorCuit1) & Trim(auxEntregadorCuit2) & Trim(auxEntregadorCuit3)
    end if
    if ((Cdbl(auxDestinatario) <> Cdbl(auxDestinatarioOld))or(Trim(auxDestinatarioCuit1)&Trim(auxDestinatarioCuit2)&Trim(auxDestinatarioCuit3) <> Trim(auxDestinatarioCuitOld))) then
        strSQL = "UPDATE HCAMIONESDESCARGA SET CDCLIENTE = "& auxDestinatario &" WHERE NUCARTAPORTE = '"& g_ctaPte &"' AND DTCONTABLE = '"& Left(g_dtContable ,4) &"-"& Mid(g_dtContable ,5,2) &"-"& Right(g_dtContable ,2)&"'"
        Call logMig.info("1 - Modifico el Destinatario: " & strSQL)
        Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
        Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "DESTINATARIO|"& GF_STR2CUIT(auxDestinatarioCuitOld) &" ("& auxDestinatarioOld & " - " & getDsCliente(auxDestinatarioOld) & ")|"& Trim(auxDestinatarioCuit1)&"-"&Trim(auxDestinatarioCuit2)&"-"&Trim(auxDestinatarioCuit3) &" ("& auxDestinatario & " - " & getDsCliente(auxDestinatario) & ")")
    end if    
    if ((Cdbl(auxTransportista) <> Cdbl(auxTransportistaOld))or(Trim(auxTransportistaCuit1)&Trim(auxTransportistaCuit2)&Trim(auxTransportistaCuit3) <> Trim(auxTransportistaCuitOld))) then
        strSQL = "UPDATE HCAMIONES SET CDTRANSPORTISTA = "& auxTransportista &" WHERE IDCAMION = '"& g_IdCamion &"' AND DTCONTABLE = '"& Left(g_dtContable ,4) &"-"& Mid(g_dtContable ,5,2) &"-"& Right(g_dtContable ,2)&"'"
        Call logMig.info("1 - Modifico el Transportista: " & strSQL)
        Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
        Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "TRANSPORTISTA|"& GF_STR2CUIT(auxTransportistaCuitOld) &" ("& auxTransportistaOld & " - " & getDsTransportista(auxTransportistaOld) & ")|"& Trim(auxTransportistaCuit1)&"-"&Trim(auxTransportistaCuit2)&"-"&Trim(auxTransportistaCuit3) &" ("& auxTransportista & " - " & getDsTransportista(auxTransportista) & ")")
        auxTransportistaOld = auxTransportista
        auxTransportistaCuitOld = Trim(auxTransportistaCuit1)&Trim(auxTransportistaCuit2) & Trim(auxTransportistaCuit3)
    end if

     if (Trim(auxChoferNumDoc1)&Trim(auxChoferNumDoc2)&Trim(auxChoferNumDoc3) <> Trim(auxChoferCuitOld)) then
    
        dsChofer = Split(auxChoferDs,",")
        strSQL = "UPDATE HCAMIONES SET NUDOCUMENTO="&Trim(auxChoferNumDoc2)&"," &_
                 " DSAPELLIDOCONDUCTOR='"& dsChofer(0) &"',DSNOMBRECONDUCTOR='"&dsChofer(1)&"'" &_
                 " WHERE IDCAMION = '"& g_IdCamion &"' AND DTCONTABLE = '"& Left(g_dtContable ,4) &"-"& Mid(g_dtContable ,5,2) &"-"& Right(g_dtContable ,2)&"'"
        Call logMig.info("1 - Modifico el Chofer: " & strSQL)
        Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
        Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "CHOFER|"& GF_STR2CUIT(auxChoferCuitOld) &"|"& Trim(auxChoferNumDoc1)&"-"&Trim(auxChoferNumDoc2)&"-"&Trim(auxChoferNumDoc3))
        auxChoferCuitOld = Trim(auxChoferNumDoc1) & Trim(auxChoferNumDoc2) & Trim(auxChoferNumDoc3)
    end if
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function grabarProductoCtaPte()
    Dim strSQL, mySet
    mySet = ""
    'Agrupo los campos a modificar por las tablas implicadas
    if (Cdbl(auxGrano) <> Cdbl(auxGranoOld)) then
        Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "PRODUCTO|"& getDsProducto(auxGranoOld) &"|"& getDsProducto(auxGrano))
        mySet = "CDPRODUCTO = "& auxGrano &","
    end if
    if (auxCupo <> auxCupoOld) then
        Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "CUPO|"& auxCupoOld &"|"& auxCupo )
        mySet = mySet & "NUCUPO = '"& auxCupo &"',"
        'Dependiendo del numero de cupo si se agrega o no se van actualizando la tabla cuposAsignados
        if (Cdbl(auxCupoOld) <> 0) then Call actualizarCupoAsignado(auxCupo, OPERATOR_RESTA)
        if (Cdbl(auxCupo) <> 0) then Call actualizarCupoAsignado(auxCupo, OPERATOR_SUMA)
        auxCupoOld = auxCupo
    end if
    if (mySet <> "") then        
        mySet = left(mySet,len(mySet)-1)
        strSQL = "UPDATE HCAMIONES SET "& mySet & " WHERE IDCAMION = '"& g_IdCamion &"' AND DTCONTABLE = '"& Left(g_dtContable ,4) &"-"& Mid(g_dtContable ,5,2) &"-"& Right(g_dtContable ,2)&"'"
        Call logMig.info("1 - Modifico datos de los granos: " & strSQL)
        Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
    end if
    mySet = ""
    if (auxCosecha <> auxCosechaOld) then
        Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "COSECHA|"& auxCosechaOld &"|"& auxCosecha )
        mySet = mySet & "CDCOSECHA = "& auxCosecha &","
        auxCosechaOld = auxCosecha
    end if
    if (Cdbl(auxPesoBruto) <> Cdbl(auxPesoBrutoOld)) then
        Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "PESO BRUTO PROCEDENCIA|"& auxPesoBrutoOld &"|"& auxPesoBruto )
        mySet = mySet & "VLBRUTOORIGEN = "& Cdbl(auxPesoBruto) &","
        auxPesoBrutoOld = auxPesoBruto
    end if
    if (Cdbl(auxPesoTara) <> Cdbl(auxPesoTaraOld)) then
        Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "PESO TARA PROCEDENCIA|"& auxPesoTaraOld &"|"& auxPesoTara )
        mySet = mySet & "VLTARAORIGEN = "& Cdbl(auxPesoBruto) &","
        auxPesoTaraOld = auxPesoTara
    end if
    if (Cdbl(auxProcedenciaCd) <> Cdbl(auxProcedenciaCdOld)) then
        Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "PROCEDENCIA|"& auxProcedenciaCdOld &"|"& auxProcedenciaCd )
        mySet = mySet & "CDPROCEDENCIA = "& Cdbl(auxProcedenciaCd) &","
        auxProcedenciaCdOld = auxProcedenciaCd
    end if
    if (mySet <> "") then
        mySet = left(mySet,len(mySet)-1)
        strSQL = "UPDATE HCAMIONESDESCARGA SET "& mySet & " WHERE NUCARTAPORTE = '"& g_ctaPte &"' AND DTCONTABLE = '"& Left(g_dtContable ,4) &"-"& Mid(g_dtContable ,5,2) &"-"& Right(g_dtContable ,2)&"'"
        Call logMig.info("2 - Modifico datos de los granos: " & strSQL)
        Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
    end if
    mySet = ""
    if (editText4DB(Trim(Ucase(auxObservaciones))) <> editText4DB(Trim(Ucase(auxObservacionesOld)))) then
        'Verifico el camion ya tenia una descripcion
        Call GF_BD_Puertos(g_pto, rsObs, "OPEN", "SELECT * FROM HOBSERVACIONESCAMION WHERE IDCAMION = '"& g_IdCamion &"' AND DTCONTABLE = '"& Left(g_dtContable ,4) &"-"& Mid(g_dtContable ,5,2) &"-"& Right(g_dtContable ,2) &"'")
        if (not rsObs.Eof) then
            strSQL = "UPDATE HOBSERVACIONESCAMION SET DSOBSERVACIONES = '"& editText4DB(Trim(auxObservaciones)) &"' WHERE IDCAMION = '"& g_IdCamion &"' AND DTCONTABLE = '"& Left(g_dtContable ,4) &"-"& Mid(g_dtContable ,5,2) &"-"& Right(g_dtContable ,2) &"'"
        else
            strSQL = "INSERT INTO HOBSERVACIONESCAMION (DTCONTABLE,IDCAMION,DSOBSERVACIONES)VALUES('"& Left(g_dtContable ,4) &"-"& Mid(g_dtContable ,5,2) &"-"& Right(g_dtContable ,2) &"','"&g_IdCamion&"','"&editText4DB(Trim(auxObservaciones))&"')"
        end if
        Call logMig.info("1 - Modifico observaciones: " & strSQL)
        Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
        Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "OBSERVACIONES CAMION|"& Trim(auxObservacionesOld) &"|"& Trim(auxObservaciones))
        auxObservacionesOld = auxObservaciones
    end if
    if (Cdbl(auxBiotecnologia)<>Cdbl(auxBiotecnologiaOld)) then
        Call GF_BD_Puertos(g_pto, rsBio, "OPEN", "SELECT * FROM TBLBIOTECNOLOGIASDECLARADAS WHERE TIPOTRANSPORTE="& TIPO_TRANSPORTE_CAMION &" AND NUCARTAPORTE = '"& g_ctaPte &"'")
        if (not rsBio.Eof) then
            strSQL = "UPDATE TBLBIOTECNOLOGIASDECLARADAS SET IDBIOTECNOLOGIA = "& auxBiotecnologia &" WHERE TIPOTRANSPORTE="& TIPO_TRANSPORTE_CAMION &" AND NUCARTAPORTE = '"& g_ctaPte &"'"
        else
            strSQL = "INSERT INTO TBLBIOTECNOLOGIASDECLARADAS (TIPOTRANSPORTE,NUCARTAPORTE,IDBIOTECNOLOGIA)VALUES("& TIPO_TRANSPORTE_CAMION &",'"& g_ctaPte &"',"& auxBiotecnologia &")"
        end if
        Call logMig.info("1 - Modifico Biotecnologia: " & strSQL)
        Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
        Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "BIOTECNOLOGIA|"& auxBiotecnologiaOld&"|"& auxBiotecnologia )
        auxBiotecnologiaOld = auxBiotecnologia
    end if
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function grabarTransporteCtaPte()
    Dim strSQL, mySet
    mySet = ""
    if (Trim(Ucase(auxChapa)) <> auxChapaOld) then
        Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "CHAPA CAMION|"& auxChapaOld &"|"& Trim(Ucase(auxChapa)))
        mySet = "CDCHAPACAMION = '"& Trim(Ucase(auxChapa)) &"',"
        auxChapaOld = Trim(Ucase(auxChapa))
    end if
    if (Trim(Ucase(auxAcoplado)) <> auxAcopladoOld) then
        Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "ACOPLADO CAMION|"& auxAcopladoOld &"|"& Trim(Ucase(auxAcoplado)))
        mySet = mySet & "CDCHAPAACOPLADO = '"& Trim(Ucase(auxAcoplado)) &"',"
        auxAcopladoOld = Trim(Ucase(auxAcoplado))
    end if
    if (mySet <> "") then
		mySet = left(mySet,len(mySet)-1)
        strSQL = "UPDATE HCAMIONES SET "& mySet & " WHERE IDCAMION = '"& g_IdCamion &"' AND DTCONTABLE = '"& Left(g_dtContable ,4) &"-"& Mid(g_dtContable ,5,2) &"-"& Right(g_dtContable ,2)&"'"
        Call logMig.info("1 - Modifico Transporte: " & strSQL)
        Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
    end if
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function grabarDescargaCtaPte()
    Dim strSQL, mySet, myFechaArribo, myFechaEgreso,myDtContable
    mySet = ""
    if (Cdbl(auxTurno) <> Cdbl(auxTurnoOld)) then
        strSQL = "UPDATE HCAMIONES SET SQTURNO = "& Cdbl(auxTurno) &" WHERE IDCAMION = '"& g_IdCamion &"' AND DTCONTABLE = '"& Left(g_dtContable ,4) &"-"& Mid(g_dtContable ,5,2) &"-"& Right(g_dtContable ,2)&"'"
        Call logMig.info("1 - Modifico Turno: " & strSQL)
        Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
        Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "TURNO|"& auxTurnoOld &"|"& auxTurno)
        auxTurnoOld = auxTurno
    end if
    '******************************************************** aca graba los nuevos datos de kilos ********************************************************
    if ((Cdbl(auxPesadaTara) <> Cdbl(auxPesadaTaraOld))or(Cdbl(auxPesadaBruto) <> Cdbl(auxPesadaBrutoOld))) then
        myDtContable = GF_FN2DTCONTABLE(g_dtContable)    
        'Primero debo actualizar la tabla HAcondicProductoCamiones el cual indica la merma que debera tener. Ojo esta tabla tambien
        ' se deberia cambiar en caso que se modifique la merma en el POP Up
        Call actualizarAcondProductoMerma(auxMermaPorcentaje,auxNetoSMerma,g_IdCamion,myDtContable)
        if (Cdbl(auxPesadaBruto) <> Cdbl(auxPesadaBrutoOld)) then
            Call agregarPesada(g_IdCamion, myDtContable, auxPesadaBruto, PESADA_BRUTO, g_pto)
            Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "PESADA BRUTO|"& auxPesadaBrutoOld &"|"& auxPesadaBruto)
        end if
        if (Cdbl(auxPesadaTara) <> Cdbl(auxPesadaTaraOld)) then
            Call agregarPesada(g_IdCamion, myDtContable, auxPesadaTara, PESADA_TARA, g_pto)
            Call oDiccModificaciones.Add(oDiccModificaciones.Count + 1, "PESADA TARA|"& auxPesadaTaraOld &"|"& auxPesadaTara)
        end if
        Call actualizarMerma(myDtContable,g_IdCamion,auxMerma,g_pto)
    end if
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
'Funcion encargada de agregar una pesada a un camion
'Nota: el puesto que se registra en la modificacion con 3 siempre y ICMETODO es M, ya en el RMD lo hace asi
Function agregarPesada(p_IdCamion, p_dtContable, p_PesadaBruto, p_TipoPesada, p_strPuerto)
    Dim strSQL,auxDtPesada
    auxDtPesada = Year(Now())&"-"&GF_nDigits(Month(Now()),2)&"-"&GF_nDigits(Day(Now()),2)&" "&GF_nDigits(Hour(Now()),2)&":"&GF_nDigits(Minute(Now()),2)&":"&GF_nDigits(Second(Now()),2)
    strSQL = "INSERT INTO DBO.HPESADASCAMION "&_
             "VALUES ('"& p_dtContable &"','"& p_IdCamion &"',(SELECT MAX(SQPESADA)+1 AS SQPESADA FROM HPESADASCAMION WHERE DTCONTABLE='"&p_dtContable&"' AND IDCAMION='"& p_IdCamion&"'), "&_
             "         "& p_TipoPesada &","& p_PesadaBruto &",'M','3','"& Session("Usuario") &"','"& auxDtPesada &"',0)"
    Call logMig.info("Agrega HPESADASCAMION: "& strSQL)
    Call GF_BD_Puertos(p_strPuerto, rs, "EXEC", strSQL)
End function
'-----------------------------------------------------------------------------------------------------------------------------------
Function agregarAuditoriaPesada(p_IdCamion, p_dtContable, p_PesadaBruto, p_TipoPesada)
    Dim strSQL,auxDtPesada
    auxDtPesada = Year(Now())&"-"&GF_nDigits(Month(Now()),2)&"-"&GF_nDigits(Day(Now()),2)&" "&GF_nDigits(Hour(Now()),2)&":"&GF_nDigits(Minute(Now()),2)&":"&GF_nDigits(Second(Now()),2)
    strSQL = "INSERT INTO DBO.HAUDPESADASCAMION "&_
             "VALUES ('"& p_dtContable &"',(SELECT case when MAX(SQAUDITORIA) is Null then 1 else MAX(SQAUDITORIA)+1 end AS SQAUDITORIA FROM HAUDPESADASCAMION WHERE DTCONTABLE='"&p_dtContable&"' AND IDCAMION='"& p_IdCamion&"'),'"& p_IdCamion &"', "&_
             "         (SELECT MAX(SQPESADA)AS SQPESADA FROM HPESADASCAMION WHERE DTCONTABLE='"&p_dtContable&"' AND IDCAMION='"& p_IdCamion&"'),"& p_TipoPesada &","& p_PesadaBruto &",'M','3','"& Session("Usuario") &"','"& auxDtPesada &"',0)"
    Call GF_BD_Puertos(g_strPuerto, rs, "EXEC", strSQL)
End function
'-----------------------------------------------------------------------------------------------------------------------------------
Function agregarHAudMermas(p_IdCamion, p_dtContable,p_strPuerto)
    Dim strSQL
    strSQL = "Insert into dbo.HAudMermasCamiones "&_
             " select A.dtcontable,(SELECT case when MAX(SQAUDITORIA) is Null then 1 else MAX(SQAUDITORIA)+1 end AS SQAUDITORIA FROM HAudMermasCamiones WHERE DTCONTABLE='"&p_dtContable&"' AND IDCAMION='"& p_IdCamion&"'),A.idcamion,A.sqpesada,A.vlmermakilos "&_
             " from HMERMASCAMIONES A where A.DtContable='" & p_dtContable & "' and A.idCamion='"& p_IdCamion &"'"&_
             "   AND A.sqpesada=(SELECT MAX(SQPESADA) FROM HPESADAsCAMION WHERE IDCAMION = A.IDCAMION AND DTCONTABLE = A.DTCONTABLE)"
    Call logMig.info("Agrega HAudMermasCamiones: "& strSQL)
    Call GF_BD_Puertos(p_strPuerto, rs, "EXEC",strSQL)    
End function
'-----------------------------------------------------------------------------------------------------------------------------------
Function actualizarMerma(p_DtContable,p_idCamion,p_MermaKilo,p_strPuerto)
    Dim strSQL
    strSQL = "INSERT INTO HMERMASCAMIONES VALUES('"& p_DtContable &"','"&p_idCamion&"',(SELECT MAX(SQPESADA) FROM HPESADASCAMION WHERE IDCAMION='"&p_idCamion&"' AND DTCONTABLE='"&p_DtContable&"'),"&p_MermaKilo&")"
    Call logMig.info("Agrega HMERMASCAMIONES: "& strSQL )
    Call GF_BD_Puertos(p_strPuerto, rs, "EXEC",strSQL)
End function
'-----------------------------------------------------------------------------------------------------------------------------------
Function  actualizarAcondProductoMerma(p_PorcentajeMerma,p_NetoSMerma,p_IdCamion,p_DtContable)
    Dim strSQL, mermaAcond
    mermaAcond = Cdbl((p_NetoSMerma * p_PorcentajeMerma / 100))
    strSQL = " UPDATE HAPC " &_
             " SET HAPC.PCMERMA = " & p_PorcentajeMerma & ", HAPC.VLMERMAKILOS=" & mermaAcond &_ 
             " FROM DBO.HACONDICPRODUCTOCAMIONES HAPC "&_
             " WHERE HAPC.IDCAMION='" & p_IdCamion & "' AND HAPC.DTCONTABLE='" & p_DtContable & "'" &_
             "      AND HAPC.SQCALADA = (SELECT MAX(SQCALADA) " &_
             "                           FROM DBO.HCALADADECAMIONES " &_
             "                           WHERE IDCAMION = HAPC.IDCAMION AND DTCONTABLE = HAPC.DTCONTABLE)"
    Call GF_BD_Puertos(g_strPuerto, rs, "EXEC", strSQL)
    Call logMig.info("2 - Modifico Acond. Producto Merma: " & strSQL)
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
'Envia todos los cambios guardados en el Diccionario, sus valores antes y despues
Function enviarMailCtaPte()
    Dim msj,strDs,mailDestino, mailOrigen
    Call logMig.info("Preparando para enviar mail: ")
    msj = "El usuario " & getUserDescription(session("Usuario")) & " ("& session("Usuario") &") realizo cambios en una Carta de Porte: " & vbCrLf
    msj = msj & "Numero   : " & GF_EDIT_CTAPTE(g_ctaPteOld) & vbCrLf
    msj = msj & "Producto : " & auxGranoOld & "-" & getDsProducto(auxGranoOld) & vbCrLf    
    msj = msj & "Fecha    : " & GF_FN2DTCONTABLE(g_dtContable) & vbCrLf
    msj = msj & "Titular       : " & GF_STR2CUIT(auxTitularCuitOld) & "-" & getDsVendedorPto(auxTitularCuitOld,"") & vbCrLf
    msj = msj & "Intermediario : " & auxIntermediarioCdOld & "-" & getDsCorredor(auxIntermediarioCdOld) & vbCrLf
    msj = msj & "Rte. Comercial: " & auxRemitenteCdOld  & "-" & getDsCorredor(auxRemitenteCdOld) & vbCrLf
    msj = msj & "Destinatario  : " & auxDestinatarioOld & "-" & getDsCliente(auxDestinatarioOld) & vbCrLf
    msj = msj & "Corredor      : " & auxCorredorOld & "-" & getDsCorredor(auxCorredorOld) & vbCrLf & vbCrLf
    msj = msj & "Los siguientes datos fueron modificados:" & vbCrLf & vbCrLf
    for each myKey in oDiccModificaciones.Keys
        if (InStr(1, oDiccModificaciones.item(myKey), "|") > 0) then
            strDS = Split(oDiccModificaciones.item(myKey),"|")
            msj = msj & myKey & ") "& strDS(0) & " | Antes: " & strDS(1) &" | Ahora: "& strDS(2) & vbCrLf
        else
            msj = msj & myKey & ") "& oDiccModificaciones.item(myKey) & vbCrLf
        end if
    next    
	Call logMig.info("Preparacion completa. Se envia el mail a los destinatarios de la tarea " & TASK_POS_MODIFICACION_HISTORICA & ", lista: " & MAIL_TASK_INFO_LIST)
	Call SendMail(TASK_POS_MODIFICACION_HISTORICA, MAIL_TASK_INFO_LIST, "POSEIDON - MODIFICACION HISTORICA - ("&Ucase(g_pto)&")", msj, "")    
    oDiccModificaciones.RemoveAll
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function getMailsModificacionCtaPte()

    Dim rs, strSQL, ret
    strSQL= "select rel.cdtarea ,em.cdemail, em.dsemail from svcemails em " &_
            " left join SVCRELTASKLISTEMAIL rel on " &_
            " rel.cdemail = em.cdemail " &_
            " where rel.cdtarea = 2" &_
            " union " &_
            "select rl.cdtarea ,em.cdemail, em.dsemail from SVCdetaillist dl " &_
            " left join svcemails em on " &_ 
            " dl.cdemail = em.cdemail " &_
            " left join SVCRELTASKLISTEMAIL rl on " &_
            " rl.cdlist = dl.cdlist " &_
            " Where rl.cdtarea = 2 "
    Call GF_BD_Puertos(g_pto, rs, "OPEN", strSQL)
    while (not rs.eof)
        ret = ret & rs("cdemail") & ";"
        rs.MoveNext()
    wend
    getMailsModificacionCtaPte = Left(ret, Len(ret)-1)

End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function agregarHAudCamiones(pIdCamion, pDtContable)
    Dim strSQL
    strSQL = "INSERT INTO dbo.HAudCamiones "&_
             " SELECT DtContable,(SELECT case when MAX(SQAUDITORIA) is Null then 1 else MAX(SQAUDITORIA)+1 end AS SQAUDITORIA FROM HAudCamiones WHERE DTCONTABLE='"&pDtContable&"' AND IDCAMION='"& pIdCamion&"'),IDCAMION , CDCHAPACAMION, CDCHAPAACOPLADO, CDTIPOCAMION, CDTRANSPORTISTA, DSNOMBRECONDUCTOR,"&_
             " DSAPELLIDOCONDUCTOR,CDTIPODOC, NUDOCUMENTO, DTINGRESO,DTEGRESO , CDESTADO, CDCIRCUITO, CDFILA, CDSILO, CDPLATAFORMA,ICTRANSMITIDO, "&_
             " sqCamion, CDPRODUCTO, NUAUTSALIDA, SQTURNO, ICCUPO, DSCUPO, ICCONTRATOESP, IDCUPOASIGNADO, NUCUITREM, NUCUPO "&_
             " FROM dbo.HCAMIONES A WHERE IdCamion = '"& pIdCamion &"' AND DTCONTABLE = '"& pDtContable &"'"
    Call logMig.info("1 - Agrega HAudCamiones: " & strSQL)
    Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function agregarHAuditoriaCamiones(pIdCamion, pDtContable)
    Dim auxEstadoCamion,WshNetwork,auxTerminal
    auxEstadoCamion = getEstadoCamion(pIdCamion, pDtContable)
    Set WshNetwork = CreateObject("WScript.Network")
    auxTerminal = Right(WshNetwork.ComputerName,10)
    strSQL = "INSERT INTO dbo.HAuditoriaCamiones "&_
             "(DtContable,DtAuditoria,SqAuditoria,IdCamion,CdTransaccion,CdEstadoAnterior,CdEstadoPosterior,CdUserName,CdTerminal,CdSupervisor)values  "&_
             "('"& pDtContable &"','"&Year(Now())&"-"&GF_nDigits(Month(Now()),2)&"-"&GF_nDigits(Day(Now()),2)&" "&GF_nDigits(Hour(Now()),2)&":"&GF_nDigits(Minute(Now()),2)&":"&GF_nDigits(Second(Now()),2)&"',(SELECT case when MAX(SQAUDITORIA) is Null then 1 else MAX(SQAUDITORIA)+1 end AS SQAUDITORIA FROM HAuditoriaCamiones WHERE DTCONTABLE='"&pDtContable&"' AND IDCAMION='"& pIdCamion&"'),"&_
             "'"& pIdCamion &"',"& TRANSACCION_MODIFICACION_HISTORICA &","& auxEstadoCamion &","& auxEstadoCamion &",'"& session("Usuario") &"','"&auxTerminal&"','')"
    Call logMig.info("1 - Agrega HAuditoriaCamiones: " & strSQL)
    Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function agregarHAudCamionesDescarga(pIdCamion, pDtContable)
    Dim strSQL
    strSQL = "INSERT INTO dbo.HAudCamionesDescarga "&_
             " SELECT DtContable,(SELECT case when MAX(SQAUDITORIA) is Null then 1 else MAX(SQAUDITORIA)+1 end AS SQAUDITORIA FROM HAudCamionesDescarga WHERE DTCONTABLE='"&pDtContable&"' AND IDCAMION='"& pIdCamion&"') ,IdCamion,nuCartaPorte,dtCartaPOrte,CdEmpresa, CdCliente, CdCorredor,"&_
             " CdVendedor,CdCosecha,CdProcedencia,cdEntregador,vlBrutoOrigen,vlTaraOrigen,nuInfoAnalisis,nuRecibo,dtCPVencimiento,nuCtaPteDig,CTG,NuTicketPlaya "&_
             " FROM dbo.HCAMIONESDESCARGA A WHERE IdCamion = '"& pIdCamion &"' AND DTCONTABLE = '"& pDtContable &"'"
    Call logMig.info("1 - Agrega HAudCamionesDescarga: " & strSQL)
    Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function agregarRegistrosAuditoria(pIdCamion, pDtContable, pPto)
    Dim auxDtContable
    auxDtContable = Left(pDtContable,4) &"-"& Mid(pDtContable,5,2) &"-"& Right(pDtContable,2)    
    Call agregarHAudCamiones(pIdCamion, auxDtContable)
    Call agregarHAuditoriaCamiones(pIdCamion, auxDtContable)
    Call agregarHAudCamionesDescarga(pIdCamion, auxDtContable)
    'Solo actualizo la tabla de Auditoria Cuenta y Orden si tiene un Remitente o Intermediario
    if (Cdbl(auxIntermediarioCd) <> 0 and Cdbl(auxRemitenteCd) <> 0 ) then Call agregarHAudCuentayOrden(pIdCamion, auxDtContable)
    'Si actualiza la pesada Bruto agrego la auditoria de pesada
    if (Cdbl(auxPesadaBruto) <> Cdbl(auxPesadaBrutoOld)) then Call agregarAuditoriaPesada(pIdCamion,auxDtContable,auxPesadaBruto,PESADA_BRUTO)
    if (Cdbl(auxPesadaTara) <> Cdbl(auxPesadaTaraOld)) then   Call agregarAuditoriaPesada(pIdCamion,auxDtContable,auxPesadaTara,PESADA_TARA)
    if ((Cdbl(auxPesadaTara) <> Cdbl(auxPesadaTaraOld))or(Cdbl(auxPesadaBruto) <> Cdbl(auxPesadaBrutoOld))) then   Call agregarHAudMermas(pIdCamion,auxDtContable,g_strPuerto)
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function getEstadoCamion(pIdCamion, pDtContable)
    Dim strSQL,rsEst
    getEstadoCamion = 0
    strSQL = "SELECT CDESTADO FROM HCAMIONES WHERE DTCONTABLE ='"& pDtContable &"' AND IDCAMION = '"& pIdCamion &"'"
    Call GF_BD_Puertos(g_pto, rsEst, "OPEN", strSQL)
    if (not rsEst.Eof) then getEstadoCamion = rsEst("CDESTADO")
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
'Agrega a la tabla de Auditoria de cuenta y orden el nuevo resgistro del vendedor
Function agregarHAudCuentayOrden(pIdCamion, pDtContable)
    Dim strSQL ,rsCyOAud,rs
    Call GF_BD_Puertos(g_pto, rsCyOAud, "OPEN", "SELECT * FROM HCUENTAYORDENESCAMIONES WHERE IDCAMION ='"& pIdCamion &"' AND DTCONTABLE = '"& pDtContable &"'")
    while not rsCyOAud.Eof 
        strSQL = "INSERT INTO HAudCuentayOrdenesCamiones(dtcontable,sqauditoria,idcamion,sqorden,cdvendedor)"&_
                 "values('"& pDtContable &"',(SELECT case when MAX(SQAUDITORIA) is Null then 1 else MAX(SQAUDITORIA)+1 end AS SQAUDITORIA FROM HAudCuentayOrdenesCamiones WHERE DTCONTABLE='"&pDtContable&"' AND IDCAMION='"& pIdCamion&"'),'"&pIdCamion&"',"& rsCyOAud("sqorden") &","& rsCyOAud("cdvendedor") &")"
        Call logMig.info("1 - Agrega HAudCuentayOrdenesCamiones: " & strSQL)
        Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
        rsCyOAud.MoveNext()
    wend
End Function 
'-----------------------------------------------------------------------------------------------------------------------------------
'Si hay cambios en el Producto o el Destinatario o peso Bruto o Tara, Neto o Merma se actualizan las tablas de Stock Fisico, cuenta corriente
'NOTA: en el Rmd figura tambien si hay cambios en la Empresa o Silo pero estos datos 
'      no son modificables aca por lo tanto no se tienen en cuenta
Function actualizarAjusteCamion(pIdCamion, pDtContable,pCdProducto,pCdProductoOld,pCliente,pClienteOld,pMerma,pMermaOld,pPesadaBruto,pPesadaBrutoOld,pPesadaTara,pPesadaTaraOld)
    Dim auxNumerador,auxDtContable,auxSilo, myNetoOld, myNetoNew
    auxNumerador = 0
    Call logMig.info(" ************** Ingresa a la Funcion actualizarAjusteCamion( "&pIdCamion&", "&pDtContable&", "&pCdProducto&", "&pCdProductoOld&", "&pCliente&", "&pClienteOld&", "&pMerma&", "&pMermaOld&", "&pPesadaBruto&", "&pPesadaBrutoOld&", "&pPesadaTara&", "&pPesadaTaraOld&") ***************")
    if (Cdbl(pCliente) <> Cdbl(pClienteOld))or(Cdbl(pCdProducto) <> Cdbl(pCdProductoOld))or(Cdbl(pMerma) <> Cdbl(pMermaOld))or(Cdbl(pPesadaBruto) <> Cdbl(pPesadaBrutoOld))or(Cdbl(pPesadaTara) <> Cdbl(pPesadaTaraOld)) then
       'Primero debo obtener y actualizar el ultimo Contador de Numeradores de Ajuste de Camiones
        auxNumerador = getNumeradorAjstCamion()
        auxDtContable = GF_FN2DTCONTABLE(pDtContable)
        'Calculo los netos (vijeos y nuevos) para porder actualizar la ExCuentaCorriente
        myNetoOld = (Cdbl(pPesadaBrutoOld) - Cdbl(pPesadaTaraOld)) - Cdbl(pMermaOld)
        myNetoNew = (Cdbl(pPesadaBruto)- Cdbl(pPesadaTara)) - Cdbl(pMerma)
        
        'Agrega el ajuste del camion con los kilos netos y merma viejos(los que tenia antes de editar) pero con el producto y cliente nuevo
        Call logMig.info("Buscar y agregar el ajuste del camion")
        Call agregarExAjusteCamion(pIdCamion, auxDtContable,auxNumerador,pCdProducto,pCliente,myNetoOld,pMermaOld)
        'Actauliza la tabla de ExCuentaCorriente con los kilos que deja el registro y con los que tendra nuevos
        Call logMig.info("Buscar y actualizar Cuenta Corriente para debitar los kilos viejos (antes de modificar)")
        Call actualizarExCtaCorrienteCamion(pIdCamion, auxDtContable,pCdProductoOld,pClienteOld,pMermaOld,myNetoOld,"D")
        Call logMig.info("Buscar y actualizar Cuenta Corriente para acredirtar los kilos nuevos (luego de modificar)")
        Call actualizarExCtaCorrienteCamion(pIdCamion, auxDtContable,pCdProducto,pCliente,pMerma,myNetoNew,"C")

        'Agrego el movimiento del camion
        Call logMig.info("Buscar Movimientos con los datos originales (antes de modificar)")
        Call actualizarExMovimientos(pIdCamion, auxDtContable,pCdProductoOld,pClienteOld,pMermaOld,myNetoOld,"D",auxNumerador)
        Call logMig.info("Buscar Movimientos con los datos nuevos (luego de modificar)")
        Call actualizarExMovimientos(pIdCamion, auxDtContable,pCdProducto,pCliente,pMerma,myNetoNew,"C",auxNumerador)        
        'Obtengo el codigo de Silo para realizar la modificacion del Stock Fisico(debido a que este dato no es modificable en la pantalla)
        auxSilo = getCdSiloByCamion(pIdCamion, auxDtContable)
        Call logMig.info("Busco el stock fisico con los datos originales (antes de modificar)")
        Call actualizarExStockFisico(pIdCamion, auxDtContable,pCdProductoOld,pMermaOld,myNetoOld,"D",auxSilo)
        Call logMig.info("Busco el stock fisico con los datos nuevos (luego de modificar)")
        Call actualizarExStockFisico(pIdCamion, auxDtContable,pCdProducto,pMerma,myNetoNew,"C",auxSilo)
        Call logMig.info("Busco el movimiento stock fisico con los datos originales (antes de modificar)")
        Call agregarExMovimientoStockFisico(pIdCamion, auxDtContable,pCdProductoOld,myNetoOld,"D",auxSilo,auxNumerador)
        Call logMig.info("Busco el movimiento stock fisico con los datos nuevos (luego de modificar)")
        Call agregarExMovimientoStockFisico(pIdCamion, auxDtContable,pCdProducto,myNetoNew,"C",auxSilo,auxNumerador)
    end if
    Call logMig.info(" ************** Sale Funcion actualizarAjusteCamion *************** ")
    actualizarAjusteCamion = auxNumerador
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function rearmarCartaPorte(pIdCamion,pDtContable,pCdProducto,pCdProductoOld,pCliente,pClienteOld,pMerma,pMermaOld,pPesadaBruto,pPesadaBrutoOld,pPesadaTara,pPesadaTaraOld)
    Dim strSQL,auxDtContable, myNetoOld, myNetoNew
    'Primero calculo los netos que tenian antes y luego de modificar.
    myNetoOld = (Cdbl(pPesadaBrutoOld) - Cdbl(pPesadaTaraOld)) - Cdbl(pMermaOld)
    myNetoNew = (Cdbl(pPesadaBruto)- Cdbl(pPesadaTara)) - Cdbl(pMerma)
    Call logMig.info(" ************** Ingresa a la Funcion rearmarCartaPorte( "&pIdCamion&", "&pDtContable&", "&pCdProducto&", "&pCdProductoOld&", "&pCliente&", "&pClienteOld&", "&myNetoNew&", "&myNetoOld&" ) *************** ")
    'Si hay modificacion en los siguientes datos se rearma la carta de porte
    if (Cdbl(pCliente) <> Cdbl(pClienteOld))or(Cdbl(pCdProducto) <> Cdbl(pCdProductoOld))or(Cdbl(myNetoNew) <> Cdbl(myNetoOld)) then
        auxDtContable = GF_FN2DTCONTABLE(pDtContable)
        Call agregarFechasInexistentesCartaPorte(auxDtContable)
        Call agregarExSaldoInicial(auxDtContable)
    end if
    Call logMig.info(" ************** Sale Funcion rearmarCartaPorte *************** ")
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function rearmarStockFisico(pIdCamion,pDtContable,pCdProducto,pCdProductoOld,pCliente,pClienteOld,pMerma,pMermaOld,pPesadaBruto,pPesadaBrutoOld,pPesadaTara,pPesadaTaraOld)
    Dim strSQL, auxDtContable,myNetoNew,myNetoOld
    'Primero calculo los netos que tenian antes y luego de modificar, no se tiene en cuenta la MERMA a la hora aplicar la modificacion
    'Pero a la hora de la comparacion si hay cambios si se tiene en cuenta
    myNetoOld = Cdbl(pPesadaBrutoOld) - Cdbl(pPesadaTaraOld)
    myNetoNew = Cdbl(pPesadaBruto)- Cdbl(pPesadaTara)
    Call logMig.info(" ************** Ingresa a la Funcion rearmarStockFisico( "&pIdCamion&", "&pDtContable&", "&pCdProducto&", "&pCdProductoOld&", "&pCliente&", "&pClienteOld&", "&pMerma&", "&pMermaOld&", "&myNetoNew&", "&myNetoOld&" ) *************** ")
    
    if (Cdbl(pCdProducto) <> Cdbl(pCdProductoOld))or(Cdbl(myNetoNew) - CDbl(pMerma) <> Cdbl(myNetoOld) - CDbl(pMermaOld)) then
        auxDtContable = GF_FN2DTCONTABLE(pDtContable)
        Call agregarFechasInexistentesStockFisico(pIdCamion,auxDtContable,myNetoNew)
        Call agregarSaldoInicialStockFisico(auxDtContable)
    end if
    Call logMig.info(" ************** Sale Funcion rearmarStockFisico *************** ")
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function agregarSaldoInicialStockFisico(pDtContable)
    Dim strSQL,rsSF1,auxSaldoInicial,auxSilo,auxProd,auxContab,auxCredito,auxDebito,totSaldoInicial,smsg,auxContabSaldo
    Call logMig.info("Busca y actualiza el saldo inicial en el stock fisico")
    strSQL = "Select dtContable,"&_
             "      case when cdSilo is null then '' else cdSilo end as cdSilo, "&_
             "      case when cdProducto is null then 0 else cdProducto end as cdProducto,"&_
             "      case when vlSaldoInicial is null then 0 else vlSaldoInicial end as vlSaldoInicial,"&_
             "      case when vlCredito is null then 0 else vlCredito end as vlCredito,"&_
             "      case when Vldebito is null then 0 else Vldebito end as Vldebito "&_
             "from dbo.ExStockFisico "&_
             "where dtContable >= '"& pDtContable &"'"&_
             " order by cdSilo, cdProducto,dtContable "
    Call logMig.info("--> Busca Stock Fisico 3: "&strSQL)
    Call GF_BD_Puertos(g_pto, rsSF1, "OPEN", strSQL)
    If Not rsSF1.EOF Then
        Do While Not rsSF1.EOF
            auxSilo = Trim(rsSF1("CdSilo"))
            auxProd = Cdbl(rsSF1("CdProducto"))
            auxContab = Year(rsSF1("dtContable")) &"-"& GF_nDigits(Month(rsSF1("dtContable")),2) &"-"& GF_nDigits(Day(rsSF1("dtContable")),2)
            auxSaldoInicial = Cdbl(rsSF1("VlSaldoInicial"))
            auxCredito = Cdbl(rsSF1("vlCredito"))
            auxDebito = Cdbl(rsSF1("VlDebito"))
            If (Cstr(auxContab) = Cstr(pDtContable))And(Cdbl(auxSaldoInicial) <> 0) Then
                totSaldoInicial = Cdbl(auxSaldoInicial) + Cdbl(auxCredito) - Cdbl(auxDebito)
                Call logMig.info("----> Busca Stock Fisico 3-A: SON IGUALES LAS FECHAS Y SALDO MAYOR A CERO, "& auxContab &"="& pDtContable &" Y "&auxSaldoInicial&" > 0, HAGO LA CUENTA ("&auxCredito &" - "&auxDebito&") Y SE VA ACUMULANDO" )
                rsSF1.MoveNext()
            Else
                strSQL = "Select a.dtContable,"&_
                         "       case when a.vlsaldoInicial is null then 0 else a.vlsaldoInicial end as vlsaldoInicial,"&_
                         "       case when a.vlcredito is null then 0 else a.vlcredito end as vlcredito, "&_
                         "       case when a.vlDebito is null then 0 else a.vlDebito end as vlDebito "&_
                         "from dbo.ExStockFisico a "&_
                         "where dtContable < '"& pDtContable & "' and cdSilo = '"& auxSilo &"' and cdProducto = "& auxProd &_
                         " order by a.dtContable desc"   
                Call logMig.info("----> Busca Stock Fisico 4: "&strSQL)
                Call GF_BD_Puertos(g_pto, rsSF2, "OPEN", strSQL)
                If Not rsSF2.EOF Then
                    totSaldoInicial = Cdbl(rsSF2("VlSaldoInicial")) + Cdbl(rsSF2("vlCredito")) - Cdbl(rsSF2("VlDebito"))
                Else
                    totSaldoInicial = 0
                End If
            End If
            If rsSF1.EOF Then Exit Do
            Do Until rsSF1.EOF Or (Trim(rsSF1("CdSilo")) <> auxSilo) Or (Cdbl(rsSF1("CdProducto")) <> Cdbl(auxProd))
                auxContabSaldo = Year(rsSF1("dtContable")) &"-"& GF_nDigits(Month(rsSF1("dtContable")),2) &"-"& GF_nDigits(Day(rsSF1("dtContable")),2)
                strSQL = "UPDATE dbo.ExStockFisico Set vlsaldoInicial = "& totSaldoInicial &" where "&_
                         " dtContable='"& auxContabSaldo &"' and cdproducto="& auxProd &" and cdSilo='"& auxSilo&"'"
                Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
                Call logMig.info("Actualiza stock fisico: "&strSQL)
                Call logMig.info("Recalculando Saldos iniciales a partir de fecha ")
                Call logMig.info(string(10," ") & "Silo: " & auxSilo)
                Call logMig.info(string(10," ") & "Producto: " & auxProd)
                Call logMig.info(string(10," ") & "Fecha Empresa: " & auxContabSaldo)
                Call logMig.info(string(10," ") & "Saldo Inicial: " & auxSaldoInicial)
                Call logMig.info(string(10," ") & "Saldo Recalculado: " & totSaldoInicial)

                totSaldoInicial = Cdbl(totSaldoInicial) + Cdbl(rsSF1("vlCredito")) - Cdbl(rsSF1("VlDebito"))
                rsSF1.MoveNext()
                Ind = 1
                If rsSF1.EOF Then Exit Do
            Loop
            If Ind = 0 And Not rsSF1.EOF Then rsSF1.MoveNext()
        Loop
    end if

End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function agregarFechasInexistentesStockFisico(pIdCamion,pDtContable,pNeto)
    Dim strSQL,auxSilo,auxProd
    Call logMig.info("Busca y actualiza Fechas inexistentes en el stock fisico")
    strSQL = "Select Distinct Case when CdSilo is null then '' else CdSilo end as CdSilo, "&_
             "                Case when CdProducto is null then 0 else CdProducto end as CdProducto "&_
             "FROM dbo.ExStockFisico "
    Call logMig.info("--> Busca Stock Fisico 1: "&strSQL)
    Call GF_BD_Puertos(g_pto, rsSF1, "OPEN", strSQL)
    strSQL = "Select Distinct dtContable from dbo.ExStockFisico Where dtContable > '"& pDtContable &"' order by dtContable"
    Call logMig.info("--> Busca Stock Fisico 2: "&strSQL)
    Call GF_BD_Puertos(g_pto, rsSF2, "OPEN", strSQL)
    If Not rsSF1.EOF Then
        Do While Not rsSF1.EOF
            auxSilo = Trim(rsSF1("CdSilo"))
            auxProd = Cdbl(rsSF1("CdProducto"))
            If Not rsSF2.EOF Then
                Do While Not rsSF2.EOF
                    auxFecha = Year(rsSF2("dtContable")) &"-"& GF_nDigits(Month(rsSF2("dtContable")),2) &"-"& GF_nDigits(Day(rsSF2("dtContable")),2)
                    strSQL = "Select dtContable, cdProducto,cdSilo from dbo.ExStockFisico where "&_
                             " dtContable = '"& auxFecha &"' and cdSilo='"& auxSilo &"' and cdproducto=" & auxProd
                    Call logMig.info("----> Busca Stock Fisico 3: "&strSQL)
                    Call GF_BD_Puertos(g_pto, rsSF3, "OPEN", strSQL)
                    If rsSF3.EOF Then Call actualizarExStockFisico(pIdCamion,auxFecha,auxProd,0,pNeto,"C",auxSilo)
                    rsSF2.MoveNext()
                Loop
                rsSF2.MoveFirst
            End If
            rsSF1.MoveNext()
        Loop
    End If
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function agregarFechasInexistentesCartaPorte(pDtContable)
    Dim rsCC1,rsCC2,rsCC3,auxEmp,auxClie,auxProd,auxFecha
    Call logMig.info("Busca y agrega Fechas inexistentes en la cuenta corriente")
    strSQL = "Select Distinct CdEmpresa, CdCliente, CdProducto FROM dbo.ExCuentCorrientes "
    Call logMig.info("--> Busco Cuenta Corriente 1: " & strSQL)
    Call GF_BD_Puertos(g_pto, rsCC1, "OPEN", strSQL)
    strSQL = "Select Distinct dtContable from dbo.ExCuentCorrientes Where dtContable >= '"& pDtContable &"' order by dtcontable asc"
    Call logMig.info("--> Busco Cuenta Corriente 2: " & strSQL)
    Call GF_BD_Puertos(g_pto, rsCC2, "OPEN", strSQL)
    While Not rsCC1.EOF
        auxEmp = rsCC1("CdEmpresa")
        auxClie = rsCC1("CdCliente")
        auxProd = rsCC1("CdProducto")
        If Not rsCC2.EOF Then
            While Not rsCC2.EOF
                'NOTA: El campo DTCONTABLE en la DB es de tipo DATE, pero al manejarlo con ASP lo transfirma automaticamente en un STRING
                '      pero con un formato distinto al que espera para comparar contra la DB nuevamente, por este motivo se lo maneja como Fecha tambien en ASP 
                auxFecha = Year(rsCC2("dtContable")) &"-"& GF_nDigits(Month(rsCC2("dtContable")),2) &"-"& GF_nDigits(Day(rsCC2("dtContable")),2)
                strsql = "Select * from dbo.ExCuentCorrientes "&_
                         "where dtContable = '"& auxFecha &"' and cdEmpresa = " & auxEmp & " and cdCliente = " & auxClie & " and cdproducto = " & auxProd
                Call logMig.info("----> Busco Cuenta Corriente 3: " & strSQL)
                Call GF_BD_Puertos(g_pto, rsCC3, "OPEN", strSQL)
                If rsCC3.EOF Then Call agregarExCuentCorrientes(auxFecha, auxEmp, auxClie, auxProd)
                rsCC2.MoveNext()
            wend
            rsCC2.MoveFirst
        End If
        rsCC1.MoveNext()
   wend
End function
'-----------------------------------------------------------------------------------------------------------------------------------'
'ver ma�ana probar con fecha 2014-11-20 para que traiga pocos registros, y copiar el formateo de fecha par que lo tome bien
Function agregarExSaldoInicial(pDtContable)
    Dim strSQL,rsESI,auxSaldoInicial,auxEmp,auxClie,auxProd,auxContab,auxCredito,auxDebito,totSaldoInicial,auxContabSaldo
    Call logMig.info("Busca y actualiza el saldo inicial en la cuenta corriente")
    strSQL = " Select dtContable,cdEmpresa,cdCliente,cdProducto,"&_
             "        CASE WHEN vlSaldoInicial IS NULL THEN 0 ELSE vlSaldoInicial END AS vlSaldoInicial, "&_
             "        CASE WHEN vlCredito IS NULL THEN 0 ELSE vlCredito END AS vlCredito, "&_
             "        CASE WHEN Vldebito IS NULL THEN 0 ELSE Vldebito END AS Vldebito "&_
             " from dbo.ExCuentCorrientes "&_
             " where dtContable >= '"& pDtContable &"' order by cdEmpresa, cdCliente, cdProducto,dtContable "
    Call logMig.info("--> Busco Cuenta Corriente 4: " & strSQL)
    Call GF_BD_Puertos(g_pto, rsESI, "OPEN", strSQL)
    Do While Not rsESI.EOF
        auxEmp = Cdbl(rsESI("cdEmpresa"))
        auxClie = Cdbl(rsESI("cdCliente"))
        auxProd = Cdbl(rsESI("cdProducto"))
        auxContab = Year(rsESI("dtContable")) &"-"& GF_nDigits(Month(rsESI("dtContable")),2) &"-"& GF_nDigits(Day(rsESI("dtContable")),2)
        auxSaldoInicial = Cdbl(rsESI("vlSaldoInicial"))
        auxCredito = Cdbl(rsESI("vlCredito"))
        auxDebito = Cdbl(rsESI("Vldebito"))
        'Si la fecha del Camion que se esta modificando historicamente es igual a la fecha contable que esta procesando el recordset sumo los creditos y resto los debitos
        If Cstr(auxContab) = Cstr(pDtContable) Then
            Call logMig.info("----> SON IGUALES LAS FECHAS "& auxContab &"="& pDtContable &", HAGO LA CUENTA ("&auxCredito &" - "&auxDebito&") Y SE VA ACUMULANDO" )
            totSaldoInicial = Cdbl(auxSaldoInicial) + Cdbl(auxCredito) - Cdbl(auxDebito)
            rsESI.MoveNext()
        Else
            strSQL = "Select a.dtcontable,"&_
                     "       CASE WHEN a.vlsaldoInicial IS NULL THEN 0 ELSE a.vlsaldoInicial END AS vlsaldoInicial,"&_
                     "       CASE WHEN a.vlcredito IS NULL THEN 0 ELSE a.vlcredito END AS vlcredito,"&_
                     "       CASE WHEN a.vlDebito IS NULL THEN 0 ELSE a.vlDebito END AS vlDebito "&_
                     "from dbo.ExCuentCorrientes a "&_
                     "where a.dtContable < '"& pDtContable &"'"&_
                     " and a.cdEmpresa="& auxEmp &" and a.cdcliente="& auxClie &" and a.cdProducto = " & auxProd &_
                     " order by dtContable desc "
            Call logMig.info("----> Busco Cuenta Corriente 5: " & strSQL)
            Call GF_BD_Puertos(g_pto, rsESI1, "OPEN", strSQL)
            If Not rsESI1.Eof Then
                totSaldoInicial = Cdbl(rsESI1("VlSaldoInicial")) + Cdbl(rsESI1("vlCredito")) - Cdbl(rsESI1("VlDebito"))
            Else
                totSaldoInicial = 0
            End If
        End if
        Ind = 0
        If rsESI.EOF Then Exit Do
        Do Until rsESI.EOF Or (Cdbl(rsESI("cdEmpresa")) <> Cdbl(auxEmp)) Or (Cdbl(rsESI("CDCliente")) <> Cdbl(auxClie)) Or (Cdbl(rsESI("CdProducto")) <> Cdbl(auxProd))
            auxContabSaldo = Year(rsESI("dtContable")) &"-"& GF_nDigits(Month(rsESI("dtContable")),2) &"-"& GF_nDigits(Day(rsESI("dtContable")),2)
            strSQL = "UPDATE dbo.ExCuentCorrientes Set vlsaldoInicial = "&totSaldoInicial&" where "&_
                     " dtContable = '"& auxContabSaldo &"' and cdproducto = "&auxProd&" and cdEmpresa = "&auxEmp&" and cdCliente = "&auxClie
            Call logMig.info("----> Actualizo cuenta corriente: "& strSQL)
            Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
            totSaldoInicial = Cdbl(totSaldoInicial) + Cdbl(rsESI("vlCredito")) - Cdbl(rsESI("VlDebito"))
            rsESI.MoveNext()
            Ind = 1
            If rsESI.EOF Then Exit Do
        Loop
        If Ind = 0 And Not rsESI.EOF Then rsESI.MoveNext()
    Loop
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function agregarExCuentCorrientes(pDtContable,pEmpresa,pCliente,pProducto)
    Dim strSQL
    strSQL =  "Insert into dbo.ExCuentCorrientes values('"&pDtContable&"',"&pEmpresa&","&pCliente&","&pProducto&",0,0,0)"
    Call logMig.info("----> Agrega Cuenta Corriente: " & strSQL)
    Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function agregarExMovimientoStockFisico(pIdCamion, pDtContable,pCdProducto,pNeto,pTransaccion,pSilo,pNumerador)
    Dim strSQL,auxTransaccion,auxDtMovimiento
    auxDtMovimiento = session("MmtoSistema")
    if (pTransaccion = "D") then
        auxTransaccion = 21
    else
        auxTransaccion = 121
    end if
    auxDtMovimiento = Left(auxDtMovimiento,4) &"-"& Mid(auxDtMovimiento,5,2)  &"-"& Mid(auxDtMovimiento,7,2) &" "& Mid(auxDtMovimiento,9,2) &":"& Mid(auxDtMovimiento,11,2) &":"& Mid(auxDtMovimiento,13,2)
    
    strSQL = "INSERT into dbo.ExMovStockFisico Values('"&auxDtMovimiento&"','"&pDtContable&"','"&pSilo&"',"&pCdProducto&","&auxTransaccion&","&pNeto&",'"&pNumerador&"')"
    Call logMig.info("--> Agrego movimiento stock fisico: "&strSQL)
    Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function getCdSiloByCamion(pIdCamion, pDtContable)
    Dim strSQL,rsSilo
    getCdSiloByCamion = ""
    strSQL = "SELECT CASE WHEN CDSILO IS NULL THEN '' ELSE CDSILO END AS CDSILO FROM HCAMIONES WHERE IDCAMION = '"&pIdCamion&"' AND DTCONTABLE='"&pDtContable&"'"
    Call GF_BD_Puertos(g_pto, rsSilo, "OPEN", strSQL)
    if (not rsSilo.Eof) then getCdSiloByCamion = Trim(rsSilo("CDSILO"))
    Call logMig.info("Busco el Silo de la Cta.Porte: " & strSQL)
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function actualizarExStockFisico(pIdCamion, pDtContable,pCdProducto,pMerma,pNeto,pTransaccion,pSilo)
    Dim strSQL,rsSto
    strSQL = "SELECT * FROM dbo.ExStockFisico where cdProducto="&pCdProducto&" and cdSilo='"&pSilo&"' and dtContable='"&pDtContable&"'"
    Call logMig.info("----> Obtengo el stock Fisico: " & strSQL)
    Call GF_BD_Puertos(g_pto, rsSto, "OPEN", strSQL)
    if (not rsSto.Eof) then 
        strSQL = "UPDATE dbo.ExStockFisico SET "
        if (pTransaccion = "D") then
            strSQL = strSQL & " vlDebito = vlDebito + "& Cdbl(pNeto) 
        else
            strSQL = strSQL & " vlCredito = vlCredito + "& Cdbl(pNeto)  
        end if
        strSQL = strSQL & " where cdProducto="&pCdProducto&" and cdSilo='"&pSilo&"' and dtContable ='"&pDtContable&"'"
        Call logMig.info("----> Actualizo el stock fisico: " & strSQL)
    else
        strSQL = "Insert into dbo.ExStockFisico values('"&pDtContable&"','"&pSilo&"',"&pCdProducto&",0,"
        if (pTransaccion = "D") then
            strSQL = strSQL & "0,"&pNeto &")"
        else
            strSQL = strSQL & pNeto & ",0)"
        end if
        Call logMig.info("----> Agrego el stock fisico: " & strSQL)
    end if
    Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function actualizarExMovimientos(pIdCamion, pDtContable,pCdProducto,pCliente,pMerma,pNeto,pTransaccion,pAjuste)
    Dim strSQL,auxDtMovimiento, auxTransaccion
    auxDtMovimiento = session("MmtoSistema")
    if (pTransaccion = "D") then
        auxTransaccion = 121
    else
        auxTransaccion = 21
    end if
    auxDtMovimiento = Left(auxDtMovimiento,4) &"-"& Mid(auxDtMovimiento,5,2)  &"-"& Mid(auxDtMovimiento,7,2) &" "& Mid(auxDtMovimiento,9,2) &":"& Mid(auxDtMovimiento,11,2) &":"& Mid(auxDtMovimiento,13,2)
    strSQL = "INSERT into dbo.ExMovimientos Values('"&auxDtMovimiento&"','"&pDtContable&"',1,"&pCliente&","&pCdProducto&","&auxTransaccion&","&pNeto&",'"&pAjuste&"')"
    Call logMig.info("--> Agrego los movimientos: " & strSQL)
    Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
End function
'-----------------------------------------------------------------------------------------------------------------------------------
Function actualizarExCtaCorrienteCamion(pIdCamion, pDtContable,pCdProducto,pCliente,pMerma,pNeto,pTransaccion)
    Dim strSQL,rsCta
    strSQL = "SELECT * FROM ExCuentCorrientes WHERE cdProducto="&pCdProducto&" and cdEmpresa=1 and cdCliente="&pCliente&" and dtContable='"&pDtContable&"'"
    Call logMig.info("--> Obtengo los movientos del CTA.CTE : " & strSQL)
    Call GF_BD_Puertos(g_pto, rsCta, "OPEN", strSQL)
    if (not rsCta.Eof) then        
        strSQL = "UPDATE ExCuentCorrientes SET "
        if (pTransaccion = "D") then
            strSQL = strSQL & "vlDebito = vlDebito + " & pNeto
        else
            strSQL = strSQL & "vlCredito = vlCredito + " & pNeto
        end if
        strSQL = strSQL & " WHERE cdProducto="&pCdProducto&" and cdEmpresa=1 and cdCliente="&pCliente&" and dtContable='"&pDtContable&"'"
        Call logMig.info("--> Actualizo la CTA.CTE: " & strSQL)
    else
        strSQL = "INSERT INTO ExCuentCorrientes VALUES('"&pDtContable&"',1,"&pCliente&","&pCdProducto&",0,"
        if (pTransaccion = "D") then
            strSQL = strSQL &" 0,"& pNeto & ")"
        else
            strSQL = strSQL & pNeto & ",0)"             
        end if
        Call logMig.info("-->Agrego la CTA.CTE: " & strSQL)
    end if
    Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function agregarExAjusteCamion(pIdCamion, pDtContable,pNumerador,pCdProducto,pCliente,pNeto,pMermaKl)
    Dim strSQL, auxDtAjuste
    auxDtAjuste = Year(Now())&"-"&GF_nDigits(Month(Now()),2)&"-"&GF_nDigits(Day(Now()),2)&" "&GF_nDigits(Hour(Now()),2)&":"&GF_nDigits(Minute(Now()),2)&":"&GF_nDigits(Second(Now()),2)
    strSQL = "INSERT INTO ExhCamionAjustes VALUES('"& pDtContable &"','"&pIdCamion&"',"&pNumerador&",0,'"&auxDtAjuste&"',1,"&pCliente&","&pCdProducto&","&pNeto&","&pMermaKl&")"
    Call logMig.info("--> Agrego el ajuste del camion: " & strSQL)
    Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function getKilosMermaCamion(pIdCamion, pDtContable,pCdProducto,pCliente)
    Dim strSQL,rsMerma
    getKilosMermaCamion = 0
    strSQL = "SELECT CASE WHEN VLMERMA IS NULL THEN 0 ELSE VLMERMA END AS VLMERMA FROM ExhCamionAjustes WHERE DTCONTABLE='"&pDtContable&"' AND IDCAMION='"&pIdCamion&"' AND CDPRODUCTO="&pCdProducto&" AND CDCLIENTE="&pCliente&" AND CDEMPRESA=1"
    Call GF_BD_Puertos(g_pto, rsMerma, "OPEN", strSQL)
    if not rsMerma.Eof then
        getKilosMermaCamion = rsMerma("VLMERMA")
    end if
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function getKilosPesadaCamion(pIdCamion, pDtContable,pTipoPesada,p_strPuerto)
    Dim strSQL,rsPesada,myDtContable
    getKilosPesadaCamion = 0
    myDtContable = GF_FN2DTCONTABLE(pDtContable)
    strSQL = "SELECT MAX(SQPESADA),CASE WHEN VLPESADA IS NULL THEN 0 ELSE VLPESADA END AS VLPESADA FROM HPESADASCAMION WHERE DTCONTABLE ='"& myDtContable &"' AND IDCAMION = '"& pIdCamion &"' AND CDPESADA = "& pTipoPesada &" GROUP BY VLPESADA "
    Call GF_BD_Puertos(p_strPuerto, rsPesada, "OPEN", strSQL)
    if not rsPesada.Eof then getKilosPesadaCamion = rsPesada("VLPESADA")
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
'Actualiza la cantidad de cupos asignados
Function actualizarCupoAsignado(p_NroCupo, p_Operator)
    Dim strSQL,rsCupo
    strSQL = "UPDATE DBO.ASIGNACIONCUPOS SET QTINGRESADOS = QTINGRESADOS " & p_Operator & " 1 WHERE IDCUPO= " & p_NroCupo
    Call GF_BD_Puertos(g_pto, rsCupo, "EXEC", strSQL)
End Function 
'-----------------------------------------------------------------------------------------------------------------------------------
Function getNumeradorAjstCamion()
    Dim strSQL, numerador,rsNum
    strSQL = "Update dbo.ContadoresNumeradores Set VlUltimoAsignado = VlUltimoAsignado + 1,dtUltimoValor='"&  Year(Now())&"-"&GF_nDigits(Month(Now()),2)&"-"&GF_nDigits(Day(Now()),2) &"' where CDNUMERADOR='ULTAJCAMION'"
    Call GF_BD_Puertos(g_pto, rs, "EXEC", strSQL)
    Call logMig.info("Actualizo el numero asignado de ajuste: " & strSQL)
    Call GF_BD_Puertos(g_pto, rsNum, "OPEN", "Select VlUltimoAsignado from dbo.ContadoresNumeradores where CDNUMERADOR='ULTAJCAMION'")
    numerador = 0
    if not rsNum.Eof then numerador = rsNum("VlUltimoAsignado")
    Call logMig.info("Obtengo el ultimo numero asignado de ajuste: " & numerador)
    getNumeradorAjstCamion = numerador
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function inicializarLogCartaPorte()
    Set logMig = new classLog
    Call startLog(HND_FILE,MSG_INF_LOG+MSG_ERR_LOG+MSG_WRN_LOG)
    logMig.fileName = "MODIFICACION_HISTORICA_"& left(session("MmtoDato"),8) & "_" & g_pto
    Call logMig.info("-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* INICIA LOG -*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*")
    Call logMig.info("-- CARTA DE PORTE: "& g_ctaPte )
    Call logMig.info("-- CAMION: "& g_IdCamion )
    Call logMig.info("-- FECHA CONTABLE: "& GF_FN2DTCONTABLE(g_dtContable) )
    Call logMig.info("-- PUERTO: "& g_pto )
End Function 
'----------------------------------------------------------------------------------------------------------
'Funcion que calcula la merma para un determinado producto y rubro dependiendo del valor reabaja, solo aplica a rubros de Humedad y Zaranda
Function CalcularMerma(pCdProducto, pCdRubro, pValor, p_Pto)
    Dim strSQL, rs,listaRubrosHumedad,listaRubrosZaranda,myMerma
    myMerma = 0
    listaRubrosHumedad = "," & getValueParametro(PARAM_CD_RUBRO_HUMEDAD,p_Pto) & ","
    listaRubrosZaranda = "," & getValueParametro(PARAM_CD_RUBRO_ZARANDA,p_Pto) & ","
    
    if (InStr(1, listaRubrosHumedad, "," & pCdRubro & ",") > 0) then
        strSQL= "Select (VLMERMAXTABLA + MERMAXMANIPULEO) PORCMERMA from " &_                
                " MERMAXSECADO MXS " &_
                " INNER JOIN GASTOSXSECADO GXS ON GXS.CDPRODUCTO=MXS.CDPRODUCTO" &_
                " where MXS.CDPRODUCTO=" & pCdProducto & " and MXS.VLHUMEDAD=" & pValor
        Call GF_BD_Puertos(p_Pto, rs, "OPEN", strSQL) 
        if (not rs.eof) then myMerma = CDbl(rs("PORCMERMA"))
    end if    
    if (InStr(1, listaRubrosZaranda, "," & pCdRubro & ",") > 0) then
        strSQL="Select * from MERMASAUTOMATICASPENALIZACION where CDPRODUCTO=" & pCdProducto & " and CDRUBRO=" & pCdRubro & " and VALORMINIMO<=" & pValor & " and VALORMAXIMO>=" & pValor
        Call GF_BD_Puertos(p_Pto, rs, "OPEN", strSQL)
        if (not rs.eof) then myMerma = CDbl(rs("MERMAVARIABLE"))
    end if
    CalcularMerma = myMerma
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Dim g_ctaPte, g_dtContable, strSQL, g_rs, conn,g_pto, auxCartaPorte1,auxCartaPorte2,auxCTG,auxDtCartaPorte,auxDtVencimiento
Dim auxTitularDs,auxTitularCuit1,auxTitularCuit2,auxTitularCuit3,auxIntermediarioDs,auxIntermediarioCuit1,auxIntermediarioCuit2,auxIntermediarioCuit3,auxRemitenteDs,auxTransportista,auxChoferTipoDoc,auxChoferNumDoc,auxChoferDs
Dim auxCorredor,auxCorredorCuit1,auxCorredorCuit2,auxCorredorCuit3,auxCorredorDs,auxEntregador,auxEntregadorCuit1,auxEntregadorCuit2,auxEntregadorCuit3,auxEntregadorDs,auxDestinatario,auxDestinatarioCuit1,auxDestinatarioCuit2,auxDestinatarioCuit3,auxDestinatarioDs,auxRemitenteCuit1,auxRemitenteCuit2,auxRemitenteCuit3
Dim auxTransportistaDs,auxTransportistaCuit1,auxTransportistaCuit2,auxTransportistaCuit3,auxTitularCd,auxIntermediarioCd,auxRemitenteCd,auxProcedenciaCd,auxProcedenciaDs
Dim auxChoferNumDoc1,auxChoferNumDoc2,auxChoferNumDoc3,auxProcedenciaProv,auxCosecha,auxGrano,auxPesoBruto,auxCupo,auxBiotecnologia,auxObservaciones,auxPesoTara,auxPesoNeto
Dim auxAcoplado,auxChapa,auxFechaArribo,auxHoraArribo,auxFechaEgreso,auxHoraEgreso,auxTurno,auxPesadaBruto,auxPesadaTara,auxObservacionesDescarga
Dim errorCabecera,errorInterviniente,errorProducto,g_Error,errorTransporte,errorDescarga,oDiccModificaciones,g_IdCamion,g_ctaPteOld,g_Merma
Dim auxDtCartaPorteOld,auxDtVencimientoOld,auxCTGOld,auxCosechaOld,auxGranoOld,auxPesoBrutoOld,auxCupoOld,auxPesoTaraOld,auxPesadaBrutoOld,auxObservacionesDescargaOld,auxFechaEgresoOld,auxProcedenciaProvOld
Dim auxBiotecnologiaOld,auxPesoNetoOld,auxObservacionesOld,auxProcedenciaCdOld,auxChapaOld,auxAcopladoOld,auxTurnoOld,auxChoferTipoDocOld
Dim auxTitularCdOld,auxTitularCuitOld,auxIntermediarioCdOld,auxIntermediarioCuitOld,auxRemitenteCdOld,auxRemitenteCuitOld,auxCorredorOld
Dim auxCorredorCuitOld,auxEntregadorOld,auxEntregadorCuitOld,auxDestinatarioOld,auxDestinatarioCuitOld,auxTransportistaOld,auxTransportistaCuitOld,auxChoferCuitOld,logMig,auxMerma,auxMermaPorcentaje
Dim auxMermaPorcentajeOld,auxPesadaTaraOld,auxMermaOld,auxNetoCMerma,auxNetoSMerma,auxNroAjuste


accion = GF_PARAMETROS7("accion", "" ,6)
g_ctaPte = GF_PARAMETROS7("cartaPorte", "" ,6)
g_idCamion = GF_PARAMETROS7("idCamion", "" ,6)
call addParam("cartaPorte", g_ctaPte, params)
g_dtContable = GF_PARAMETROS7("dtContable", "" ,6)
call addParam("dtContable", g_dtContable, params)
g_pto = GF_PARAMETROS7("pto", "" ,6)
g_strPuerto = g_pto
call addParam("pto", g_pto, params)

flagGrabo = false
Set oDiccModificaciones = createObject("Scripting.Dictionary")

%>
