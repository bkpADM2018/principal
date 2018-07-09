<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientos.asp"-->
<!--#include file="Includes/procedimientosPuertos.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosLog.asp"-->
<!--#include file="Includes/procedimientosMail.asp"-->
<%

'-----------------------------------------------------------------------------------------------------------

Function CalcularMerma(pCdProducto, pCdRubro, pValor, ByRef pMerma, ByRef pEsZaranda)
    Dim strSQL, rs
    
    pMerma = 0
    pEsZaranda = false
    if (InStr(1, g_listaRubrosHumedad, "," & pCdRubro & ",") > 0) then
        strSQL= "Select (VLMERMAXTABLA + MERMAXMANIPULEO) PORCMERMA from " &_                
                " MERMAXSECADO MXS " &_
                " INNER JOIN GASTOSXSECADO GXS ON GXS.CDPRODUCTO=MXS.CDPRODUCTO" &_
                " where MXS.CDPRODUCTO=" & pCdProducto & " and MXS.VLHUMEDAD=" & pValor                
        Call GF_BD_Puertos(g_strPuerto, rs, "OPEN", strSQL) 
        if (not rs.eof) then pMerma = CDbl(rs("PORCMERMA"))
    end if    
    if (InStr(1, g_listaRubrosZaranda, "," & pCdRubro & ",") > 0) then
        pEsZaranda = true
        strSQL="Select * from MERMASAUTOMATICASPENALIZACION where CDPRODUCTO=" & pCdProducto & " and CDRUBRO=" & pCdRubro & " and VALORMINIMO<=" & pValor & " and VALORMAXIMO>=" & pValor                
        Call GF_BD_Puertos(g_strPuerto, rs, "OPEN", strSQL) 
        if (not rs.eof) then pMerma = CDbl(rs("MERMAVARIABLE"))
    end if     
End Function
'-----------------------------------------------------------------------------------------------------------
Function corteCamion(pRs, pDtContable, pIdCamion)
    Dim ret 
    
    ret = true
    if (not pRs.eof) then
        if ((pRs("DTCONTABLE") = pDtContable) and (pRs("IDCAMION") = pIdCamion)) then ret = false
    end if    
    corteCamion = ret
End Function
'-----------------------------------------------------------------------------------------------------------
Function corteVagon(pRs, pDtContable, pCdOperativo, pCdVagon)
    Dim ret 
    
    ret = true
    if (not pRs.eof) then
        if ((pRs("DTCONTABLE") = pDtContable) and (pRs("CDOPERATIVO") = pCdOperativo) and (pRs("CDVAGON") = pCdVagon)) then ret = false
    end if    
    corteVagon = ret
End Function
'-----------------------------------------------------------------------------------------------------------
'                                               CAMIONES
'-----------------------------------------------------------------------------------------------------------
Function procesarCamiones(pPto, pDtContable)

    Dim strSQL, rs1, myTotalMermaDetalleCalc, myMerma, myTotalMermaDetalleTabla, strLog, vecRubros(100, 5), cantOK, cantError
    Dim ret, esZaranda

    '--Ppia Produccion
    strSQL="Select  HCD.NUCARTAPORTE,   " &_
           "        HCD.CDCORREDOR,     " &_ 
           "        HCD.CDVENDEDOR,     " &_
           "        'V' PpiaProd,    " &_
           "        DO.NCESTABLEPROCE,    " &_
           "        DO.CRAZONDESTINATARIO,           " &_
           "        HC.NUCUITREM, " &_         
           "        V.NUDOCUMENTO " &_
           "        from HCAMIONESDESCARGA HCD  " &_
           "        inner join HCAMIONES  HC on HC.DTCONTABLE=HCD.DTCONTABLE and HC.IDCAMION=HCD.IDCAMION   " &_
           "        inner join DATOSONCCA DO on HCD.NUCARTAPORTE=DO.NCCARTAPORTE    " &_
           "        inner join VENDEDORES V on V.CDVENDEDOR=HCD.CDVENDEDOR  " &_
           "        where HCD.DTCONTABLE='" & pDtContable & "'   " &_
           "        and DO.NCESTABLEPROCE = '1'     " &_
           "        and HCD.CDCLIENTE=1 " &_
           "        and HC.NUCUITREM = V.NUDOCUMENTO    " &_
           " Union  "
    '--No Ppia Produccion
    strSQL= strSQL & "Select  HCD.NUCARTAPORTE,   " &_
           "        HCD.CDCORREDOR,     " &_
           "        HCD.CDVENDEDOR,     " &_
           "        'F' PpiaProd,    " &_
           "        DO.NCESTABLEPROCE,   " &_
           "        DO.CRAZONDESTINATARIO,   " &_
           "        HC.NUCUITREM, " &_
           "        V.NUDOCUMENTO   " &_
           "        from HCAMIONESDESCARGA HCD  " &_
           "        inner join HCAMIONES  HC on HC.DTCONTABLE=HCD.DTCONTABLE and HC.IDCAMION=HCD.IDCAMION   " &_
           "        inner join DATOSONCCA DO on HCD.NUCARTAPORTE=DO.NCCARTAPORTE    " &_
           "        inner join VENDEDORES V on V.CDVENDEDOR=HCD.CDVENDEDOR  " &_
           "        where HCD.DTCONTABLE='" & pDtContable & "'   " &_
           "        and (DO.NCESTABLEPROCE <> '1'   " &_
	       "            or HCD.CDCLIENTE<>1 " &_
	       "            or HC.NUCUITREM <> V.NUDOCUMENTO)"
    Call GF_BD_Puertos(pPto, rs, "OPEN", strSQL)
    while (not rs.eof)
%>        
    <tr>
        <td><% =rs("NUCARTAPORTE") %></td>
        <td><% =rs("CDCORREDOR") %></td>
        <td><% =rs("CDVENDEDOR") %></td>
        <td><% =rs("PpiaProd") %></td>
        <td><% =rs("NCESTABLEPROCE") %></td>
        <td><% =rs("CRAZONDESTINATARIO") %></td>
        <td><% =rs("NUCUITREM") %></td>
        <td><% =rs("NUDOCUMENTO") %></td>
    </tr>
<%                      
        rs.MoveNext()
    wend
    
End Function
'-----------------------------------------------------------------------------------------------------------
'                                               VAGONES
'-----------------------------------------------------------------------------------------------------------
Function procesarVagones(pDtContable, pAction)
    
    Dim strSQL, rs1, myTotalMermaDetalleCalc, myMerma, myTotalMermaDetalleTabla, strLog, vecRubros(100, 5), cantOK, cantError
    Dim ret, esZaranda
    
    logMig.info(GF_nChars("",180, "-", CHR_AFT))
    logMig.info("***    CONTROL MERMAS    ***")
    logMig.info("Puerto:" & g_strPuerto)
    logMig.info("Fecha :" & pDtContable)
    logMig.info("Transporte: VAGONES        ")
    logMig.info(GF_nChars("",180, "-", CHR_AFT))
    
    strSQL=" Select HC.DTCONTABLE, HC.CDOPERATIVO, HC.CDVAGON, NUCARTAPORTE CARTAPORTE, MERMATOTAL, SQCALADA, B.CDRUBRO, DSRUBRO, HC.CDPRODUCTO, VLBONREBAJA," &_
        "ROUND(( (SELECT PC.vlpesada  " &_
        "			   FROM   db2admin.hpesadasvagon PC  " &_
        "			   WHERE  PC.dtcontable = HC.dtcontable  " &_
        "			          AND PC.CDOPERATIVO = HC.CDOPERATIVO  " &_
        "			          AND PC.CDVAGON = HC.CDVAGON  " &_
        "			          AND PC.cdpesada = 1  " &_
        "			          AND PC.sqpesada = (SELECT Max(sqpesada)  " &_
        "			                             FROM   db2admin.hpesadasvagon  " &_
        "			                             WHERE  PC.dtcontable = dtcontable  " &_
        "			                                    AND PC.CDOPERATIVO = CDOPERATIVO  " &_
        "			                                    AND PC.CDVAGON = CDVAGON  " &_
        "			                                    AND cdpesada = 1)) -  " &_
        "			    (SELECT PC.vlpesada  " &_
        "				    FROM   db2admin.hpesadasvagon PC  " &_
        "			     WHERE  PC.dtcontable = HC.dtcontable  " &_
        "			            AND PC.CDOPERATIVO = HC.CDOPERATIVO  " &_
        "			            AND PC.CDVAGON = HC.CDVAGON  " &_
        "			            AND PC.cdpesada = 2  " &_
        "			            AND PC.sqpesada = (SELECT Max(sqpesada)  " &_
        "			                               FROM   db2admin.hpesadasvagon  " &_
        "			                               WHERE  PC.dtcontable = dtcontable  " &_
        "			                                      AND PC.CDOPERATIVO = CDOPERATIVO  " &_
        "			                                      AND PC.CDVAGON = CDVAGON  " &_
        "			                                      AND cdpesada = 2)) ) * MERMACABECERA/100, 0) " &_
        "			KGMERMACALCULADA, MERMACABECERA, MERMADETALLE " &_
        "from " &_
        "(Select * from HVAGONES where CDESTADO in (6, 8) and DTCONTABLE='" & DtContable & "') HC  " &_
        "inner join " &_
        "(Select DTCONTABLE, CDOPERATIVO, CDVAGON, SUM(VLMERMAKILOS) MERMATOTAL from HMERMASVAGONES A where SQPESADA = (Select MAX(SQPESADA) from HMERMASVAGONES where DTCONTABLE=A.DTCONTABLE and CDOPERATIVO=A.CDOPERATIVO and CDVAGON=A.CDVAGON) group by DTCONTABLE, CDOPERATIVO, CDVAGON) MC on MC.DTCONTABLE=HC.DTCONTABLE and MC.CDOPERATIVO=HC.CDOPERATIVO and MC.CDVAGON=HC.CDVAGON " &_
        "inner join " &_
        "(Select DTCONTABLE, CDOPERATIVO, CDVAGON, SQCALADA, CDRUBRO, VLBONREBAJA, VLMERMA MERMADETALLE from HRUBROSVISTEOVAGONES A where SQCALADA = (Select MAX(SQCALADA) from HRUBROSVISTEOVAGONES where  DTCONTABLE=A.DTCONTABLE and CDOPERATIVO=A.CDOPERATIVO and CDVAGON=A.CDVAGON)) B  " &_
        "on B.DTCONTABLE=HC.DTCONTABLE and B.CDOPERATIVO=HC.CDOPERATIVO and B.CDVAGON=HC.CDVAGON " &_
        "inner join  " &_
        "(Select DTCONTABLE, CDOPERATIVO, CDVAGON, sum(PCMERMA) MERMACABECERA from HCALADADEVAGONES A where SQCALADA = (Select MAX(SQCALADA) from HCALADADEVAGONES where DTCONTABLE=A.DTCONTABLE and CDOPERATIVO=A.CDOPERATIVO and CDVAGON=A.CDVAGON) group by DTCONTABLE, CDOPERATIVO, CDVAGON) C on B.DTCONTABLE=C.DTCONTABLE and B.CDOPERATIVO=C.CDOPERATIVO and B.CDVAGON=C.CDVAGON " &_
        "inner join RUBROS R on R.CDRUBRO=B.CDRUBRO " &_
        "ORDER BY HC.DTCONTABLE, HC.CDOPERATIVO, HC.CDVAGON"
    'response.write strSQL        
    'response.end
    Call GF_BD_Puertos(g_strPuerto, rs1, "OPEN", strSQL)                
    strLog = GF_nChars("DTCONTABLE", 10, " ", CHR_AFT)
    strLog = strLog & " | " & GF_nChars("OPERATIVO",       12, " ", CHR_AFT)
    strLog = strLog & " | " & GF_nChars("VAGON",           10, " ", CHR_AFT)
    strLog = strLog & " | " & GF_nChars("CTAPORTE",        12, " ", CHR_AFT)
    strLog = strLog & " | " & GF_nChars("PRODUCTO",         8, " ", CHR_AFT)
    strLog = strLog & " | " & GF_nChars("MERMA KG",         8, " ", CHR_AFT)
    strLog = strLog & " | " & GF_nChars("MERMA KG CALC.",  14, " ", CHR_AFT)
    strLog = strLog & " | " & GF_nChars("VISTEO CABECERA", 15, " ", CHR_AFT)
    strLog = strLog & " | " & GF_nChars("VISTEO DETALLE",  43, " ", CHR_AFT)
    strLog = strLog & " | SOLUCION"          
    logMig.info(strLog)
    strLog = GF_nChars("", 10, " ", CHR_AFT)        
    strLog = strLog & " | " & GF_nChars("", 12, " ", CHR_AFT)
    strLog = strLog & " | " & GF_nChars("", 10, " ", CHR_AFT)
    strLog = strLog & " | " & GF_nChars("", 12, " ", CHR_AFT)
    strLog = strLog & " | " & GF_nChars("",  8, " ", CHR_AFT)
    strLog = strLog & " | " & GF_nChars("",  8, " ", CHR_AFT)
    strLog = strLog & " | " & GF_nChars("", 14, " ", CHR_AFT)
    strLog = strLog & " | " & GF_nChars("", 15, " ", CHR_AFT)
    strLog = strLog & " | " & GF_nChars("RUBRO"     , 20, " ", CHR_AFT)
    strLog = strLog & " | " & GF_nChars("EN TABLA"  ,  8, " ", CHR_AFT)            
    strLog = strLog & " | " & GF_nChars("CALCULADO" ,  9, " ", CHR_AFT) & " |"        
    logMig.info(strLog)        
    logMig.info(GF_nChars("",180, "-", CHR_AFT))
    cantTotal = 0
    cantError = 0
    while (not rs1.eof)
        dtContableOld = rs1("DTCONTABLE")
        cdOperativoOld = rs1("CDOPERATIVO")
        cdVagonOld = rs1("CDVAGON")
        cartaPorteOld = rs1("CARTAPORTE")
        productoOld = rs1("CDPRODUCTO")
        mermaTotalOld = rs1("MERMATOTAL")
        mermaCalculadaOld = rs1("KGMERMACALCULADA")
        mermaCabeceraOld = rs1("MERMACABECERA")    
        sqCaladaOld = rs1("SQCALADA")        
        'Calculo la merma rubro x rubro.
        idx=0
        myTotalMermaDetalleCalc = 0
        myTotalMermaDetalleTabla = 0
        while (not corteVagon(rs1, dtContableOld, cdOperativoOld, cdVagonOld))        
            Call CalcularMerma(rs1("CDPRODUCTO"), rs1("CDRUBRO"), rs1("VLBONREBAJA"), myMerma, esZaranda)
            myTotalMermaDetalleCalc = myTotalMermaDetalleCalc + myMerma
            myTotalMermaDetalleTabla = myTotalMermaDetalleTabla + CDbl(rs1("MERMADETALLE"))            
            idx=idx+1        
            vecRubros(idx,1) = GF_nChars(rs1("CDRUBRO"), 4, "0", CHR_FWD) & "-" & left(rs1("DSRUBRO"), 15)
            vecRubros(idx,2) = myMerma
            vecRubros(idx,3) = rs1("MERMADETALLE")
            vecRubros(idx,4) = true
            vecRubros(idx,5) = rs1("CDRUBRO")   
            'Si el rubro tiene grabada una merma que no coincide con la calculada, se almacena para mostrar el error.
            if (esZaranda) then
                'Si es Zaranda se toma como valido lo del puerto, solo se controla que se haya echo una merma en el puerto.
                if (((myMerma > 0) and (CDbl(rs1("MERMADETALLE")) = 0)) or _
                   ((myMerma = 0) and (CDbl(rs1("MERMADETALLE")) > 0))) then vecRubros(idx,4) = false
            else
                if (round(CDbl(myMerma), 2) <> round(CDbl(rs1("MERMADETALLE")), 2)) then vecRubros(idx,4) = false                        
            end if
            rs1.MoveNext()
        wend                 
        'Imprimo los datos del camion indicando los errores encontrados.
        strLog = GF_nChars(GF_STANDARIZAR_FECHA_RTRN(dtContableOld), 10, " ", CHR_FWD)
        strLog = strLog & " | " & GF_nChars(cdOperativoOld,          12, " ", CHR_FWD)        
        strLog = strLog & " | " & GF_nChars(cdVagonOld,              10, " ", CHR_FWD)
        strLog = strLog & " | " & GF_nChars(cartaPorteOld,           12, " ", CHR_FWD)
        strLog = strLog & " | " & GF_nChars(productoOld,              8, " ", CHR_FWD)
        strLog = strLog & " | " & GF_nChars(mermaTotalOld,            8, " ", CHR_FWD)
        strLog = strLog & " | " & GF_nChars(mermaCalculadaOld,       14, " ", CHR_FWD)
        strLog = strLog & " | " & GF_nChars(mermaCabeceraOld & "%",  15, " ", CHR_FWD)
        strLog = strLog & " | " & GF_nChars(vecRubros(1,1), 20, " ", CHR_AFT)
        strLog = strLog & " | " & GF_nChars(vecRubros(1,3) & "%",  8, " ", CHR_FWD)                
        strLog = strLog & " | " & GF_nChars(vecRubros(1,2) & "%",  9, " ", CHR_FWD)        
        if (not vecRubros(1,4)) then             
            'Hay algun rubro mal calculado.
            strSQL = "Update hrubrosvisteovagones set VLMERMA=" & vecRubros(1,2) & " WHERE  DTCONTABLE='" & pDtContable & "' and CDOPERATIVO='" & cdOperativoOld & "' and CDVAGON='" & cdVagonOld & "' and SQCALADA=" & sqCaladaOld & " and CDRUBRO=" & vecRubros(1,5) & "; --antes " & vecRubros(1,3)
            if (pAction <> ACCION_CONTROLAR) then Call GF_BD_Puertos(g_strPuerto, rsX, "EXEC", strSQL)                    
            strLog = strLog & " | ERROR: Rubro mal calculado!"
            cantError = cantError + 1
        else                                        
            if  (round(CDbl(mermaCabeceraOld), 2) <> round(CDbl(myTotalMermaDetalleTabla), 2)) then
                'El % de la cabecera y el detalle no coinciden.                
                strLog = strLog & " | ERROR: % Cabecera <> % Detalle!"
                cantError = cantError + 1
            else                                                
                if (CLng(mermaTotalOld) <> CLng(mermaCalculadaOld)) then
                    'Los kilos de merma estan mal.                
                    strLog = strLog & " | ERROR: Kg Tabla <> Kg Calculados!"
                    cantError = cantError + 1
                else
                    strLog = strLog & " | OK!" 
                    cantOK = cantOK + 1 
                end if
            end if
        end if
        logMig.info(strLog)
        'Se dibujan los rubros faltantes.    
        for jdx = 2 to idx
            strLog = GF_nChars("", 10, " ", CHR_FWD)        
            strLog = strLog & " | " & GF_nChars("", 12, " ", CHR_FWD)
            strLog = strLog & " | " & GF_nChars("", 10, " ", CHR_FWD)
            strLog = strLog & " | " & GF_nChars("", 12, " ", CHR_FWD)
            strLog = strLog & " | " & GF_nChars("",  8, " ", CHR_FWD)
            strLog = strLog & " | " & GF_nChars("",  8, " ", CHR_FWD)
            strLog = strLog & " | " & GF_nChars("", 14, " ", CHR_FWD)
            strLog = strLog & " | " & GF_nChars("", 15, " ", CHR_FWD)
            strLog = strLog & " | " & GF_nChars(vecRubros(jdx,1), 20, " ", CHR_AFT)
            strLog = strLog & " | " & GF_nChars(vecRubros(jdx,3) & "%",  8, " ", CHR_FWD)                    
            strLog = strLog & " | " & GF_nChars(vecRubros(jdx,2) & "%",  9, " ", CHR_FWD)
            if (not vecRubros(jdx,4)) then             
                'Hay algun rubro mal calculado.
                strSQL = "Update hrubrosvisteovagones set VLMERMA=" & vecRubros(jdx,2) & " WHERE  DTCONTABLE='" & pDtContable & "' and CDOPERATIVO='" & cdOperativoOld & "' and CDVAGON='" & cdVagonOld & "' and SQCALADA=" & sqCaladaOld & " and CDRUBRO=" & vecRubros(jdx,5) & "; --antes " & vecRubros(jdx,3)
                if (pAction <> ACCION_CONTROLAR) then Call GF_BD_Puertos(g_strPuerto, rsX, "EXEC", strSQL)                    
                strLog = strLog & " | ERROR: Rubro mal calculado!"
                cantError = cantError + 1
            else
                strLog = strLog & " | OK!"  
                cantOK = cantOK + 1
            end if                
            logMig.info(strLog)
        next         
    wend    
    logMig.info(GF_nChars("",180, "-", CHR_AFT))           
    logMig.info("TOTAL PROCESADOS: " & cantOK + cantError & " (OK:" & cantOK & ", ERROR:" & cantError & ")")       
    logMig.info(GF_nChars("",180, "-", CHR_AFT))  
    
    ret = true
    if (cantError > 0) then ret = false
    procesarVagones = ret
             
End Function
'***********************************************
'*****      COMIENZO DE LA PAGINA          *****
'***********************************************
Dim strSQL, rs, DtContable, myHoy, g_listaRubrosHumedad, g_listaRubrosZaranda, listaRubros
Dim myHasta, transporte, myNext, logMig, rsX, resultC, resultV, action

'----------------------------------
'PARAMETROS DEL APLICATIVO
'
' pto: Puerto ARROYO, TRANSITO, PIEDRABUENA
' f: Fecha a controlar [Opcional]
' t: Transporte C: Camiones, V:Vagones [Opcional]
'----------------------------------
g_strPuerto = GF_PARAMETROS7("pto", "", 6)
'Asumo que se va a migrar los datos del dia de ayer.
myHoy = GF_DTE2FN(day(date) & "/" & month(date) & "/" & year(date))
myHoy = GF_DTEADD(myHoy,-1,"D")
'Si pidio una fecha particular tomo esa fecha.
if (GF_PARAMETROS7("f", 0, 6) <> 0) then myHoy = GF_PARAMETROS7("f", 0, 6)
DtContable = Left(myHoy, 4) & "-" & mid(myHoy, 5, 2) & "-" & Right(myHoy, 2)	 

transporte = "T"
if (GF_PARAMETROS7("t", "", 6) <> "") then transporte = GF_PARAMETROS7("t", "", 6)

%>
<html>
<head>
</head>
<body onload="bodyOnLoad()">
<table>
    <tr>
        <td rowspan="2">Carta de Porte</td>
        <td rowspan="2">Corredor</td>
        <td rowspan="2">Vendedor</td>
        <td rowspan="2">Ppia Prod</td>
        <td colspan="4">CONTROL</td>
        
    </tr>
    <tr>        
        <td>Estab. Proce.</td>
        <td>Destinatario</td>
        <td>Titular</td>
        <td>Rte. Comercial</td>
    </tr>
</table>
<%
if ((transporte = "C") or (transporte = "T")) then Call procesarCamiones(g_strPuerto, DtContable)
'if ((transporte = "V") or (transporte = "T")) then resultV = procesarVagones(g_strPuerto, DtContable)
%>
</body>
</html>