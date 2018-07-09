<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientosPuertos.asp"-->
<!--#include file="../Includes/procedimientosFechas.asp"-->
<!--#include file="../Includes/procedimientosTraducir.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->
<!--#include file="../Includes/procedimientosUnificador.asp"-->
<%
CONST PARAM_RUBRO_HUMEDAD    = "CDRUBROHUMEDAD"
CONST PARAM_RUBRO_PH         = "CDRUBROPESOHECTO"
CONST SEPERATOR_RUBROS       = "$"    
CONST SECCION_TN_TOTAL = 1
CONST SECCION_VELOCIDAD = 2
CONST SECCION_TIEMPO_PROMEDIO = 3
CONST SECCION_CAMION_PUERTO = 4
CONST SECCION_CALIDAD = 5
'---------------------------------------------------------------------------------------------------------------------------------------
Function getSQLDescargaKg(pFechaDesde,pFechaHasta,pTransporte,pPto,pOrderByHora)
	Dim strSQL,rs,diaHoy
	diaHoy = Year(Now()) & "-" & GF_nDigits(Month(Now()), 2) & "-" & GF_nDigits(Day(Now()), 2)
    strSQLCam = ""
    strSQLVag = ""
    if ((Cdbl(pTransporte) = TIPO_TRANSPORTE_CAMION)or(Cdbl(pTransporte) = TIPO_TRANSPORTE_CAMVAG)) then
             strSQLCam = " (SELECT CASE WHEN Sum (vlneto) IS NULL THEN 0 ELSE Sum (vlneto) END AS neto_descarga "
             if (pOrderByHora) then strSQLCam = strSQLCam & ",hora "
             strSQLCam = strSQLCam & " FROM (SELECT a.*, "&_
             "                          ( "&_
             "           		            (SELECT CASE WHEN HC.vlpesada IS NULL THEN 0 ELSE HC.vlpesada END AS vlpesada "&_
             "              	             FROM   hpesadascamion HC "&_
             "              	             WHERE  HC.cdpesada = 1 AND HC.dtcontable = a.dtcontable AND HC.idcamion = a.idcamion "&_
             "                     	            AND HC.sqpesada = (SELECT Max(sqpesada) "&_
             "                          				            FROM  hpesadascamion "&_
             "                          				            WHERE dtcontable = HC.dtcontable AND HC.idcamion = idcamion AND cdpesada = 1)) - "&_
             "               	            (SELECT CASE WHEN HC.vlpesada IS NULL THEN 0 ELSE HC.vlpesada END AS vlpesada "&_
             "               	             FROM hpesadascamion HC "&_
             "               	             WHERE HC.cdpesada = 2 AND HC.dtcontable = a.dtcontable AND HC.idcamion = a.idcamion "&_
             "               		            AND HC.sqpesada = (SELECT Max(sqpesada) "&_
             "                                  		            FROM  hpesadascamion "&_
             "                                  		            WHERE HC.dtcontable = dtcontable AND HC.idcamion = idcamion AND cdpesada = 2)) - "&_
             "             	            (SELECT CASE WHEN HMC.vlmermakilos IS NULL THEN 0 ELSE HMC.vlmermakilos END AS VLMERMAKILOS "&_
             "              	             FROM hmermascamiones HMC "&_
             "              	             WHERE HMC.dtcontable = a.dtcontable AND HMC.idcamion = a.idcamion "&_
             "               		            AND HMC.sqpesada = (SELECT Max(sqpesada) "&_
             "                                   		            FROM   hpesadascamion "&_
             "                                   		            WHERE HMC.dtcontable = dtcontable AND HMC.idcamion = idcamion and cdpesada = 2))  "&_
             "                         ) AS vlNETO "&_
             "                 FROM   (SELECT B.* "&_
             "                         FROM (SELECT dtcontable, idcamion "
             if (pOrderByHora) then strSQLCam = strSQLCam & ",DATEPART(HOUR, dtegreso) AS hora "
             strSQLCam = strSQLCam & "               FROM   hcamiones "&_
             "                                 WHERE  dtcontable >= '"& pFechaDesde &"'  AND dtcontable <= '"& pFechaHasta &"'  AND cdestado IN( 6, 8 )) B "&_
             "                         INNER JOIN hcamionesdescarga c ON c.idcamion = B.idcamion AND c.dtcontable = b.dtcontable) A "&_
             "               ) AS T1 "
             if (pOrderByHora) then strSQLCam = strSQLCam & " GROUP BY hora "
             strSQLCam = strSQLCam & " ) "&_
             "         UNION  "&_
             "         (SELECT CASE WHEN Sum (vlneto) IS NULL THEN 0 ELSE Sum (vlneto) END AS neto_descarga "
             if (pOrderByHora) then strSQLCam = strSQLCam & ",hora "
             strSQLCam = strSQLCam & " FROM (SELECT a.*, "&_
             "                       ( (SELECT CASE WHEN HC.vlpesada IS NULL THEN 0 ELSE HC.vlpesada END AS vlpesada "&_
             "                          FROM   pesadascamion HC "&_
             "                          WHERE  HC.cdpesada = 1 AND A.idcamion = hc.idcamion "&_
             "                                 AND HC.sqpesada = (SELECT Max(sqpesada) "&_
             "                          				          FROM  pesadascamion "&_
             "                          				          WHERE  cdpesada = 1 AND idcamion = Hc.idcamion)) - "&_
             "                           (SELECT CASE WHEN HC.vlpesada IS NULL THEN 0 ELSE HC.vlpesada END AS vlpesada "&_
             "                            FROM   pesadascamion HC "&_
             "                            WHERE  HC.cdpesada = 2 AND A.idcamion = hc.idcamion "&_
             "                                   AND sqpesada = (SELECT Max(sqpesada) "&_
             "                            			             FROM  pesadascamion "&_
             "                            			             WHERE  cdpesada = 2 AND idcamion = HC.idcamion)) - "&_
             "                           (SELECT CASE WHEN MC.vlmermakilos IS NULL THEN 0 ELSE MC.vlmermakilos END AS VLMERMAKILOS "&_
             "                            FROM   mermascamiones MC "&_
             "                            WHERE  MC.idcamion = A.idcamion "&_
             "                                   AND MC.sqpesada = (SELECT Max(sqpesada) "&_
             "                                                      FROM   pesadascamion "&_
             "                                                      WHERE MC.idcamion = idcamion and cdpesada = 2)) "&_
             "           	            ) AS vlNETO "&_
             "                FROM   (SELECT B.* "&_
             "                        FROM   (SELECT '"&diaHoy&"'  AS dtcontable, idcamion "
             if (pOrderByHora) then strSQLCam = strSQLCam & ",DATEPART(HOUR, dtegreso) AS hora "
             strSQLCam = strSQLCam & "        FROM   camiones "&_
             "                                WHERE  cdestado IN( 6, 8 )) B "&_
             "                        INNER JOIN camionesdescarga c "&_
             "             	            ON c.idcamion = B.idcamion"&_
             "                        WHERE B.DTCONTABLE >= '"& pFechaDesde &"' AND B.DTCONTABLE <= '"& pFechaHasta &"'"
             if (pOrderByHora) then strSQLCam = strSQLCam & " AND B.HORA IS NOT NULL "
             strSQLCam = strSQLCam & " ) A "&_
             "        ) AS T1 "
             if (pOrderByHora) then strSQLCam = strSQLCam & " GROUP  BY hora "
             strSQLCam = strSQLCam & " ) "
        end if
        if ((Cdbl(pTransporte) = TIPO_TRANSPORTE_VAGON)or(Cdbl(pTransporte) = TIPO_TRANSPORTE_CAMVAG)) then
             strSQLVag = " ( SELECT CASE WHEN Sum (vlneto) IS NULL THEN 0 ELSE Sum (vlneto) END AS neto_descarga "
             if (pOrderByHora) then strSQLVag = strSQLVag & ",hora "
             strSQLVag = strSQLVag & " FROM (SELECT a.*, "&_
             "                       ( "&_
             "           		            (SELECT CASE WHEN HC.vlpesada IS NULL THEN 0 ELSE HC.vlpesada END AS vlpesada "&_
             "              	             FROM   HPESADASVAGON HC "&_
             "              	             WHERE  HC.cdpesada = 1 AND HC.dtcontable = a.dtcontable AND HC.CDOPERATIVO = a.CDOPERATIVO and hc.cdoperativoserie = a.cdoperativoserie AND HC.CDVAGON = A.CDVAGON "&_
             "                     	            AND HC.sqpesada = (SELECT Max(sqpesada) "&_
             "                          				            FROM HPESADASVAGON "&_
             "                          				            WHERE dtcontable = HC.dtcontable AND HC.CDOPERATIVO = CDOPERATIVO AND hc.cdoperativoserie = cdoperativoserie AND HC.CDVAGON = CDVAGON AND cdpesada = 1)) - "&_
             "               	            (SELECT CASE WHEN HC.vlpesada IS NULL THEN 0 ELSE HC.vlpesada END AS vlpesada "&_
             "              	             FROM   HPESADASVAGON HC "&_
             "              	             WHERE  HC.cdpesada = 2 AND HC.dtcontable = a.dtcontable AND HC.CDOPERATIVO = a.CDOPERATIVO and hc.cdoperativoserie = a.cdoperativoserie AND HC.CDVAGON = A.CDVAGON "&_
             "                     	            AND HC.sqpesada = (SELECT Max(sqpesada) "&_
             "                          				            FROM HPESADASVAGON "&_
             "                          				            WHERE dtcontable = HC.dtcontable AND HC.CDOPERATIVO = CDOPERATIVO AND hc.cdoperativoserie = cdoperativoserie AND HC.CDVAGON = CDVAGON AND cdpesada = 2)) - "&_
             "             	                (SELECT CASE WHEN HMC.vlmermakilos IS NULL THEN 0 ELSE HMC.vlmermakilos END AS VLMERMAKILOS "&_
             "              	             FROM HMERMASVAGONES HMC "&_
             "              	             WHERE HMC.dtcontable = a.dtcontable AND HMC.cdoperativo = a.cdoperativo and hmc.cdoperativoserie = a.cdoperativoserie and hmc.cdvagon = a.cdvagon "&_
             "               		            AND HMC.sqpesada = (SELECT Max(sqpesada) "&_
             "                                   		            FROM   HPesadasVagon "&_
             "                                   		            WHERE HMC.dtcontable = dtcontable AND HMC.cdoperativo = cdoperativo AND HMC.cdoperativoSERIE = cdoperativoSERIE AND HMC.cdvagon = cdvagon and cdpesada = 2))  "&_
             "                         ) AS vlNETO "&_
             "                 FROM   (SELECT B.* "&_
             "                         FROM   (SELECT dtcontable, cdoperativo,cdoperativoserie,cdvagon "
             if (pOrderByHora) then strSQLVag = strSQLVag & ",DATEPART(HOUR, dtfin) AS hora "
             strSQLVag = strSQLVag & "         FROM   hvagones "&_
             "                                 WHERE  dtcontable >= '"& pFechaDesde &"' AND dtcontable <= '"& pFechaHasta &"' AND cdestado = 8 ) B ) A "&_
             "               ) AS T1 "
             if (pOrderByHora) then strSQLVag = strSQLVag & " GROUP  BY hora "
             strSQLVag = strSQLVag & " ) "&_
 	         "       UNION "&_
             "        ( "&_
		     "       SELECT CASE WHEN Sum (vlneto) IS NULL THEN 0 ELSE Sum (vlneto) END AS neto_descarga "
             if (pOrderByHora) then strSQLVag = strSQLVag & ",hora "
             strSQLVag = strSQLVag & " FROM (SELECT a.*, "&_
             "                       ( "&_
             "           		            (SELECT CASE WHEN HC.vlpesada IS NULL THEN 0 ELSE HC.vlpesada END AS vlpesada "&_
             "              	             FROM   PESADASVAGON HC "&_
             "              	             WHERE  HC.cdpesada = 1 AND HC.CDOPERATIVO = a.CDOPERATIVO and hc.cdoperativoserie = a.cdoperativoserie AND HC.CDVAGON = A.CDVAGON "&_
             "                     	            AND HC.sqpesada = (SELECT Max(sqpesada) "&_
             "                          				            FROM PESADASVAGON  "&_
             "                          				            WHERE HC.CDOPERATIVO = CDOPERATIVO AND hc.cdoperativoserie = cdoperativoserie AND HC.CDVAGON = CDVAGON AND cdpesada = 1)) - "&_
             "               	            (SELECT CASE WHEN HC.vlpesada IS NULL THEN 0 ELSE HC.vlpesada END AS vlpesada "&_
             "              	             FROM   PESADASVAGON HC "&_
             "              	             WHERE  HC.cdpesada = 2 AND HC.CDOPERATIVO = a.CDOPERATIVO and hc.cdoperativoserie = a.cdoperativoserie AND HC.CDVAGON = A.CDVAGON "&_
             "                     	            AND HC.sqpesada = (SELECT Max(sqpesada) "&_
             "                          				            FROM PESADASVAGON  "&_
             "                          				            WHERE HC.CDOPERATIVO = CDOPERATIVO AND hc.cdoperativoserie = cdoperativoserie AND HC.CDVAGON = CDVAGON AND cdpesada = 2)) - "&_
             "             	                (SELECT CASE WHEN HMC.vlmermakilos IS NULL THEN 0 ELSE HMC.vlmermakilos END AS VLMERMAKILOS "&_
             "              	             FROM MERMASVAGONES HMC "&_
             "              	             WHERE HMC.cdoperativo = a.cdoperativo and hmc.cdoperativoserie = a.cdoperativoserie and hmc.cdvagon = a.cdvagon "&_
             "               		            AND HMC.sqpesada = (SELECT Max(sqpesada) "&_
             "                                   		            FROM   PesadasVagon "&_
             "                                   		            WHERE HMC.cdoperativo = cdoperativo AND HMC.cdoperativoSERIE = cdoperativoSERIE AND HMC.cdvagon = cdvagon and cdpesada = 2))  "&_
             "                         ) AS vlNETO "&_
             "                 FROM   (SELECT B.* "&_
             "                         FROM   (SELECT  '"&diaHoy&"' as dtcontable, cdoperativo,cdoperativoserie,cdvagon "
             if (pOrderByHora) then strSQLVag = strSQLVag & ",DATEPART(HOUR, dtfin) AS hora "
             strSQLVag = strSQLVag & "         FROM   vagones "&_
             "                                 WHERE  cdestado = 8 ) B "&_
             "                          WHERE B.DTCONTABLE >= '"& pFechaDesde &"' AND B.DTCONTABLE <= '"& pFechaHasta &"'"
             if (pOrderByHora) then strSQLVag = strSQLVag & " AND B.HORA IS NOT NULL "
             strSQLVag = strSQLVag & " ) A "&_
             "               ) AS T1 "
             if (pOrderByHora) then strSQLVag = strSQLVag & " GROUP  BY hora "
             strSQLVag = strSQLVag & " ) "
        end if
        
        if ((strSQLCam <> "")and(strSQLVag <> "")) then auxUnion= " UNION "
        strSQL = "SELECT SUM(NETO_DESCARGA) AS KGDESCARGA "
        if (pOrderByHora) then strSQL = strSQL & ",HORA "
        strSQL = strSQL & "FROM ( "& strSQLCam & auxUnion & strSQLVag & " ) AS TFINAL "
        if (pOrderByHora) then strSQL = strSQL & " GROUP BY HORA "
    Call GF_BD_Puertos(pPto, rs, "OPEN", strSQL)
	Set getSQLDescargaKg = rs
End Function
'---------------------------------------------------------------------------------------------------------------------------------------
'Obtiene los kilos/Toneladas que se descargaron en un determinado periodo. Dependiendo del parametro pUnidadPeso determina si es en Kilos o Toneladas
Function getDescargaKg(pFechaDesde,pFechaHasta,pTransporte,pStrPuerto)
    Dim rsDesKg
    Set rsDesKg = getSQLDescargaKg(pFechaDesde,pFechaHasta,pTransporte,pStrPuerto,False)
    rtrn = 0
    if (not rsDesKg.Eof) then
        rtrn = Cdbl(rsDesKg("KGDESCARGA"))/1000
        rtrn = GF_EDIT_DECIMALS(rtrn*100,2)
    end if
    getDescargaKg = rtrn
End Function
'---------------------------------------------------------------------------------------------------------------------------------------
Function getSQLDescargaCalidad(pFechaDesde,pFechaHasta,pCdProducto,pRubroHumedad,pRubroPh,pTransporte,pPto)
    Dim strSQL,rs,diaHoy,strSQLVag,strSQLCam
	diaHoy = Year(Now()) & "-" & GF_nDigits(Month(Now()), 2) & "-" & GF_nDigits(Day(Now()), 2)
    strSQLCam = ""
    strSQLVag = ""
    if ((Cdbl(pTransporte) = TIPO_TRANSPORTE_CAMION)or(Cdbl(pTransporte) = TIPO_TRANSPORTE_CAMVAG)) then
        strSQLCam = "(SELECT DISTINCT "&TIPO_TRANSPORTE_CAMION&" as transporte,hcc.dtcontable AS FContable, hcc.dtcalada AS DtCalada, hcc.idcamion AS ID, hrvc.cdrubro AS Rubro, hrvc.vlbonrebaja AS Valor, "&_
                    "                 (SELECT CASE WHEN HPC.vlpesada IS NULL THEN 0 ELSE HPC.vlpesada END AS vlpesada "&_
                    "                  FROM hpesadascamion HPC "&_
                    "                  WHERE HPC.dtcontable = hcc.dtcontable AND HPC.idcamion = hcc.idcamion AND HPC.cdpesada = 1 "&_
                    "                           AND HPC.sqpesada = (SELECT Max(sqpesada) "&_
                    "                                               FROM   hpesadascamion "&_
                    "                                               WHERE dtcontable = HPC.dtcontable AND idcamion = HPC.idcamion AND cdpesada = 1)) AS Bruto, "&_
                    "                   (SELECT CASE WHEN HPC.vlpesada IS NULL THEN 0 ELSE HPC.vlpesada END AS vlpesada "&_
                    "                    FROM   hpesadascamion HPC "&_
                    "                    WHERE  HPC.dtcontable = hcc.dtcontable AND HPC.idcamion = hcc.idcamion AND HPC.cdpesada = 2 "&_
                    "                           AND HPC.sqpesada = (SELECT Max(sqpesada) "&_
                    "                                               FROM   hpesadascamion "&_
                    "                                               WHERE dtcontable = HPC.dtcontable AND idcamion = HPC.idcamion AND cdpesada = 2)) AS Tara, "&_
                    "                   (SELECT CASE WHEN HMC.vlmermakilos IS NULL THEN 0 ELSE HMC.vlmermakilos END AS VLMERMAKILOS "&_
                    "                    FROM hmermascamiones HMC "&_
                    "                    WHERE  HMC.dtcontable = hcc.dtcontable AND HMC.idcamion = hcc.idcamion "&_
                    "                           AND HMC.sqpesada = (SELECT Max(sqpesada) "&_
                    "                                               FROM   hpesadascamion "&_
                    "                                               WHERE dtcontable = HMC.dtcontable AND idcamion = HMC.idcamion AND cdpesada = 2)) AS KgMerma "&_
                    "FROM   (SELECT HCAL.DTCONTABLE,HCAL.DTCALADA,HCAL.IDCAMION,HCAL.SQCALADA "&_
        	        "	     FROM  hcaladadecamiones HCAL  "&_
        	        "	     WHERE  cast(HCAL.dtcalada AS date) >= '"& pFechaDesde &"'"&_
                    "	            AND cast(HCAL.dtcalada AS date) <= '"& pFechaHasta &"'"&_
        	        "	 	        AND HCAL.sqcalada = (SELECT Max(sqcalada) "&_
                    "                  		             FROM   hcaladadecamiones "&_
                    "                  		             WHERE  idcamion = HCAL.idcamion AND dtcontable = HCAL.dtcontable))hcc "&_
                    "          INNER JOIN hcamiones hc ON hcc.dtcontable = hc.dtcontable AND hcc.idcamion = hc.idcamion "
                    if (Cdbl(pCdProducto) <> 0) then strSQLCam = strSQLCam & " and hc.cdproducto = "& pCdProducto 
                    strSQLCam = strSQLCam & " INNER JOIN hrubrosvisteocamiones hrvc ON hcc.idcamion = hrvc.idcamion AND hcc.dtcontable = hrvc.dtcontable AND hcc.sqcalada = hrvc.sqcalada "&_
                    "   WHERE hc.cdestado IN (SELECT cdestado FROM taskestados WHERE cdtask = 412 AND cdtipocamion = 1))  "&_
                    " UNION "&_
	                " (SELECT DISTINCT "&TIPO_TRANSPORTE_CAMION&" as transporte,"&_
                    "                  hcc.dtcontable AS FContable, "&_
                    "                  hcc.dtcalada AS DtCalada, "&_
                    "                  hcc.idcamion AS ID, "&_
                    "                  hrvc.cdrubro AS Rubro, "&_
                    "                  hrvc.vlbonrebaja AS Valor, "&_
                    "                  (SELECT CASE WHEN PC.vlpesada IS NULL THEN 0 ELSE PC.vlpesada END AS vlpesada "&_
                    "                   FROM   pesadascamion PC "&_
                    "                   WHERE  PC.idcamion = hcc.idcamion AND PC.cdpesada = 1 "&_
                    "                                   AND PC.sqpesada = (SELECT Max(sqpesada) "&_
                    "                                                      FROM pesadascamion "&_
                    "                                                      WHERE idcamion = PC.idcamion AND cdpesada = 1)) AS Bruto, "&_
                    "                   (SELECT CASE WHEN PC.vlpesada IS NULL THEN 0 ELSE PC.vlpesada END AS vlpesada "&_
                    "                    FROM   pesadascamion PC "&_
                    "                    WHERE  PC.idcamion = hcc.idcamion AND PC.cdpesada = 2 "&_
                    "                                   AND PC.sqpesada = (SELECT Max(sqpesada) "&_
                    "                                                      FROM  pesadascamion "&_
                    "                                                      WHERE idcamion = PC.idcamion AND cdpesada = 2)) AS Tara, "&_
                    "                   (SELECT CASE WHEN HMC.vlmermakilos IS NULL THEN 0 ELSE HMC.vlmermakilos END AS VLMERMAKILOS "&_
                    "                    FROM mermascamiones HMC "&_
                    "                    WHERE HMC.idcamion = hcc.idcamion "&_
                    "                                   AND HMC.sqpesada = (SELECT Max(sqpesada) "&_
                    "                                                       FROM pesadascamion "&_
                    "                                                       WHERE idcamion = HMC.idcamion AND cdpesada = 2)) AS KgMerma "&_
                    " FROM   (SELECT '"& diaHoy &"' AS DTCONTABLE,HCAL.DTCALADA,HCAL.IDCAMION,HCAL.SQCALADA "&_
			        "         FROM  caladadecamiones HCAL  "&_
        	        "	      WHERE  cast(HCAL.dtcalada AS date) >= '"& pFechaDesde &"'"&_
                    "	            and cast(HCAL.dtcalada AS date) <= '"& pFechaHasta &"'"&_
        	        "	 	        AND HCAL.sqcalada = (SELECT Max(sqcalada) "&_
                    "                  		             FROM   caladadecamiones "&_
                    "                  		             WHERE  idcamion = HCAL.idcamion) "&_
        	        "	      )hcc "&_
                    "  INNER JOIN camiones hc ON hcc.idcamion = hc.idcamion "
                    if (Cdbl(pCdProducto) <> 0) then strSQLCam = strSQLCam & " and hc.cdproducto = "& pCdProducto 
                    strSQLCam = strSQLCam & " INNER JOIN rubrosvisteocamiones hrvc ON hcc.idcamion = hrvc.idcamion AND hcc.sqcalada = hrvc.sqcalada "&_
                    "  WHERE hc.cdestado IN (SELECT cdestado FROM taskestados WHERE cdtask = 412 AND cdtipocamion = 1))"
    end if
    if ((Cdbl(pTransporte) = TIPO_TRANSPORTE_VAGON)or(Cdbl(pTransporte) = TIPO_TRANSPORTE_CAMVAG)) then
        strSQLVag = "(SELECT DISTINCT "&TIPO_TRANSPORTE_VAGON&" as transporte,"&_
                    "                 hcc.dtcontable AS FContable, "&_
                    "                 hcc.dtcalada AS DtCalada, "&_
                    "                 hcc.cdvagon as ID, "&_
                    "                 hrvc.cdrubro AS Rubro , "&_
                    "                 hrvc.vlbonrebaja AS Valor, "&_
                    "                (SELECT CASE WHEN HPV.vlpesada IS NULL THEN 0 ELSE HPV.vlpesada END AS vlpesada "&_
                    "                 FROM   HPESADASVAGON HPV "&_
                    "                 WHERE  HPV.dtcontable = hcc.dtcontable AND HPV.cdoperativo = hcc.cdoperativo AND HPV.cdoperativoserie = hcc.cdoperativoserie AND HPV.cdvagon = hcc.cdvagon AND HPV.cdpesada = 1 "&_
                    "                     AND HPV.sqpesada = (SELECT Max(sqpesada) "&_
                    "                                       FROM   HPESADASVAGON "&_
                    "                                       WHERE dtcontable = HPV.dtcontable AND cdoperativo = HPV.cdoperativo AND HPV.cdoperativoserie = cdoperativoserie AND HPV.cdvagon = cdvagon AND cdpesada = 1)) AS Bruto , "&_
                    "                (SELECT CASE WHEN HPV.vlpesada IS NULL THEN 0 ELSE HPV.vlpesada END AS vlpesada "&_
                    "                 FROM   HPESADASVAGON HPV "&_
                    "                 WHERE  HPV.dtcontable = hcc.dtcontable AND HPV.cdoperativo = hcc.cdoperativo AND HPV.cdoperativoserie = hcc.cdoperativoserie AND HPV.cdvagon = hcc.cdvagon AND HPV.cdpesada = 2 "&_
                    "                     AND HPV.sqpesada = (SELECT Max(sqpesada) "&_
                    "                                       FROM   HPESADASVAGON "&_
                    "                                       WHERE dtcontable = HPV.dtcontable AND cdoperativo = HPV.cdoperativo AND HPV.cdoperativoserie = cdoperativoserie AND HPV.cdvagon = cdvagon AND cdpesada = 2)) AS  Tara, "&_
					"                (SELECT CASE WHEN pv.vlMermaKilos IS NULL THEN 0 ELSE pv.vlMermaKilos END AS vlMermaKilos  "&_
					"                 FROM HMermasvagones pv "&_
    				"	              WHERE pv.dtcontable = hcc.dtcontable and pv.cdoperativo = hcc.cdoperativo and pv.cdoperativoSERIE = hcc.cdoperativoSERIE and pv.cdvagon = hcc.cdvagon "&_
    			    "			          AND pv.sqpesada = (SELECT MAX(sqPesada)  "&_
    				"		 				               FROM HPesadasVagon "&_
    				"						               WHERE dtcontable = pv.dtcontable and cdoperativo = pv.cdoperativo and cdoperativoSERIE = pv.cdoperativoSERIE and cdvagon=pv.cdvagon and cdPesada =2)) AS KgMerma "&_
                    "FROM (SELECT HCAL.dtcontable, HCAL.cdoperativo,HCAL.cdoperativoserie,HCAL.cdvagon,HCAL.dtcalada,HCAL.sqcalada "&_
                    "      FROM   HCALADADEVAGONES HCAL "&_
                    "      WHERE  cast(HCAL.dtcalada AS date) >= '"& pFechaDesde &"'"&_
                    "             AND cast(HCAL.dtcalada AS date) <= '"& pFechaHasta &"'"&_
                    "             AND HCAL.sqcalada = (SELECT Max(sqcalada)  "&_
                    "               					FROM HCALADADEVAGONES  "&_
                    "               					WHERE dtcontable = HCAL.dtcontable AND cdoperativo = HCAL.cdoperativo AND cdoperativoserie = HCAL.cdoperativoserie AND cdvagon = HCAL.cdvagon) "&_
		            "	  )hcc "&_
                    "       INNER JOIN hvagones hc "&_
                    "               ON hcc.dtcontable = hc.dtcontable AND hcc.cdoperativo = hc.cdoperativo and hcc.cdoperativoserie = hc.cdoperativoserie and hcc.cdvagon = hc.cdvagon "
                    if (Cdbl(pCdProducto) <> 0) then strSQLVag = strSQLVag & " and hc.cdproducto = "& pCdProducto
                    strSQLVag = strSQLVag & " INNER JOIN HRUBROSVISTEOVAGONES hrvc "&_
                    "       ON hcc.dtcontable = hrvc.dtcontable AND hcc.cdoperativo = hrvc.cdoperativo and hcc.cdoperativoserie = hrvc.cdoperativoserie AND hc.cdvagon = hrvc.cdvagon AND hcc.sqcalada = hrvc.sqcalada "&_
                    "WHERE HC.CDESTADO IN (6,8)) "&_
                    "UNION "&_
                    "(SELECT DISTINCT  "&TIPO_TRANSPORTE_VAGON&" as transporte,"&_
                    "                 hcc.dtcontable AS FContable, "&_
                    "                 hcc.dtcalada AS DtCalada, "&_
                    "                 hcc.cdvagon AS ID, "&_
                    "                 hrvc.cdrubro AS Rubro , "&_
                    "                 hrvc.vlbonrebaja AS Valor, "&_
                    "                (SELECT CASE WHEN HPV.vlpesada IS NULL THEN 0 ELSE HPV.vlpesada END AS vlpesada "&_
                    "                 FROM PESADASVAGON HPV "&_
                    "                 WHERE  HPV.cdoperativo = hcc.cdoperativo AND HPV.cdoperativoserie = hcc.cdoperativoserie AND HPV.cdvagon = hcc.cdvagon AND HPV.cdpesada = 1 "&_
                    "                       AND HPV.sqpesada = (SELECT Max(sqpesada) "&_
                    "                                           FROM   PESADASVAGON "&_
                    "                                           WHERE cdoperativo = HPV.cdoperativo AND HPV.cdoperativoserie = cdoperativoserie AND HPV.cdvagon = cdvagon AND cdpesada = 1)) AS Bruto , "&_
                    "                (SELECT CASE WHEN HPV.vlpesada IS NULL THEN 0 ELSE HPV.vlpesada END AS vlpesada "&_
                    "                 FROM   PESADASVAGON HPV "&_
                    "                 WHERE  HPV.cdoperativo = hcc.cdoperativo AND HPV.cdoperativoserie = hcc.cdoperativoserie AND HPV.cdvagon = hcc.cdvagon AND HPV.cdpesada = 2 "&_
                    "                       AND HPV.sqpesada = (SELECT Max(sqpesada) "&_
                    "                                           FROM   PESADASVAGON "&_
                    "                                           WHERE cdoperativo = HPV.cdoperativo AND HPV.cdoperativoserie = cdoperativoserie AND HPV.cdvagon = cdvagon AND cdpesada = 2)) AS  Tara, "&_
					"                (SELECT CASE WHEN pv.vlMermaKilos IS NULL THEN 0 ELSE pv.vlMermaKilos END AS vlMermaKilos  "&_
					"                 FROM Mermasvagones pv "&_
    				"	              WHERE pv.cdoperativo = hcc.cdoperativo and pv.cdoperativoSERIE = hcc.cdoperativoSERIE and pv.cdvagon = hcc.cdvagon "&_
    				"		                and pv.sqpesada = (SELECT MAX(sqPesada)  "&_
    				"		 				                   FROM PesadasVagon "&_
    				"						                   WHERE cdoperativo = pv.cdoperativo and cdoperativoSERIE = pv.cdoperativoSERIE and cdvagon=pv.cdvagon and cdPesada =2)) AS KgMerma "&_
                    "FROM   (SELECT '"& diaHoy &"' as dtcontable, HCAL.cdoperativo,HCAL.cdoperativoserie,HCAL.cdvagon,HCAL.dtcalada,HCAL.sqcalada "&_
                    "        FROM   CALADADEVAGONES HCAL "&_
                    "        WHERE  cast(HCAL.dtcalada AS date) >= '"& pFechaDesde &"' "&_
                    "               AND cast(HCAL.dtcalada AS date) <= '"& pFechaHasta &"' "&_
                    "               AND HCAL.sqcalada = (SELECT Max(sqcalada)  "&_
                    "               					FROM CALADADEVAGONES  "&_
                    "               					WHERE cdoperativo = HCAL.cdoperativo AND cdoperativoserie = HCAL.cdoperativoserie AND cdvagon = HCAL.cdvagon) "&_
		            "	    )hcc "&_
                    "       INNER JOIN vagones hc "&_
                    "               ON hcc.cdoperativo = hc.cdoperativo and hcc.cdoperativoserie = hc.cdoperativoserie and hcc.cdvagon = hc.cdvagon "
                    if (Cdbl(pCdProducto) <> 0) then strSQLVag = strSQLVag & " and hc.cdproducto = "& pCdProducto
                    strSQLVag = strSQLVag & " INNER JOIN RUBROSVISTEOVAGONES hrvc "&_
                    "               ON hcc.cdoperativo = hrvc.cdoperativo and hcc.cdoperativoserie = hrvc.cdoperativoserie AND hc.cdvagon = hrvc.cdvagon AND hcc.sqcalada = hrvc.sqcalada "&_
                    " WHERE hc.CDESTADO IN(6,8)) "
    end if
    if (((Cdbl(pTransporte) = TIPO_TRANSPORTE_VAGON)and(Cdbl(pTransporte) = TIPO_TRANSPORTE_CAMION))or(Cdbl(pTransporte) = TIPO_TRANSPORTE_CAMVAG)) then strUnion = "UNION"
    strSQL = "SELECT CCC.transporte,"&_
             "       ((YEAR(CCC.fcontable) * 10000) + (MONTH(CCC.fcontable) * 100) + DAY(CCC.fcontable)) AS fcontable, "&_
             "       id,CCC.rubro,R.dsrubro, "&_
             "       CASE WHEN CCC.valor IS NULL THEN 0 ELSE CCC.valor END AS valor, "&_
             "       CASE WHEN CCC.bruto IS NULL THEN 0 ELSE CCC.bruto END AS bruto, "&_
             "       CASE WHEN CCC.tara IS NULL THEN 0 ELSE CCC.tara END AS tara, "&_
             "       CASE WHEN CCC.KgMerma IS NULL THEN 0 ELSE CCC.KgMerma END AS KgMerma "&_
             " FROM ( "& strSQLCam & strUnion & strSQLVag & " ) CCC "&_
             "LEFT JOIN rubros r ON CCC.rubro = r.cdrubro "&_
             "WHERE (CCC.RUBRO IN ("& pRubroHumedad &","& pRubroPh &") OR (CCC.KgMerma <> 0)) "&_
             "ORDER  BY CCC.transporte,CCC.fcontable, CCC.ID, CCC.rubro ASC "
    Call GF_BD_Puertos(pPto, rs, "OPEN", strSQL)
	Set getSQLDescargaCalidad = rs
End Function
'---------------------------------------------------------------------------------------------------------------------------------------
'Obtiene el promedio de un rubro, dependiendo del camion y de la fecha contable
Function promediarValorRubroCamion(pCdRubro,pDtContable,pIdCamion,pDsRubro,pValor,pStrPuerto)
    Select Case Cdbl(pCdRubro)
        Case Cdbl(g_auxRubrosPH)
            strSQL = "SELECT ROUND(AVG(VLPESO),2)valor FROM HMUESTRASHUMEDCAMIONES HMHC WHERE HMHC.IDCAMION='" & pIdCamion & "' AND HMHC.DTCONTABLE ='" & GF_FN2DTCONTABLE(pDtContable) & "' group by dtcontable,idcamion"
        Case Cdbl(g_auxRubrosHumedad)
            strSQL = "SELECT AVG(VLHUMEDAD)valor FROM HMUESTRASHUMEDCAMIONES HMHC WHERE HMHC.IDCAMION='" & pIdCamion & "' AND HMHC.DTCONTABLE ='" & GF_FN2DTCONTABLE(pDtContable) & "' AND HMHC.sqcalada =(SELECT MAX(sqcalada) from hmuestrashumedcamiones where idcamion=HMHC.idcamion and dtcontable=HMHC.dtcontable) group by dtcontable,idcamion"
    End Select
    Call GF_BD_Puertos(pStrPuerto, rsVl, "OPEN", strSQL)
    auxVal = pValor
    if (not rsVl.Eof) Then auxVal = Round(rsVl("valor"), 2)
	promediarValorRubroCamion = auxVal
End Function
'--------------------------------------------------------------------------------------------------------------
'Obtiene los datos de la Calidad, en un parametro almacenado se encuentra los rubros que se deben calcular.
'La operatoria es similar al que tenemos en reporte Calador
'El formato que devuelve es un string como el siguiente:
'       Valor_PH $ Valor_Humedad $ Valor_Merma
Function getDescargaCalidad(pFechaDesde,pFechaHasta,pCdProducto,pTransporte,pStrPuerto)
    Dim rsCalCam,rsCalVag,auxParametrosRubros,oDiccNetoXRubro,oDiccNeto,rtrn,auxMerma
    Set oDiccNeto = createObject("Scripting.Dictionary")
    Set oDiccNetoXRubro = createObject("Scripting.Dictionary")
    rtrn = ""
    g_auxRubrosPH = getValueParametro(PARAM_RUBRO_PH,pStrPuerto)
    g_auxRubrosHumedad = getValueParametro(PARAM_RUBRO_HUMEDAD,pStrPuerto)
    if ((g_auxRubrosHumedad <> "")and(g_auxRubrosPH <> "")) then
        oDiccNetoXRubro.Add Cdbl(g_auxRubrosPH), 0
        oDiccNetoXRubro.Add Cdbl(g_auxRubrosHumedad), 0
        Set rsCalCam = getSQLDescargaCalidad(pFechaDesde,pFechaHasta,pCdProducto,g_auxRubrosHumedad,g_auxRubrosPH,pTransporte,pStrPuerto)
        if (not rsCalCam.Eof) then
            while not rsCalCam.Eof
                'Primero pregunto si el rubro es el que debo mostrar en pantalla, por que hay otros rubros para calcular el promedio de la merma
                if (oDiccNetoXRubro.Exists(cdbl(rsCalCam("rubro")))) then
                    if (Cdbl(rsCalCam("transporte")) = TIPO_TRANSPORTE_CAMION) then
                        auxValor = promediarValorRubroCamion(rsCalCam("rubro"), rsCalCam("fcontable"), rsCalCam("id"),rsCalCam("dsRubro"),rsCalCam("valor"),pStrPuerto)
                    else
                        auxValor = Cdbl(rsCalCam("valor"))
                    end if
                    dNeto = Cdbl(rsCalCam("bruto")) - Cdbl(rsCalCam("tara"))
                    oDiccNetoXRubro.Item(cdbl(rsCalCam("rubro"))) = oDiccNetoXRubro.Item(cdbl(rsCalCam("rubro"))) + (dNeto * CDbl(auxValor))
                    oDiccNeto.Item(cdbl(rsCalCam("rubro"))) = oDiccNeto.Item(cdbl(rsCalCam("rubro"))) + Cdbl(dNeto)
                end if
                auxMerma = auxMerma + Cdbl(rsCalCam("KgMerma"))
                rsCalCam.MoveNext()
            wend
        End if
        'Primero obtengo valores de Rubros (Camiones y/o Vagones)
        For each theKey in oDiccNetoXRubro.Keys
            auxVl = 0
            if ((Cdbl(oDiccNetoXRubro.Item(theKey)) <> 0)and(Cdbl(oDiccNeto.Item(theKey)) <> 0 )) then auxVl = GF_EDIT_DECIMALS(round(Cdbl(oDiccNetoXRubro.Item(theKey))/Cdbl(oDiccNeto.Item(theKey)),2)*100,2)
            rtrn = rtrn & auxVl & SEPERATOR_RUBROS
	    Next
        'Obtengo valores totales de la Merma (Toneladas)
        rtrn = rtrn & GF_EDIT_DECIMALS(round(Cdbl(auxMerma)/1000,2)*100,2)
    end if
    getDescargaCalidad = rtrn
End Function 
'--------------------------------------------------------------------------------------------------------------
'Obtengo la cantidad de camiones que se encuentran en circuito, solo busca en la tabla CAMIONES (diaria) debido a que el camion 
'pasa al historico cuando este cambia a un estado terminal. Los estados terminales son obtienen de ESTADOSTERMINALES y solo obtengo el circuito Descarga
'Se tomo en cuenta que puede haber camiones que ingresaron el dia de ayer y todavia estan en circuito (DTINGRESO)
Function getDescargaCamionesCirculacion(pStrPuerto,pPeriodo,pFechaDesde,pTransporte,pFechaHasta)
    Dim strSQL,rsCan,rtrn,diaHoy
    rtrn = 0
    diaHoy = Year(Now()) &"-"& GF_nDigits(Month(Now()), 2) &"-"& GF_nDigits(Day(Now()), 2)
    if ((diaHoy = pFechaDesde)and(diaHoy= pFechaHasta)) then
        if (Cdbl(pTransporte) = TIPO_TRANSPORTE_CAMION) then
            strSQL = "SELECT COUNT(*) AS CANTIDAD "&_
                    "FROM CAMIONES "&_
                    "WHERE CDTIPOCAMION = "& CIRCUITO_CAMION_DESCARGA &_
                    "  AND CDESTADO NOT IN (SELECT CDESTADO "&_
                    "                       FROM ESTADOSTERMINALES "&_
                    "                       WHERE CDTIPOCAMION = "& CIRCUITO_CAMION_DESCARGA &" ) "
            Call GF_BD_Puertos(pStrPuerto, rsCan, "OPEN", strSQL)
            If (not rsCan.Eof) then rtrn = Cdbl(rsCan("CANTIDAD"))
        end if
        if (Cdbl(pTransporte) = TIPO_TRANSPORTE_VAGON) then
            strSQL ="SELECT COUNT(*) AS CANTIDAD "&_
                    "FROM VAGONES "&_
                    "WHERE CDESTADO NOT IN (SELECT CDESTADO "&_
                    "                       FROM ESTADOSTERMINALES "&_
                    "                       WHERE CDTIPOCAMION = 99 ) "
            Call GF_BD_Puertos(pStrPuerto, rsCan, "OPEN", strSQL)
            If (not rsCan.Eof) then rtrn = Cdbl(rsCan("CANTIDAD"))
        end if
    end if
    getDescargaCamionesCirculacion  = rtrn
End Function
'--------------------------------------------------------------------------------------------------------------
'Se encarga de redondear el tiempo para abajo, esto sucede para las horas y minutos 
Function redondearHoraPromedio(ByRef p_Hora,ByRef p_Minuto)
    if (p_Hora - CInt(p_Hora) < 0) Then
        p_Hora = CInt(p_Hora) - 1
    else
        p_Hora = CInt(p_Hora)
    end if
    if (p_Minuto - CInt(p_Minuto) < 0) Then
        p_Minuto = CInt(p_Minuto) - 1
    else
        p_Minuto = CInt(p_Minuto)
    end if
End Function
'--------------------------------------------------------------------------------------------------------------
'Obtiene el tiempo promedio que un camion tarda en realizar todo el recorrido de la planta
'Se devuelve la Hora y Minuto
Function getDescargaTiempoPromedio(pFechaDesde,pFechaHasta,pTransporte,pStrPuerto)
    Dim strSQL,auxMinutos,strSQLCam,strSQLVag
    strSQLCam = ""
    strSQLVag = ""
    diaHoy = Year(Now()) & "-" & GF_nDigits(Month(Now()), 2) & "-" & GF_nDigits(Day(Now()), 2)
    if ((Cdbl(pTransporte) = TIPO_TRANSPORTE_CAMION)or(Cdbl(pTransporte) = TIPO_TRANSPORTE_CAMVAG)) then
        strSQLCam = "     ( SELECT ((Year(dtingreso) * 10000) + (Month(dtingreso) * 100) + Day(dtingreso)) AS FECHAINGRESO,"&_
                    "              ((DATEPART(HOUR, dtingreso) * 10000) + (DATEPART(MINUTE, dtingreso) * 100) + DATEPART(SECOND, dtingreso)) AS HORAINGRESO,"&_
                    "              ((Year(dtegreso) * 10000) + (Month(dtegreso) * 100) + Day(dtegreso)) AS FECHAEGRESO, "&_
                    "              ((DATEPART(HOUR, dtegreso) * 10000) + (DATEPART(MINUTE, dtegreso) * 100) + DATEPART(SECOND, dtegreso)) AS HORAEGRESO, "&_
                    "               '"& diaHoy &"' AS DTCONTABLE, "&TIPO_TRANSPORTE_CAMION&" AS TRANSPORTE "&_
	                "         FROM CAMIONES "&_
	                "         WHERE DTINGRESO IS NOT NULL AND DTEGRESO IS NOT NULL AND CDTIPOCAMION = "& CIRCUITO_CAMION_DESCARGA &" ) "&_
	                "       UNION "&_
	                "       (SELECT ((Year(dtingreso) * 10000) + (Month(dtingreso) * 100) + Day(dtingreso)) AS FECHAINGRESO,"&_
                    "               ((DATEPART(HOUR, dtingreso) * 10000) + (DATEPART(MINUTE, dtingreso) * 100) + DATEPART(SECOND, dtingreso)) AS HORAINGRESO,"&_
                    "               ((Year(dtegreso) * 10000) + (Month(dtegreso) * 100) + Day(dtegreso)) AS FECHAEGRESO, "&_
                    "               ((DATEPART(HOUR, dtegreso) * 10000) + (DATEPART(MINUTE, dtegreso) * 100) + DATEPART(SECOND, dtegreso)) AS HORAEGRESO, "&_
                    "               DTCONTABLE,"&TIPO_TRANSPORTE_CAMION&" AS TRANSPORTE "&_
	                "         FROM HCAMIONES "&_
	                "         WHERE DTINGRESO IS NOT NULL AND DTEGRESO IS NOT NULL AND CDTIPOCAMION = "& CIRCUITO_CAMION_DESCARGA &_
	  	            "            AND DTCONTABLE >= '"& pFechaDesde &"' AND DTCONTABLE <= '"& pFechaHasta &"' )"
    end if
    if ((Cdbl(pTransporte) = TIPO_TRANSPORTE_VAGON)or(Cdbl(pTransporte) = TIPO_TRANSPORTE_CAMVAG)) then
        strSQLVag = "     (SELECT ((Year(dtarribo) * 10000) + (Month(dtarribo) * 100) + Day(dtarribo)) AS FECHAINGRESO,"&_
                    "             ((DATEPART(HOUR, dtarribo) * 10000) + (DATEPART(MINUTE, dtarribo) * 100) + DATEPART(SECOND, dtarribo)) AS HORAINGRESO,"&_
                    "             ((Year(dtfin) * 10000) + (Month(dtfin) * 100) + Day(dtfin)) AS FECHAEGRESO, "&_
                    "             ((DATEPART(HOUR, dtfin) * 10000) + (DATEPART(MINUTE, dtfin) * 100) + DATEPART(SECOND, dtfin)) AS HORAEGRESO, "&_
                    "             '"& diaHoy &"' AS DTCONTABLE, "&TIPO_TRANSPORTE_VAGON&" AS TRANSPORTE "&_
                    "      FROM   vagones "&_
                    "      WHERE  dtarribo IS NOT NULL AND dtfin IS NOT NULL ) "&_
                    "      UNION "&_
                    "     (SELECT ((Year(dtarribo) * 10000) + (Month(dtarribo) * 100) + Day(dtarribo)) AS FECHAINGRESO,"&_
                    "             ((DATEPART(HOUR, dtarribo) * 10000) + (DATEPART(MINUTE, dtarribo) * 100) + DATEPART(SECOND, dtarribo)) AS HORAINGRESO,"&_
                    "             ((Year(dtfin) * 10000) + (Month(dtfin) * 100) + Day(dtfin)) AS FECHAEGRESO, "&_
                    "             ((DATEPART(HOUR, dtfin) * 10000) + (DATEPART(MINUTE, dtfin) * 100) + DATEPART(SECOND, dtfin)) AS HORAEGRESO, "&_ 
                    "             DTCONTABLE, "&TIPO_TRANSPORTE_VAGON&" AS TRANSPORTE "&_
                    "      FROM   hvagones "&_
                    "      WHERE  dtarribo IS NOT NULL AND dtfin IS NOT NULL "&_
                    "           AND dtcontable >= '"& pFechaDesde &"' AND dtcontable <= '"& pFechaHasta &"') "
    end if
    if (((Cdbl(pTransporte) = TIPO_TRANSPORTE_VAGON)and(Cdbl(pTransporte) = TIPO_TRANSPORTE_CAMION))or(Cdbl(pTransporte) = TIPO_TRANSPORTE_CAMVAG)) then strUnion = "UNION"
    
    strSQL = "SELECT T.DTCONTABLE,T.TRANSPORTE, "&_
             "       CAST(T.FECHAINGRESO as BIGINT)*1000000 + right('000000' + cast(T.HORAINGRESO AS varchar(6)), 6) AS INGRESO, "&_
             "       CAST(T.FECHAEGRESO as BIGINT)*1000000 + right('000000' + cast(T.HORAEGRESO AS varchar(6)), 6) AS EGRESO "&_
             "FROM ("& strSQLCam & strUnion & strSQLVag & ") T "&_
             "WHERE T.DTCONTABLE >= '"& pFechaDesde &"' AND T.DTCONTABLE <= '"& pFechaHasta &"'"
    Call GF_BD_Puertos(pStrPuerto, rsTiem, "OPEN", strSQL)
    totSegundo = 0
    cantidadReg = 0
    while (not rsTiem.Eof)
        If (Cdbl(rsTiem("INGRESO")) < Cdbl(rsTiem("EGRESO"))) Then totSegundo = totSegundo + Cdbl(GF_DTEDIFF(Cdbl(rsTiem("INGRESO")),Cdbl(rsTiem("EGRESO")),"S"))
        rsTiem.MoveNext()
        cantidadReg = cantidadReg + 1
    Wend
    If (Cdbl(cantidadReg) <> 0) Then
        totSegundo  = Cdbl(totSegundo)/Cdbl(cantidadReg)
        auxHora     = totSegundo/3600
        auxMinuto   = (totSegundo Mod 3600) / 60
        auxSegundo  = totSegundo Mod 60
        Call redondearHoraPromedio(auxHora,auxMinuto)
        getDescargaTiempoPromedio = GF_nDigits(auxHora,2)&":"& GF_nDigits(auxMinuto,2)
    else
        getDescargaTiempoPromedio = 0
    End if
End Function
'----------------------------------------------------------------------------------------------------------------------------
'Esta funcion obtiene el PROMEDIO de velocidad de descarga por hora que hay en el Puerto
'Para obtener el tiempo 
Function getDescargaVelocidad(pFechaDesde,pFechaHasta,pTipoTransporte,pPto)
    Dim rsDesKg ,totKilos, countHoras,difHoras,auxFechaDesde,auxFechaHasta
    Set rsDesKg = getSQLDescargaKg(pFechaDesde,pFechaHasta,pTipoTransporte,pPto,True)
    auxFechaDesde = left(pFechaDesde,4) & Mid(pFechaDesde, 6, 2) & Right(pFechaDesde,2)
    auxFechaHasta = left(pFechaHasta,4) & Mid(pFechaHasta, 6, 2) & Right(pFechaHasta,2)
    difHoras = GF_DTEDIFF(auxFechaDesde,auxFechaHasta,"H") + 24
    rtrn = 0
    if (not rsDesKg.Eof) then
        while (not rsDesKg.Eof) 
            totKilos = Cdbl(totKilos) + Cdbl(rsDesKg("KGDESCARGA"))
            rsDesKg.MoveNext()
        wend
        totKilos = Cdbl(totKilos)/1000
        rtrn = GF_EDIT_DECIMALS((Cdbl(totKilos)/Cdbl(difHoras))*100,2)
    end if
    getDescargaVelocidad = rtrn
End Function
'--------------------------------------------------------------------------------------------------------------------------------
Function obtenerProductosDescargados(pFechaDesde,pFechaHasta,pTransporte,pPto)
    Dim strSQL,diaHoy
    diaHoy = Year(Now()) & "-" & GF_nDigits(Month(Now()), 2) & "-" & GF_nDigits(Day(Now()), 2)
    if ((Cdbl(pTransporte) = TIPO_TRANSPORTE_CAMION)or(Cdbl(pTransporte) = TIPO_TRANSPORTE_CAMVAG)) then
        strSQLCam = "  (SELECT C.CDPRODUCTO "&_
                    "  FROM (( SELECT CDPRODUCTO "&_
	                "           FROM CAMIONES "&_
	                "           WHERE CDPRODUCTO IN (SELECT DISTINCT(CDPRODUCTO) "&_
			        "		                          FROM CAMIONES "&_
                    "                                WHERE '"& diaHoy &"' >= '"& pFechaDesde &"' AND '"& diaHoy &"' <= '"& pFechaHasta &"' ) "&_
	                "           GROUP BY CDPRODUCTO "&_
	                "  )UNION( "&_
	                "          SELECT CDPRODUCTO "&_
	                "          FROM HCAMIONES "&_
	                "          WHERE CDPRODUCTO IN (SELECT DISTINCT(CDPRODUCTO)  "&_
			        "	 	                         FROM HCAMIONES "&_
			        "	 	                         WHERE DTCONTABLE >= '"& pFechaDesde &"'  AND DTCONTABLE <= '"& pFechaHasta &"') "&_
	                "          GROUP BY CDPRODUCTO "&_
	                "   )) AS C) "
    end if
    if ((Cdbl(pTransporte) = TIPO_TRANSPORTE_VAGON)or(Cdbl(pTransporte) = TIPO_TRANSPORTE_CAMVAG)) then
        strSQLVag = "  (SELECT V.CDPRODUCTO "&_
                    "  FROM (( SELECT CDPRODUCTO "&_
	                "           FROM VAGONES "&_
	                "           WHERE CDPRODUCTO IN (SELECT DISTINCT(CDPRODUCTO) "&_
			        "		                          FROM VAGONES "&_
                    "                                WHERE '"& diaHoy &"' >= '"& pFechaDesde &"' AND '"& diaHoy &"' <= '"& pFechaHasta &"' ) "&_
	                "           GROUP BY CDPRODUCTO "&_
	                "  )UNION( "&_
	                "          SELECT CDPRODUCTO "&_
	                "          FROM HVAGONES "&_
	                "          WHERE CDPRODUCTO IN (SELECT DISTINCT(CDPRODUCTO)  "&_
			        "	 	                         FROM HVAGONES "&_
			        "	 	                         WHERE DTCONTABLE >= '"& pFechaDesde &"'  AND DTCONTABLE <= '"& pFechaHasta &"') "&_
	                "          GROUP BY CDPRODUCTO "&_
	                "   )) AS V) "
    end if
    if (((Cdbl(pTransporte) = TIPO_TRANSPORTE_VAGON)and(Cdbl(pTransporte) = TIPO_TRANSPORTE_CAMION))or(Cdbl(pTransporte) = TIPO_TRANSPORTE_CAMVAG)) then strUnion = "UNION"
    strSQL = "SELECT * FROM ("& strSQLCam & strUnion & strSQLVag & ") T "
    Call GF_BD_Puertos(pPto, rs, "OPEN", strSQL)
    Set obtenerProductosDescargados = rs
End Function
'--------------------------------------------------------------------------------------------------------------------------------
'Genera e imprime el combo box con los productos descargados en el periodo establecido, por defecto se ordenara por cantidad de descargas que tuvo
'Devuelve un producto por default para trabajarlo en el resto de las consultas
Function generateSelectedProducts(pFechaDesde, pFechaHasta,pTransporte,pPto)
    Dim strCombo,auxProducto,rsProductos,listProducto,rsProExp, dsPro
    Set rsProductos = obtenerProductosDescargados(pFechaDesde,pFechaHasta,pTransporte,pPto)
    listProducto = 0
    if (not rsProductos.Eof) then
        while (not rsProductos.Eof)
            listProducto = listProducto & rsProductos("CDPRODUCTO") & ","
            rsProductos.MoveNext()
        Wend
        listProducto = left(listProducto,len(listProducto)-1)
    end if
    auxProducto = 0
    strCombo = "<select style='width:120px;height:16px;' id='cmbProducto' id='cmbProducto' onchange='changePrducto(this);'>"
    Call executeQuery(rsProExp, "OPEN", "SELECT CODIPR,INGLPR, DESCPR FROM MERFL.MER112F1 WHERE CODIPR IN ("& listProducto &")")
    if (not rsProExp.Eof) then
        auxProducto = Cdbl(rsProExp("CODIPR"))
        while (not rsProExp.Eof)
            if((Trim(rsProExp("INGLPR")) <> "") and (GF_GET_IDIOMA() = LANG_ENGLISH)) then 
                dsPro = Trim(rsProExp("INGLPR"))
            else
                dsPro = Trim(rsProExp("DESCPR"))
            end if
            strCombo =  strCombo & "<option value="& rsProExp("CODIPR") &">"& rsProExp("CODIPR") &"-"& dsPro &"</option>"
            rsProExp.MoveNext()
        wend
    else
        strCombo =  strCombo & "<option value=0>Sin datos</option>"
    end if
    strCombo =  strCombo & "</select>"&_
                "<input type='hidden' id='productoOld' name='productoOld' value="& auxProducto &">"
    Response.Write strCombo
    generateSelectedProducts = auxProducto
End function
'--------------------------------------------------------------------------------------------------------------------------------
Function getDsProductoExpByList(pListProducto)
    Dim rsPro
    getDsProductoExp = ""
    Call executeQuery(rsPro, "OPEN", "SELECT INGLPR FROM MERFL.MER112F1 WHERE CODIPR IN ("& pListProducto &")" )
    if (not rsPro.Eof) then getDsProductoExp = Trim(rsPro("INGLPR"))
End Function 
'*************************************************************************************
'***************************** COMIENZO DE LA PAGINA ***********************************
'*************************************************************************************
Dim g_strPuerto,fechaDesde,fechaHasta,listAll,unidadPeso,g_auxRubrosHumedad,g_auxRubrosPH,cdProducto,seccion
   
g_strPuerto = GF_Parametros7("pto","",6)
fechaDesde = GF_PARAMETROS7("fechaDesde", "", 6)
fechaHasta = GF_PARAMETROS7("fechaHasta", "", 6)
listAll = GF_PARAMETROS7("listAll", "", 6)
cdProducto = GF_PARAMETROS7("cdProducto", 0, 6)
transporte = GF_PARAMETROS7("transporte", 0, 6)
seccion = GF_PARAMETROS7("seccion", 0, 6)

' fechaDesde : viene en formato AAAAMMDD
' fechaHasta : viene en formato AAAAMMDD

fechaDesde = GF_FN2DTCONTABLE(fechaDesde)
fechaHasta = GF_FN2DTCONTABLE(fechaHasta)


'Si el producto es 0 significa que debe buscar todos los datos 
if (CInt(cdProducto) = 0 ) then
    if (CInt(seccion) = SECCION_TN_TOTAL) then strValue = getDescargaKg(fechaDesde,fechaHasta,transporte,g_strPuerto)
    
    if (CInt(seccion) = SECCION_VELOCIDAD) then strValue = getDescargaVelocidad(fechaDesde,fechaHasta,transporte,g_strPuerto)

    if (CInt(seccion) = SECCION_TIEMPO_PROMEDIO) then strValue = getDescargaTiempoPromedio(fechaDesde,fechaHasta,transporte,g_strPuerto)
    
    if (CInt(seccion) = SECCION_CAMION_PUERTO) then strValue = getDescargaCamionesCirculacion(g_strPuerto,periodo,fechaDesde,transporte,fechaHasta)
    
    if (CInt(seccion) = SECCION_CALIDAD) then 
        cdProducto = generateSelectedProducts(fechaDesde,fechaHasta,transporte,g_strPuerto)
        strValue = STRING_DELIMITER & getDescargaCalidad(fechaDesde,fechaHasta,cdProducto,transporte,g_strPuerto)
    end if

else
    'Si el producto tiene dato solo actualizamos la calidad
    strValue = getDescargaCalidad(fechaDesde,fechaHasta,cdProducto,transporte,g_strPuerto)
end if

Response.Write strValue
%>

