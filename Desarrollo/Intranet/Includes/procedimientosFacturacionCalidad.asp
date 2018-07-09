<%

CONST PARAM_CD_RUBRO_HUMEDAD = "CDRUBROHUMEDAD"
CONST PARAM_CD_RUBRO_ZARANDA = "CDRUBROZARANDA"
Const PARAM_FACT_FECHA_PROPIAS = "DTULTFACTACONDP"
Const PARAM_FACT_FECHA_3ROS = "DTULTFACTACONDE"
Const PARAM_FACT_CIA = "FACTURA_CIA"

'-- Constantes para estados en el proceso de facturacion de la calidad
Const FACTURA_CALIDAD_PROFORMA_PTO	= 1 'Camion descargado en puerto y Orden emitida en planta
Const FACTURA_CALIDAD_PROFORMA_BSAS = 2 'Orden transferida a Bs As.
Const FACTURA_CALIDAD_FACTURADA    	= 3 'FAC/NCR emitida OK.
Const FACTURA_CALIDAD_PRE_CANCELADA	= 8 'Orden facturación anulada por el puerto pero ya transmitida a Bs As. Queda pendiente ver su estado en la sincronización a Bs As.
Const FACTURA_CALIDAD_CANCELADA    	= 9 'Orden facturación anulada por el puerto.

Const FACT_ACOND_DESCARGA_PROPIAS = "P"
Const FACT_ACOND_DESCARGA_3ROS = "E"

Const RUBRO_EXCLUSIVO_ZARANDA = 22

Const SERVICIO_ACOND_SECADO = 21
Const SERVICIO_ACOND_SECADO_D = "SERVICIO DE SECADO"
Const SERVICIO_ACOND_ZARANDA = 22
Const SERVICIO_ACOND_ZARANDA_D = "SERVICIO DE ZARANDEO"

Const TIPO_CBTE_EMITIDO_FAC = 1
Const TIPO_CBTE_EMITIDO_NDB = 2
Const TIPO_CBTE_EMITIDO_NCR = 3

Dim gRubrosHumedad, gRubrosZaranda, gRubrosFacturacion, gCia
'-----------------------------------------------------------------------------------------------------------------
Function cargarValoresGlobalesFAC(pPto)
    
    Dim strSQL, rs, idx, ret    	
	cargarValoresGlobalesFAC = false
	gCia = getValueParametro(PARAM_FACT_CIA, pPto)		
	if (gCia <> "") then
		'Preparo los rubros de humedad
		gRubrosHumedad = getValueParametro(PARAM_CD_RUBRO_HUMEDAD,pPto)
		gRubrosFacturacion = gRubrosHumedad		
		'Preparo los rubros de Zaranda
		gRubrosZaranda = getValueParametro(PARAM_CD_RUBRO_ZARANDA,pPto)
		gRubrosFacturacion = gRubrosFacturacion & "," & gRubrosZaranda		
		cargarValoresGlobalesFAC = true		
	end if
		
End Function
'-----------------------------------------------------------------------------------------------------------------
Function armarSQLRubrosCamiones(p_ctaPorte, p_idTransporte, p_producto, p_mmtoDesde, p_mmtoHasta, p_cliente, p_Ordenar, p_excludeADM)

    Dim myDtDesde, mtDtHasta, strSQL            
     
    myDtDesde = Left(p_mmtoDesde, 4) & "-" & mid(p_mmtoDesde, 5, 2) & "-" & Right(p_mmtoDesde, 2)	 
    myDtHasta = Left(p_mmtoHasta, 4) & "-" & mid(p_mmtoHasta, 5, 2) & "-" & Right(p_mmtoHasta, 2)	 
						
    strSQL ="SELECT (YEAR(HC.dtcontable)*10000 + Month(HC.dtcontable)*100 + DAY(HC.dtcontable))  AS FECHA, "&_
			"HC.idcamion IDTRANSPORTE, "&_
			"HC.cdproducto AS PRODUCTO, "&_
			"HCD.nucartaporte AS CPORTE, "&_
			"CL.nucuit AS CUITCLIENTE, "&_
            "CL.dscliente , "&_
			"( (SELECT PC.vlpesada "&_
			   "FROM   hpesadascamion PC "&_
			   "WHERE  PC.dtcontable = HCD.dtcontable "&_
			          "AND PC.idcamion = HCD.idcamion "&_
			          "AND PC.cdpesada = 1 "&_
			          "AND PC.sqpesada = (SELECT Max(sqpesada) "&_
			                             "FROM   hpesadascamion "&_
			                             "WHERE  PC.dtcontable = dtcontable "&_
			                                    "AND PC.idcamion = idcamion "&_
			                                    "AND cdpesada = 1)) - "&_
			    "(SELECT PC.vlpesada "&_
				   " FROM   hpesadascamion PC "&_
			    " WHERE  PC.dtcontable = HCD.dtcontable "&_
			    "        AND PC.idcamion = HCD.idcamion "&_
			    "        AND PC.cdpesada = 2 "&_
			    "        AND PC.sqpesada = (SELECT Max(sqpesada) "&_
			    "                           FROM   hpesadascamion "&_
			    "                           WHERE  PC.dtcontable = dtcontable "&_
			    "                                  AND PC.idcamion = idcamion "&_
			    "                                  AND cdpesada = 2)) ) "&_
			"KILOSNETOS, "&_
			"(SELECT CASE "&_
			"          WHEN HMC.vlmermakilos IS NULL THEN 0 "&_
			"          ELSE HMC.vlmermakilos "&_
			"        END "&_
			" FROM   hmermascamiones HMC "&_
			" WHERE  HMC.dtcontable = HCD.dtcontable "&_
			"        AND HMC.idcamion = HCD.idcamion "&_
			"        AND HMC.sqpesada = (SELECT Max(sqpesada) "&_
			"                            FROM   hmermascamiones "&_
			"                            WHERE  dtcontable = HCD.dtcontable "&_
			"                                   AND idcamion = HCD.idcamion)) "&_
			"KILOSMERMA, "&_
			"HRVC.cdrubro AS RUBRO, "&_
			"R.dsrubro AS DESCRUBRO, "&_
			"HRVC.vlMerma AS PORCMERMARUBRO, "&_
			"CASE WHEN HRVC.vlbonrebaja IS NOT NULL THEN HRVC.vlbonrebaja ELSE 0 END AS VLANALISIS, "&_			
			"CASE WHEN MXS.CDPUNTO IS NOT NULL THEN MXS.CDPUNTO ELSE 0 END AS PUNTO, "&_
            "CASE WHEN HRVC.CDRUBRO in (" & gRubrosHumedad & ") THEN "&_
			"       (Select TOP 1 CDMONEDA from PRECIOSERVICIOS where CDCONCEPTO= " & SERVICIO_ACOND_SECADO & " and  PTODESDE<=MXS.CDPUNTO and PTOHASTA>=MXS.CDPUNTO and VIGENCIADESDE <= HC.dtcontable order by VIGENCIADESDE desc)"&_
			"     WHEN HRVC.CDRUBRO in (" & gRubrosZaranda & ") THEN"&_
			"       (Select TOP 1 CDMONEDA from PRECIOSERVICIOS where CDCONCEPTO= " & SERVICIO_ACOND_ZARANDA & " and VIGENCIADESDE <= HC.dtcontable order by VIGENCIADESDE desc)"&_
			"     ELSE " & MONEDA_DOLAR_NUMERICO &_
			" END MONEDAGASTO, "&_			
			"CASE WHEN HRVC.CDRUBRO in (" & gRubrosHumedad & ") and MXS.CDPUNTO IS NOT NULL THEN "&_
	        "           (Select TOP 1 (PRECIO+(PRECIOADICIONAL*(MXS.CDPUNTO-PTODESDE))) from PRECIOSERVICIOS where CDCONCEPTO= " & SERVICIO_ACOND_SECADO & " and  PTODESDE<=MXS.CDPUNTO and PTOHASTA>=MXS.CDPUNTO and VIGENCIADESDE <= HC.dtcontable order by VIGENCIADESDE desc) "&_
			"     WHEN HRVC.CDRUBRO in (" & gRubrosZaranda & ", " & RUBRO_EXCLUSIVO_ZARANDA & ") THEN"&_
			"       (Select TOP 1 PRECIO from PRECIOSERVICIOS where CDCONCEPTO= " & SERVICIO_ACOND_ZARANDA & " and VIGENCIADESDE <= HC.dtcontable order by VIGENCIADESDE desc)"&_
			"     ELSE 0 "&_
			"END IMPORTEGASTO "&_
		"FROM   (SELECT * FROM hcamiones WHERE dtcontable >= '"& myDtDesde &"' and dtcontable <= '"& myDtHasta &"' and cdestado IN ( 6, 8 ) "
		if (p_producto <> 0) then strSQL= strSQL & " and CDPRODUCTO=" & p_producto
		strSQL= strSQL & ") HC INNER JOIN (Select * from hcamionesdescarga WHERE 1=1"
		if (p_excludeADM) then	strSQL= strSQL & " and CDVENDEDOR <> 7431" 'JAS: CAMBIAR!!!!
        if (p_ctaPorte <> "") then strSQL= strSQL & " AND NUCARTAPORTE LIKE '%" & p_ctaPorte & "%'"
        if (p_idTransporte <> "") then strSQL= strSQL & " AND IDCAMION = '" & p_idTransporte & "'"
        if (p_cliente = FACT_ACOND_DESCARGA_PROPIAS) then strSQL= strSQL & " AND CDCLIENTE = 1"
        if (p_cliente = FACT_ACOND_DESCARGA_3ROS) then strSQL= strSQL & " AND CDCLIENTE <> 1"
        strSQL= strSQL & ") HCD "&_
        "       ON HC.idcamion = HCD.idcamion "&_
        "          AND HC.dtcontable = HCD.dtcontable "&_        
        "INNER JOIN ("       
        '-------------------------------------------------------------------------------------------------------
        'Esta porcion de SQL sirve para obtener los rubros de analisis que interesan para la facturacion evitando que se superpongan las cobranzas
        'Si existen rubros generales de zarandeo defibidos y se cargaron en el camion, se ignora el rubro especial de ZARANDEO ya que el servicio se cobra
        'directamente por el rubro general.
        'La logica es a una tabla con todos los reubros de facturacion le saco los registros del rubro especial de ZARANDA salvo para aquellos camiones que no tienen 
        strSQL = strSQL &   "Select * from " &_
                            " ("
        'Tabla con todos rubros de facturacion
	    strSQL = strSQL &   "   (" &_
		                    "       SELECT * " &_
		                    "       FROM   (Select * from hrubrosvisteocamiones where DTCONTABLE>='"& myDtDesde &"' and dtcontable <= '" & myDtHasta & "') A " &_
		                    "       WHERE  A.sqcalada = (SELECT Max(sqcalada) " &_
				            "		                FROM   HCALADADECAMIONES" &_
	                        "                            WHERE  idcamion = A.idcamion" &_
	                        "                                  AND dtcontable = A.dtcontable)" &_
	                        "         AND ((A.CDRUBRO in (" & gRubrosFacturacion & ") and VLMERMA > 0) or A.CDRUBRO=" & RUBRO_EXCLUSIVO_ZARANDA & ")" &_
  	                        "    ) " &_
                            " EXCEPT" &_
	                        "   (Select Z.* FROM " &_
		                    "       ("
		'HZ: Tabla con todos rubros de Generales de Zarandeo	
		strSQL = strSQL &   "           SELECT A.* FROM " &_
			                "               (Select * from hrubrosvisteocamiones where DTCONTABLE>='"& myDtDesde &"' and dtcontable <= '" & myDtHasta & "') A " &_
			                "           WHERE  A.sqcalada = (SELECT Max(sqcalada) " &_
							"                                 FROM   HCALADADECAMIONES " &_
	                        "                                 WHERE  idcamion = A.idcamion " &_
	                        "                                 AND dtcontable = A.dtcontable) " &_
	      		            "                  AND A.CDRUBRO in (" & gRubrosZaranda & ") and VLMERMA > 0" &_
	                        "       ) HZ " &_
	                        "       inner join "
	    'Z: Tabla con todas las que tienen rubro ZARANDA
	    strSQL = strSQL &   "       (" &_
	      	                "           SELECT * " &_
			                "           FROM   (Select * from hrubrosvisteocamiones where DTCONTABLE>='"& myDtDesde &"' and dtcontable <= '" & myDtHasta & "') A " &_
			                "           WHERE  A.sqcalada = (SELECT Max(sqcalada) " &_
							"                                FROM   HCALADADECAMIONES" &_
		                    "                                WHERE  idcamion = A.idcamion" &_
		                    "                                AND dtcontable = A.dtcontable)" &_
		      	            "            AND A.CDRUBRO in (" & RUBRO_EXCLUSIVO_ZARANDA & ") " &_
	                        "       ) Z " &_
	                        "       on HZ.DTCONTABLE=Z.DTCONTABLE and HZ.IDCAMION=Z.IDCAMION" &_
	                        "   ) " &_
                            " ) T " &_        
        "          ) HRVC "&_
        "       ON HC.idcamion = HRVC.idcamion "&_
        "          AND HC.dtcontable = HRVC.dtcontable "&_
        "LEFT JOIN MERMAXSECADO MXS ON MXS.CDPRODUCTO=HC.CDPRODUCTO and MXS.VLHUMEDAD=HRVC.VLBONREBAJA "&_
        "LEFT JOIN clientes CL "&_
        "      ON CL.cdcliente = HCD.cdcliente "&_               
		"LEFT JOIN rubros R "&_
        "      ON R.cdrubro = HRVC.cdrubro "		
		if (p_Ordenar) then strSQL = strSQL & "ORDER  BY HC.dtcontable, HCD.nucartaporte, HC.idcamion, HRVC.cdrubro "
        'Response.Write strSQL & "<BR>"
        'Response.End
	armarSQLRubrosCamiones = strSQL									
		
End Function
'-----------------------------------------------------------------------------------------------------------------
Function armarSQLRubrosVagones(p_ctaPorte, p_idTransporte, p_producto, p_mmtoDesde, p_mmtoHasta, p_cliente, p_Ordenar)
	dim strSQL
	strSQL  = armarSQLRubrosVagonesH(p_ctaPorte, p_idTransporte, p_producto, p_mmtoDesde, p_mmtoHasta, p_cliente, False)
	strSQL  = strSQL  & " UNION "
	strSQL  = strSQL  & armarSQLRubrosVagonesD(p_ctaPorte, p_idTransporte, p_producto, p_mmtoDesde, p_mmtoHasta, p_cliente, False)
	'response.write strSQL
	armarSQLRubrosVagones  =strSQL
End Function
'-----------------------------------------------------------------------------------------------------------------
Function armarSQLRubrosVagonesH(p_ctaPorte, p_idTransporte, p_producto, p_mmtoDesde, p_mmtoHasta, p_cliente, p_Ordenar)
	 Dim myDtDesde, mtDtHasta, strSQL, myDtHastaVisteo
	           
     
    myDtDesde = Left(p_mmtoDesde, 4) & "-" & mid(p_mmtoDesde, 5, 2) & "-" & Right(p_mmtoDesde, 2)	 
    myDtHasta = Left(p_mmtoHasta, 4) & "-" & mid(p_mmtoHasta, 5, 2) & "-" & Right(p_mmtoHasta, 2)	     
    'Como el visteo se almacena con la dtcontable en que termino todo el opeativo, pero el vagon se ve siempre con la fecha real de su descarga, puede pasar que descargo un dia y el operativo termine al siguiente
    myDtHastaVisteo = GF_DTEADD(p_mmtoHasta, 2,"D")
    
    
strSQL = "SELECT (YEAR(HC.dtcontablevagon)*10000 + Month(HC.dtcontablevagon)*100 + DAY(HC.dtcontablevagon))  AS FECHA, "&_
         "		 HC.CDOPERATIVOSERIE, "&_
		 "		 HC.CDOPERATIVO, "&_
		 "       HC.cdvagon IDTRANSPORTE, "&_
		 "		 HC.cdproducto AS PRODUCTO, "&_
		 "		 HC.nucartaporteserie + LEFT(HC.nucartaporte, 8) AS CPORTE, "&_
		 "		 CL.nucuit AS CUITCLIENTE, "&_
		 "       CL.dscliente , "&_
		 "		 ((SELECT PC.vlpesada "&_
		 "		   FROM   hpesadasvagon PC "&_
		 "		   WHERE  PC.dtcontable = HC.dtcontable "&_
		 "				AND PC.nucartaporte = HC.nucartaporte "&_
		 "				AND PC.cdvagon = HC.cdvagon "&_
		 "				AND PC.cdpesada = 1 "&_
		 "				AND PC.sqpesada = (SELECT Max(sqpesada) "&_
		 "								   FROM   hpesadasvagon "&_
		 "								   WHERE  PC.dtcontable = dtcontable "&_
		 "										AND PC.nucartaporte = nucartaporte "&_
		 "										AND PC.cdvagon = cdvagon "&_
		 "										AND cdpesada = 1)) - "&_
	     "		   (SELECT PC.vlpesada "&_
		 "			FROM   hpesadasvagon PC "&_
		 "			WHERE  PC.dtcontable = HC.dtcontable "&_
		 "				AND PC.nucartaporte = HC.nucartaporte "&_
		 "				AND PC.cdvagon = HC.cdvagon "&_
		 "				AND PC.cdpesada = 2 "&_
		 "				AND PC.sqpesada = (SELECT Max(sqpesada) "&_
		 "								   FROM   hpesadasvagon "&_
		 "								   WHERE  PC.dtcontable = dtcontable "&_
		 "										AND PC.nucartaporte = nucartaporte "&_
		 "										AND PC.cdvagon = cdvagon "&_
		 "										AND cdpesada = 2)) ) "&_
		 "			KILOSNETOS, "&_
		 "			(SELECT CASE WHEN HMC.vlmermakilos IS NULL THEN 0 ELSE HMC.vlmermakilos END "&_			
		 "			 FROM   hmermasvagones HMC "&_
		 "			 WHERE  HMC.dtcontable = HC.dtcontable "&_
		 "				AND HMC.nucartaporte = HC.nucartaporte "&_
		 "				AND HMC.cdvagon = HC.cdvagon "&_
		 "				AND HMC.sqpesada = (SELECT Max(sqpesada) "&_
		 "									FROM   hmermasvagones "&_
		 "									WHERE  dtcontable = HMC.dtcontable "&_
		 "										AND nucartaporte = HMC.nucartaporte "&_
		 "										AND cdvagon = HMC.cdvagon)) KILOSMERMA, "&_		 
		 "			HRVC.cdrubro AS RUBRO, "&_
		 "			R.dsrubro AS DESCRUBRO, "&_		
		 "          HRVC.vlMerma AS PORCMERMARUBRO, "&_	
		 "          CASE WHEN HRVC.vlbonrebaja IS NOT NULL THEN HRVC.vlbonrebaja ELSE 0 END AS VLANALISIS, "&_
	     "          CASE WHEN MXS.CDPUNTO IS NOT NULL THEN MXS.CDPUNTO ELSE 0 END AS PUNTO, "&_	
	     "          CASE WHEN HRVC.CDRUBRO in (" & gRubrosHumedad & ") THEN "&_
		 "                  (Select TOP 1 CDMONEDA from PRECIOSERVICIOS where CDCONCEPTO= " & SERVICIO_ACOND_SECADO & " and  PTODESDE<=MXS.CDPUNTO and PTOHASTA>=MXS.CDPUNTO and VIGENCIADESDE <= HC.dtcontablevagon order by VIGENCIADESDE desc)"&_
		 "               WHEN HRVC.CDRUBRO in (" & gRubrosZaranda & ") THEN"&_
		 "                  (Select TOP 1 CDMONEDA from PRECIOSERVICIOS where CDCONCEPTO= " & SERVICIO_ACOND_ZARANDA & " and  VIGENCIADESDE <= HC.dtcontablevagon order by VIGENCIADESDE desc)"&_
		 "               ELSE " & MONEDA_DOLAR_NUMERICO &_
		 "          END MONEDAGASTO, "&_			
		 "          CASE WHEN HRVC.CDRUBRO in (" & gRubrosHumedad & ") and MXS.CDPUNTO IS NOT NULL THEN "&_
	     "                      (Select TOP 1 (PRECIO+(PRECIOADICIONAL*(MXS.CDPUNTO-PTODESDE))) from PRECIOSERVICIOS where CDCONCEPTO= " & SERVICIO_ACOND_SECADO & " and  PTODESDE<=MXS.CDPUNTO and PTOHASTA>=MXS.CDPUNTO and  VIGENCIADESDE <= HC.dtcontablevagon order by VIGENCIADESDE desc) "&_
		 "              WHEN HRVC.CDRUBRO in (" & gRubrosZaranda & ") THEN"&_
		 "                  (Select TOP 1 PRECIO from PRECIOSERVICIOS where CDCONCEPTO= " & SERVICIO_ACOND_ZARANDA & " and  VIGENCIADESDE <= HC.dtcontablevagon order by VIGENCIADESDE desc)"&_
		 "              ELSE 0"&_
		 "          END IMPORTEGASTO "&_ 
		 "FROM   (SELECT * FROM hvagones WHERE dtcontablevagon >= '"& myDtDesde &"' and dtcontablevagon <= '"& myDtHasta &"' and cdestado IN ( 6, 8 )"
		 if (p_producto <> 0) then strSQL= strSQL & " and CDPRODUCTO=" & p_producto
		 if (p_idTransporte <> "") then strSQL= strSQL & " AND CDVAGON = '" & p_idTransporte & "'"
strSQL = strSQL & ") HC " &_
         "	INNER JOIN (Select * from hoperativos where 1=1"
         if (p_ctaPorte <> "") then 
            if (Len(p_ctaPorte) > 8) then
                'Ingreso serie y nro.
                strSQL= strSQL & " AND NUCARTAPORTE LIKE '%" & Right(p_ctaPorte, 8) & "%' and NUCARTAPORTESERIE LIKE '%" & Left(p_ctaPorte, Len(p_ctaPorte)-8) & "%'"
            else
                'Ingreso solo todo o parte del numero
                strSQL= strSQL & " AND NUCARTAPORTE LIKE '%" & p_ctaPorte & "%'"
            end if
         end if
         if (p_cliente = FACT_ACOND_DESCARGA_PROPIAS) then strSQL= strSQL & " AND CDCLIENTE = 1"
         if (p_cliente = FACT_ACOND_DESCARGA_3ROS) then strSQL= strSQL & " AND CDCLIENTE <> 1"
strSQL = strSQL & ") HO " &_
         "       ON HC.nucartaporte = HO.nucartaporte AND HC.dtcontable = HO.dtcontable "&_
         "  INNER JOIN ("
         '-------------------------------------------------------------------------------------------------------
        'Esta porcion de SQL sirve para obtener los rubros de analisis que interesan para la facturacion evitando que se superpongan las cobranzas
        'Si existen rubros generales de zarandeo defibidos y se cargaron en el camion, se ignora el rubro especial de ZARANDEO ya que el servicio se cobra
        'directamente por el rubro general.
        'La logica es a una tabla con todos los reubros de facturacion le saco los registros del rubro especial de ZARANDA salvo para aquellos camiones que no tienen 
        strSQL = strSQL &   "Select * from " &_
                            " ("
        'Tabla con todos rubros de facturacion
	    strSQL = strSQL &   "   (" &_
		                    "       SELECT * " &_
		                    "       FROM   (Select * from hrubrosvisteovagones where DTCONTABLE>='"& myDtDesde &"' and dtcontable <= '" & myDtHastaVisteo & "') A " &_
		                    "       WHERE  A.sqcalada = (SELECT Max(sqcalada) " &_
				            "		                     FROM   HCALADADEVAGONES" &_
	                        "                            WHERE  nucartaporte = A.nucartaporte " &_
	                        "                                   AND cdvagon = A.cdvagon" &_
	                        "                                   AND dtcontable = A.dtcontable)" &_
	                        "         AND ((A.CDRUBRO in (" & gRubrosFacturacion & ") and VLMERMA > 0) or A.CDRUBRO=" & RUBRO_EXCLUSIVO_ZARANDA & ")" &_
  	                        "    ) " &_
                            " EXCEPT" &_
	                        "   (Select Z.* FROM " &_
		                    "       ("
		'HZ: Tabla con todos rubros de Generales de Zarandeo	
		strSQL = strSQL &   "           SELECT A.* FROM " &_
			                "               (Select * from hrubrosvisteovagones where DTCONTABLE>='"& myDtDesde &"' and dtcontable <= '" & myDtHastaVisteo & "') A " &_
			                "           WHERE  A.sqcalada = (SELECT Max(sqcalada) " &_
							"                                 FROM   HCALADADEVAGONES " &_
	                        "                                 WHERE  nucartaporte = A.nucartaporte " &_
	                        "                                        AND cdvagon = A.cdvagon" &_
	                        "                                        AND dtcontable = A.dtcontable)" &_
	      		            "                  AND A.CDRUBRO in (" & gRubrosZaranda & ") and VLMERMA > 0" &_
	                        "       ) HZ " &_
	                        "       inner join "
	    'Z: Tabla con todas las que tienen rubro ZARANDA
	    strSQL = strSQL &   "       (" &_
	      	                "           SELECT * " &_
			                "           FROM   (Select * from hrubrosvisteovagones where DTCONTABLE>='"& myDtDesde &"' and dtcontable <= '" & myDtHastaVisteo & "') A " &_
			                "           WHERE  A.sqcalada = (SELECT Max(sqcalada) " &_
							"                                FROM   HCALADADEVAGONES" &_
		                    "                                WHERE  nucartaporte = A.nucartaporte " &_
	                        "                                       AND cdvagon = A.cdvagon" &_
	                        "                                       AND dtcontable = A.dtcontable)" &_
		      	            "            AND A.CDRUBRO in (" & RUBRO_EXCLUSIVO_ZARANDA & ") " &_
	                        "       ) Z " &_
	                        "       on HZ.DTCONTABLE=Z.DTCONTABLE and HZ.NUCARTAPORTE=Z.NUCARTAPORTE and HZ.CDVAGON=Z.CDVAGON" &_
	                        "   ) " &_
                            " ) T " &_        
         "      ) HRVC "&_
         "       ON HC.nucartaporte = HRVC.nucartaporte "&_
         "			AND HC.cdvagon = HRVC.cdvagon "&_
         "          AND HC.dtcontable = HRVC.dtcontable "&_
         "          AND HRVC.VLMERMA > 0 "&_  
         "          AND HRVC.CDRUBRO in (" & gRubrosFacturacion & ") "&_ 
         " LEFT JOIN clientes CL ON CL.cdcliente = HO.cdcliente "&_
         " LEFT JOIN mermaxsecado MXS on MXS.vlhumedad = HRVC.vlbonrebaja AND MXS.cdproducto = HC.cdproducto	"&_         
		 " LEFT JOIN rubros R "&_
         "      ON R.cdrubro = HRVC.cdrubro "
		 if (p_Ordenar) then strSQL = strSQL & " ORDER  BY HC.DTCONTABLE, HC.nucartaporte,HC.cdvagon,HRVC.cdrubro "
        'Response.Write strSQL & "<BR>"
	armarSQLRubrosVagonesH = strSQL
End Function
'-----------------------------------------------------------------------------------------------------------------
Function armarSQLRubrosVagonesD(p_ctaPorte, p_idTransporte, p_producto, p_mmtoDesde, p_mmtoHasta, p_cliente, p_Ordenar)
	 Dim myDtDesde, mtDtHasta, strSQL, myDtHastaVisteo
	           
     
    myDtDesde = Left(p_mmtoDesde, 4) & "-" & mid(p_mmtoDesde, 5, 2) & "-" & Right(p_mmtoDesde, 2)	 
    myDtHasta = Left(p_mmtoHasta, 4) & "-" & mid(p_mmtoHasta, 5, 2) & "-" & Right(p_mmtoHasta, 2)	     
    'Como el visteo se almacena con la dtcontable en que termino todo el opeativo, pero el vagon se ve siempre con la fecha real de su descarga, puede pasar que descargo un dia y el operativo termine al siguiente
    myDtHastaVisteo = GF_DTEADD(p_mmtoHasta, 2,"D")
    
    
strSQL = "SELECT (YEAR(HC.dtcontablevagon)*10000 + Month(HC.dtcontablevagon)*100 + DAY(HC.dtcontablevagon))  AS FECHA, "&_
         "		 HC.CDOPERATIVOSERIE, "&_
		 "		 HC.CDOPERATIVO, "&_
		 "       HC.cdvagon IDTRANSPORTE, "&_
		 "		 HC.cdproducto AS PRODUCTO, "&_
		 "		 HC.nucartaporteserie + LEFT(HC.nucartaporte, 8) AS CPORTE, "&_
		 "		 CL.nucuit AS CUITCLIENTE, "&_
		 "       CL.dscliente , "&_
		 "		 ((SELECT PC.vlpesada "&_
		 "		   FROM   pesadasvagon PC "&_
		 "		   WHERE   PC.nucartaporte = HC.nucartaporte "&_
		 "				AND PC.cdvagon = HC.cdvagon "&_
		 "				AND PC.cdpesada = 1 "&_
		 "				AND PC.sqpesada = (SELECT Max(sqpesada) "&_
		 "								   FROM   pesadasvagon "&_
		 "								   WHERE  PC.nucartaporte = nucartaporte "&_
		 "										AND PC.cdvagon = cdvagon "&_
		 "										AND cdpesada = 1)) - "&_
	     "		   (SELECT PC.vlpesada "&_
		 "			FROM   pesadasvagon PC "&_
		 "			WHERE  PC.nucartaporte = HC.nucartaporte "&_
		 "				AND PC.cdvagon = HC.cdvagon "&_
		 "				AND PC.cdpesada = 2 "&_
		 "				AND PC.sqpesada = (SELECT Max(sqpesada) "&_
		 "								   FROM   pesadasvagon "&_
		 "								   WHERE  PC.nucartaporte = nucartaporte "&_
		 "										AND PC.cdvagon = cdvagon "&_
		 "										AND cdpesada = 2)) ) "&_
		 "			KILOSNETOS, "&_
		 "			(SELECT CASE WHEN HMC.vlmermakilos IS NULL THEN 0 ELSE HMC.vlmermakilos END "&_			
		 "			 FROM   mermasvagones HMC "&_
		 "			 WHERE  HMC.nucartaporte = HC.nucartaporte "&_
		 "				AND HMC.cdvagon = HC.cdvagon "&_
		 "				AND HMC.sqpesada = (SELECT Max(sqpesada) "&_
		 "									FROM   mermasvagones "&_
		 "									WHERE  nucartaporte = HMC.nucartaporte "&_
		 "										AND cdvagon = HMC.cdvagon)) KILOSMERMA, "&_		 
		 "			HRVC.cdrubro AS RUBRO, "&_
		 "			R.dsrubro AS DESCRUBRO, "&_		
		 "          HRVC.vlMerma AS PORCMERMARUBRO, "&_	
		 "          CASE WHEN HRVC.vlbonrebaja IS NOT NULL THEN HRVC.vlbonrebaja ELSE 0 END AS VLANALISIS, "&_
	     "          CASE WHEN MXS.CDPUNTO IS NOT NULL THEN MXS.CDPUNTO ELSE 0 END AS PUNTO, "&_	
	     "          CASE WHEN HRVC.CDRUBRO in (" & gRubrosHumedad & ") THEN "&_
		 "                  (Select TOP 1 CDMONEDA from PRECIOSERVICIOS where CDCONCEPTO= " & SERVICIO_ACOND_SECADO & " and  PTODESDE<=MXS.CDPUNTO and PTOHASTA>=MXS.CDPUNTO and VIGENCIADESDE <= HC.dtcontablevagon order by VIGENCIADESDE desc)"&_
		 "               WHEN HRVC.CDRUBRO in (" & gRubrosZaranda & ") THEN"&_
		 "                  (Select TOP 1 CDMONEDA from PRECIOSERVICIOS where CDCONCEPTO= " & SERVICIO_ACOND_ZARANDA & " and  VIGENCIADESDE <= HC.dtcontablevagon order by VIGENCIADESDE desc)"&_
		 "               ELSE " & MONEDA_DOLAR_NUMERICO &_
		 "          END MONEDAGASTO, "&_			
		 "          CASE WHEN HRVC.CDRUBRO in (" & gRubrosHumedad & ") and MXS.CDPUNTO IS NOT NULL THEN "&_
	     "                      (Select TOP 1 (PRECIO+(PRECIOADICIONAL*(MXS.CDPUNTO-PTODESDE))) from PRECIOSERVICIOS where CDCONCEPTO= " & SERVICIO_ACOND_SECADO & " and  PTODESDE<=MXS.CDPUNTO and PTOHASTA>=MXS.CDPUNTO and  VIGENCIADESDE <= HC.dtcontablevagon order by VIGENCIADESDE desc) "&_
		 "              WHEN HRVC.CDRUBRO in (" & gRubrosZaranda & ") THEN"&_
		 "                  (Select TOP 1 PRECIO from PRECIOSERVICIOS where CDCONCEPTO= " & SERVICIO_ACOND_ZARANDA & " and  VIGENCIADESDE <= HC.dtcontablevagon order by VIGENCIADESDE desc)"&_
		 "              ELSE 0"&_
		 "          END IMPORTEGASTO "&_ 
		 "FROM   (SELECT * FROM vagones WHERE dtcontablevagon >= '"& myDtDesde &"' and dtcontablevagon <= '"& myDtHasta &"' and cdestado IN ( 6, 8 )"
		 if (p_producto <> 0) then strSQL= strSQL & " and CDPRODUCTO=" & p_producto
		 if (p_idTransporte <> "") then strSQL= strSQL & " AND CDVAGON = '" & p_idTransporte & "'"
strSQL = strSQL & ") HC " &_
         "	INNER JOIN (Select * from operativos where 1=1"
         if (p_ctaPorte <> "") then 
            if (Len(p_ctaPorte) > 8) then
                'Ingreso serie y nro.
                strSQL= strSQL & " AND NUCARTAPORTE LIKE '%" & Right(p_ctaPorte, 8) & "%' and NUCARTAPORTESERIE LIKE '%" & Left(p_ctaPorte, Len(p_ctaPorte)-8) & "%'"
            else
                'Ingreso solo todo o parte del numero
                strSQL= strSQL & " AND NUCARTAPORTE LIKE '%" & p_ctaPorte & "%'"
            end if
         end if
         if (p_cliente = FACT_ACOND_DESCARGA_PROPIAS) then strSQL= strSQL & " AND CDCLIENTE = 1"
         if (p_cliente = FACT_ACOND_DESCARGA_3ROS) then strSQL= strSQL & " AND CDCLIENTE <> 1"
strSQL = strSQL & ") HO " &_
         "       ON HC.nucartaporte = HO.nucartaporte "&_
         "  INNER JOIN ("
         '-------------------------------------------------------------------------------------------------------
        'Esta porcion de SQL sirve para obtener los rubros de analisis que interesan para la facturacion evitando que se superpongan las cobranzas
        'Si existen rubros generales de zarandeo defibidos y se cargaron en el camion, se ignora el rubro especial de ZARANDEO ya que el servicio se cobra
        'directamente por el rubro general.
        'La logica es a una tabla con todos los reubros de facturacion le saco los registros del rubro especial de ZARANDA salvo para aquellos camiones que no tienen 
        strSQL = strSQL &   "Select * from " &_
                            " ("
        'Tabla con todos rubros de facturacion
	    strSQL = strSQL &   "   (" &_
		                    "       SELECT * " &_
		                    "       FROM   (Select * from rubrosvisteovagones) A " &_
		                    "       WHERE  A.sqcalada = (SELECT Max(sqcalada) " &_
				            "		                     FROM   CALADADEVAGONES" &_
	                        "                            WHERE  nucartaporte = A.nucartaporte " &_
	                        "                                   AND cdvagon = A.cdvagon)" &_
	                        "         AND ((A.CDRUBRO in (" & gRubrosFacturacion & ") and VLMERMA > 0) or A.CDRUBRO=" & RUBRO_EXCLUSIVO_ZARANDA & ")" &_
  	                        "    ) " &_
                            " EXCEPT" &_
	                        "   (Select Z.* FROM " &_
		                    "       ("
		'HZ: Tabla con todos rubros de Generales de Zarandeo	
		strSQL = strSQL &   "           SELECT A.* FROM " &_
			                "               (Select * from rubrosvisteovagones) A " &_
			                "           WHERE  A.sqcalada = (SELECT Max(sqcalada) " &_
							"                                 FROM   CALADADEVAGONES " &_
	                        "                                 WHERE  nucartaporte = A.nucartaporte " &_
	                        "                                        AND cdvagon = A.cdvagon)" &_
	      		            "                  AND A.CDRUBRO in (" & gRubrosZaranda & ") and VLMERMA > 0" &_
	                        "       ) HZ " &_
	                        "       inner join "
	    'Z: Tabla con todas las que tienen rubro ZARANDA
	    strSQL = strSQL &   "       (" &_
	      	                "           SELECT * " &_
			                "           FROM   (Select * from rubrosvisteovagones ) A " &_
			                "           WHERE  A.sqcalada = (SELECT Max(sqcalada) " &_
							"                                FROM   CALADADEVAGONES" &_
		                    "                                WHERE  nucartaporte = A.nucartaporte " &_
	                        "                                       AND cdvagon = A.cdvagon)" &_
		      	            "            AND A.CDRUBRO in (" & RUBRO_EXCLUSIVO_ZARANDA & ") " &_
	                        "       ) Z " &_
	                        "       on HZ.NUCARTAPORTE=Z.NUCARTAPORTE and HZ.CDVAGON=Z.CDVAGON" &_
	                        "   ) " &_
                            " ) T " &_        
         "      ) HRVC "&_
         "       ON HC.nucartaporte = HRVC.nucartaporte "&_
         "			AND HC.cdvagon = HRVC.cdvagon "&_
         "          AND HRVC.VLMERMA > 0 "&_  
         "          AND HRVC.CDRUBRO in (" & gRubrosFacturacion & ") "&_ 
         " LEFT JOIN clientes CL ON CL.cdcliente = HO.cdcliente "&_
         " LEFT JOIN mermaxsecado MXS on MXS.vlhumedad = HRVC.vlbonrebaja AND MXS.cdproducto = HC.cdproducto	"&_         
		 " LEFT JOIN rubros R "&_
         "      ON R.cdrubro = HRVC.cdrubro "
		 if (p_Ordenar) then strSQL = strSQL & " ORDER  BY HC.nucartaporte,HC.cdvagon,HRVC.cdrubro "
        'Response.Write strSQL & "<BR>"
	armarSQLRubrosVagonesD = strSQL
End Function
'-----------------------------------------------------------------------------------------------------------------
Function migrarMermasAFacturar(p_pto, p_mmto, p_ctapte, p_idTransporte, p_transporte, p_cliente, p_logMig) 

    if (cargarValoresGlobalesFAC(p_pto)) then
		if ((p_transporte = TIPO_TRANSPORTE_CAMION) or (p_transporte = TIPO_TRANSPORTE_CAMVAG)) then Call copiarDatosCamiones(p_pto, p_mmto, p_ctapte, p_idTransporte, p_cliente, p_logMig)    		
		if ((p_transporte = TIPO_TRANSPORTE_VAGON) or (p_transporte = TIPO_TRANSPORTE_CAMVAG)) then Call copiarDatosVagones(p_pto, p_mmto, p_ctapte, p_idTransporte, p_cliente, p_logMig)		
	else	
		p_logMig.Info("NO SE PUEDE FACTURAR. ERROR AL CARGAR VALORES GLOBALES.")
	end if
	
End Function
'-----------------------------------------------------------------------------------------------------------------
Function Rubro2Concepto(pRubro, ByRef pCdConcepto, ByRef pDsConcepto)
	pCdConcepto = SERVICIO_ACOND_ZARANDA	
	pDsConcepto = SERVICIO_ACOND_ZARANDA_D
	if (InStr(1, "," & gRubrosHumedad & ",", "," & pRubro & ",") > 0) then
		pCdConcepto = SERVICIO_ACOND_SECADO
		pDsConcepto = SERVICIO_ACOND_SECADO_D
	end if
End Function
'-----------------------------------------------------------------------------------------------------------------
Function copiarDatosCamiones(p_pto, p_mmto, p_ctapte, p_idTransporte, p_cliente, p_logMig)
    Dim rs	
    'Tomo los datos a migrar y los copio en la base de Bs As.    
    Call executeQueryDb(p_pto, rs, "OPEN", armarSQLRubrosCamiones(p_ctapte, p_idTransporte, 0, p_mmto, p_mmto, p_cliente, true, true))
	Call generarOrdenesAcond(p_pto, TIPO_TRANSPORTE_CAMION, p_mmto, p_ctapte, p_idTransporte, rs, p_logMig)
	
End Function   
'--------------------------------------------------------------------------------------------------------------
Function copiarDatosVagones(p_pto, p_mmto, p_ctapte, p_idTransporte, p_cliente, p_logMig) 
	Dim rs	 
    'Tomo los datos a migrar y los copio en la base de Bs As.
	Call executeQueryDb(p_pto, rs, "OPEN", armarSQLRubrosVagones(p_ctapte, p_idTransporte, 0, p_mmto, p_mmto, p_cliente, true))
	Call generarOrdenesAcond(p_pto, TIPO_TRANSPORTE_VAGON, p_mmto, p_ctapte, p_idTransporte, rs, p_logMig)
End Function
'--------------------------------------------------------------------------------------------------------------
Function generarOrdenesAcond(p_pto, pTipoTransporte, p_mmto, p_ctapte, p_idTransporte, rs, p_logMig)

	Dim strSQLH, rsHMAF, myConcepto	
	'Leo las descargas ya registradas en la historica
	strSQLH = " Select * from (" &_
			  "Select TIPOTRANSPORTE, NUDOCUMENTO NUCARTAPORTE, IDTRANSPORTE, CDRUBRO, SUM(case when ((tipcbt = 1) or (tipcbt = 2)) then IMPORTE else -1*IMPORTE end ) IMPORTE " &_
			  " from FACTURACIONSERVICIOS " &_			  
			  " where   codcia = '"& gCia &"'" &_
			  "			and TIPOTRANSPORTE = " & pTipoTransporte	
	if (p_mmto <> "") then strSQLH = strSQLH & " AND DTCONTABLE = '"& GF_FN2DTCONTABLE(p_mmto) & "'"
	if (p_ctapte <> "") then strSQLH = strSQLH & " AND NUDOCUMENTO = '" & p_ctapte & "'"
	if (p_idTransporte <> "") then strSQLH = strSQLH & " AND IDTRANSPORTE = '" & p_idTransporte & "'"
	strSQLH = strSQLH & " group by TIPOTRANSPORTE, NUDOCUMENTO, IDTRANSPORTE, CDRUBRO ) T " &_
				" where IMPORTE > 0 " &_
				" order by NUCARTAPORTE, IDTRANSPORTE, CDRUBRO "	
    Call executeQueryDb(p_pto, rsHMAF, "OPEN", strSQLH)    
    'Copio los datos en la tabla destino.
    while ((not rs.eof) and (not rsHMAF.eof))
        if (( Cdbl(rs("CPORTE")) < CDbl(rsHMAF("NUCARTAPORTE")) )OR( Cdbl(rs("CPORTE")) = CDbl(rsHMAF("NUCARTAPORTE"))and Cdbl(rs("IDTRANSPORTE")) < Cdbl(rsHMAF("IDTRANSPORTE")))OR(Cdbl(rs("CPORTE")) = CDbl(rsHMAF("NUCARTAPORTE"))and Cdbl(rs("IDTRANSPORTE")) = Cdbl(rsHMAF("IDTRANSPORTE")) and CInt(rs("RUBRO")) < CInt(rsHMAF("CDRUBRO"))))then			
			if (CDbl(rs("IMPORTEGASTO")) > 0) then
				'Solo se puede cobrar una vez cada tipo de concepto.     			
        		Call Rubro2Concepto(rs("RUBRO"), myConcepto, myConceptoDs)		
				Call registrarOrdenAcond(p_pto, pTipoTransporte, rs("KILOSNETOS"), rs("IMPORTEGASTO"), rs("MONEDAGASTO"), rs("FECHA"), rs("CPORTE"), rs("IDTRANSPORTE"), myConcepto, myConceptoDs, rs("RUBRO"), Trim(rs("DESCRUBRO")), rs("VLANALISIS"), rs("PRODUCTO"), rs("CUITCLIENTE"), rs("PUNTO"), rs("KILOSMERMA"), p_logMig)
			end if
			'Response.Write "<BR>" & strSQL
			rs.MoveNext()
		else
			if ((Cdbl(rs("CPORTE")) = Cdbl(rsHMAF("NUCARTAPORTE"))) and (Cdbl(rsHMAF("IDTRANSPORTE")) = Cdbl(rs("IDTRANSPORTE"))) and (CInt(rs("RUBRO")) = CInt(rsHMAF("CDRUBRO")))) then rs.MoveNext()			
			rsHMAF.MoveNext()
		end if
    wend	
	'Se completan los camiones nuevos nunca antes facturados.
    while (not rs.eof)
		if (CDbl(rs("IMPORTEGASTO")) > 0) then
			Call Rubro2Concepto(rs("RUBRO"), myConcepto, myConceptoDs)		
			Call registrarOrdenAcond(p_pto, pTipoTransporte, rs("KILOSNETOS"), rs("IMPORTEGASTO"), rs("MONEDAGASTO"), rs("FECHA"), rs("CPORTE"), rs("IDTRANSPORTE"), myConcepto, myConceptoDs, rs("RUBRO"), Trim(rs("DESCRUBRO")), rs("VLANALISIS"), rs("PRODUCTO"), rs("CUITCLIENTE"), rs("PUNTO"), rs("KILOSMERMA"), p_logMig)
		end if
		'Response.Write "<BR>" & strSQL				
		rs.MoveNext()
    wend
End Function
'--------------------------------------------------------------------------------------------------------------
Function registrarOrdenAcond(p_pto, pTipoTransporte, kilosNetos, precioLiq, monedaLiq, dtFnDescarga, nuCartaPorte, pIdTransporte, pCodConc, pDescConc, pCdRubro, pDescRubro, pValorAnalisis, pCdProducto, pCuitCliente, pPtoCalidad, pKgMerma, p_logMig)
	Dim toneladasLiq, importeLiq, descServicio, strSQL, rs2
	
	toneladasLiq = CDbl(kilosNetos)/1000
	importeLiq = toneladasLiq * CDbl(precioLiq)	
	descServicio =  "Carta de Porte: " & GF_EDIT_CBTE(nuCartaPorte) & " " & pDescConc & "(" & pDescRubro & ": " & GF_EDIT_DECIMALS(CDbl(pValorAnalisis)*100, 2) & ") Tarifa: " & getSimboloMoneda(monedaLiq) & " " & GF_EDIT_DECIMALS(Cdbl(precioLiq)*100, 2) & "/Tn"		
	strSQL = "Insert into FACTURACIONSERVICIOS([tipoTransporte], [dtContable], [nudocumento], [IDTransporte], [codconce], [descripcion], [cdProducto], [cuitCliente], [cdRubro], [vlRubro], [ptoCalidad], [kilos], [merma], [codmone], [precio], [importe], [codcia], [tipcbt], [letra], [succbt], [nrocbt], [fecalta], [usualta], [estado]) " &_
			 " values(" & pTipoTransporte & ", '" & GF_FN2DTCONTABLE(dtFnDescarga) & "', '" & nuCartaPorte & "', '"& pIdTransporte &"', " & pCodConc & ", '" & Trim(descServicio) & "', " & pCdProducto & ", '" & Trim(pCuitCliente) & "', " & pCdRubro & ", " & pValorAnalisis & ", " & pPtoCalidad & "," & kilosNetos & ", " & pKgMerma & ", " & monedaLiq & ", " & precioLiq & ", " & importeLiq & ", '" & gCia & "', " & TIPO_CBTE_EMITIDO_FAC & ", '', 0, 0, '" & left(session("MmtoSistema"), 8) & "', '" & session("Usuario") & "', " & FACTURA_CALIDAD_PROFORMA_PTO &  ")"	
	'Response.Write "<BR>" & strSQL	
	Call executeQueryDb(p_pto, rs2, "EXEC", strSQL)
	p_logMig.info("CARTA PORTE:"& nuCartaPorte &"| TRANSPORTE:"&pIdTransporte&"| CONCEPTO(RUBRO): "& pCodConc & "(" & pDescRubro &")|TOTAL: "&toneladasLiq&" Tn. | IMPORTE: " & monedaLiq & " " & importeLiq)	
End Function	
'--------------------------------------------------------------------------------------------------------------
Function getDSEstadoProformaCalidad(cdEstado)
	getDSEstadoProformaCalidad = "getDSEstadoProformaCalidad - ERROR - (" & cdEstado & ")"
	Select case cdEstado
		case FACTURA_CALIDAD_PROFORMA_PTO
			getDSEstadoProformaCalidad = "Pendiente"
		case FACTURA_CALIDAD_PROFORMA_BSAS
			getDSEstadoProformaCalidad = "Proforma"
		case FACTURA_CALIDAD_FACTURADA
			getDSEstadoProformaCalidad = "Facturada"
		case FACTURA_CALIDAD_PRE_CANCELADA, FACTURA_CALIDAD_CANCELADA
			getDSEstadoProformaCalidad = "Cancelada"
	End Select
End Function	
'--------------------------------------------------------------------------------------------------------------
%>