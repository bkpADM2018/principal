<!--#include file="../../Includes/procedimientosMG.asp"-->
<!--#include file="../../Includes/procedimientostraducir.asp"-->
<!--#include file="../../Includes/procedimientosfechas.asp"-->
<!--#include file="../../Includes/procedimientosFormato.asp"-->
<!--#include file="../../Includes/procedimientosSQL.asp"-->
<!--#include file="../../Includes/procedimientosExcel.asp"-->
<!--#include file="../../Includes/procedimientosUser.asp"-->
<!--#include file="../../Includes/procedimientosPuertos.asp"-->
<!--#include file="../../Includes/procedimientos.asp"-->
<%
const OPERADOR_SUMA_KILOS = "+"
const OPERADOR_RESTA_KILOS = "-"
'------------------------------------------------------------------------------------
Function ArmarColumnaAjustes(pMyFecha,ppto,campo,pLstClientes)
Dim strSQL
strSQL = "				SELECT DISTINCT " &	campo
strSQL = strSQL & " 	FROM   dbo.excreditosdebitos cd			"		
strSQL = strSQL & " 		LEFT JOIN dbo.clientes c				"
strSQL = strSQL & " 			ON c.cdcliente = cd.cdcliente			"
strSQL = strSQL & " 		LEFT JOIN dbo.productos p				"
strSQL = strSQL & " 			ON p.cdproducto = cd.cdproducto			"
strSQL = strSQL & " 	WHERE  cd.dtcontable = '"& pMyFecha &"'			"
if (pLstClientes <> "") then strSQL = strSQL & " and cd.cdcliente in ("& pLstClientes & ")"
strSQL = strSQL & " 	ORDER  BY " & campo       
Call GF_BD_Puertos(ppto, rs, "OPEN", strSQL)		  
Set ArmarColumnaAjustes = rs
'response.write strSQL
'response.end
End function		   
'------------------------------------------------------------------------------------
Function ArmarColumnaRecargas(pMyFecha,ppto,campo,pLstClientes)
Dim strSQL
strSQL = "				SELECT t2.cdcliente,"
strSQL = strSQL &"			   t2.dscliente,"
strSQL = strSQL &"			   t2.cdproducto,"
strSQL = strSQL &"			   t2.dsproducto"
strSQL = strSQL &"		FROM   (SELECT T1.cdcliente,"
strSQL = strSQL &"					   c.dscliente,"
strSQL = strSQL &"					   T1.cdproducto,"
strSQL = strSQL &"					   p.dsproducto"
strSQL = strSQL &"				FROM   (SELECT cc.cdcliente,"
strSQL = strSQL &"							   ca.cdproducto"
strSQL = strSQL &"						FROM   dbo.hcamionescarga cc"
strSQL = strSQL &"							   JOIN dbo.hcamiones ca"
strSQL = strSQL &"								 ON ca.idcamion = cc.idcamion"
strSQL = strSQL &"									AND ca.dtcontable = cc.dtcontable"
strSQL = strSQL &"						WHERE  cc.dtcontable = '"&pMyFecha&"'"
if (pLstClientes <> "") then  strSQL = strSQL &"	and cc.cdcliente in ("& pLstClientes & ")"
strSQL = strSQL &"							   AND (SELECT Max(pc.sqpesada)"
strSQL = strSQL &"									FROM   dbo.hpesadascamion pc"
strSQL = strSQL &"									WHERE  pc.dtcontable = cc.dtcontable"
strSQL = strSQL &"										   AND pc.idcamion = cc.idcamion"
strSQL = strSQL &"										   AND pc.cdpesada = 1) <> 0"
strSQL = strSQL &"							   AND ca.cdestado NOT IN ( 12, 7, 14 )) AS T1"
strSQL = strSQL &"					   LEFT JOIN dbo.clientes c"
strSQL = strSQL &"							  ON c.cdcliente = T1.cdcliente"
strSQL = strSQL &"					   LEFT JOIN dbo.productos p"
strSQL = strSQL &"							  ON p.cdproducto = T1.cdproducto"
strSQL = strSQL &"				GROUP  BY t1.cdcliente,"
strSQL = strSQL &"						  c.dscliente,"
strSQL = strSQL &"						  T1.cdproducto,"
strSQL = strSQL &"						  p.dsproducto) AS t2  "
strSQL = strSQL &"						  order by " & campo         
'response.write strSQL
'response.end
Call GF_BD_Puertos(ppto, rs, "OPEN", strSQL)		  
Set ArmarColumnaRecargas = rs
End function						  
'------------------------------------------------------------------------------------
Function ArmarColumnaTransferencia(pMyFecha,ppto,campo,pLstClientes)
strSQL ="				SELECT DISTINCT "&campo
strSQL = strSQL & "  	FROM   dbo.exmovimientos m			"
strSQL = strSQL & " 	LEFT JOIN dbo.clientes c			"
strSQL = strSQL & " 		ON c.cdcliente = m.cdcliente		"
strSQL = strSQL & " 	LEFT JOIN dbo.productos p			"
strSQL = strSQL & " 		ON p.cdproducto = m.cdproducto		"
strSQL = strSQL & " 	WHERE  m.dtcontable = '"& pMyFecha &"'	"
if (pLstClientes <> "") then strSQL = strSQL & " AND m.cdcliente in ("& pLstClientes & ")"
strSQL = strSQL & "		    AND m.vlkilos <> 0					"
strSQL = strSQL & " 		AND ( ( m.cdtransaccion >= 102		"
strSQL = strSQL & " 				AND m.cdtransaccion <= 107 )"
strSQL = strSQL & " 			OR ( m.cdtransaccion >= 2		"
strSQL = strSQL & " 				AND m.cdtransaccion <= 7 ) )"
strSQL = strSQL & "		GROUP  BY "&campo
strSQL = strSQL & "		ORDER  BY "&campo
Call GF_BD_Puertos(ppto, rs, "OPEN", strSQL)		  
'response.write strSQL
Set ArmarColumnaTransferencia = rs
'response.write strSQL
'response.end
End function
'------------------------------------------------------------------------------------
Function ArmarColumnaOleoducto(pMyFecha,ppto,campo,pLstClientes)
Dim strSQL 
strSQL =" 				SELECT DISTINCT "&campo
strSQL = strSQL &" 		FROM   dbo.extransfoleo m				"
strSQL = strSQL &" 		LEFT JOIN dbo.clientes c				"
strSQL = strSQL &" 			ON c.cdcliente = m.cdcliente			"
strSQL = strSQL &" 		LEFT JOIN dbo.productos p				"
strSQL = strSQL &" 			ON p.cdproducto = m.cdproducto			"
strSQL = strSQL &" 		WHERE  m.dtcontable = '"& pMyFecha &"'		"
strSQL = strSQL &" 			AND m.vlkilos <> 0						"
if (pLstClientes <> "") then strSQL = strSQL &" AND m.cdcliente in ("& pLstClientes & ")"
strSQL = strSQL &" 		GROUP  BY "&campo
Call GF_BD_Puertos(ppto, rs, "OPEN", strSQL)		  
Set ArmarColumnaOleoducto = rs
'response.write strSQL
'response.end
End function
'------------------------------------------------------------------------------------
Function armarFilasClientes(pMyFecha,ppto,pLstClientes)
Dim strSQL
strSQL = "				SELECT DISTINCT e.cdcliente,										"
strSQL = strSQL & " 					C.dscliente AS dsCliente							"	
strSQL = strSQL & " 	FROM  excuentcorrientes e											"
strSQL = strSQL & " 		LEFT JOIN dbo.clientes C									"
strSQL = strSQL & " 			ON C.cdcliente = e.cdcliente								"
strSQL = strSQL & " 	WHERE  e.dtcontable = '"& pMyFecha &"'								"
if (pLstClientes <> "") then strSQL = strSQL & " and e.cdcliente in ("& pLstClientes & ")"
strSQL = strSQL & " 		AND ( e.vlsaldoinicial + e.vlcredito - e.vldebito ) <> 0		"
strSQL = strSQL & " 	GROUP  BY e.cdcliente,C.dscliente						  			"
Call GF_BD_Puertos(ppto, rs, "OPEN", strSQL)		  
Set armarFilasClientes = rs
'response.write strSQL
'response.end
End function 
'------------------------------------------------------------------------------------
Function armarColumnaClientesCamiones(pMyFecha,ppto,campo,pLstClientes)
dim strSQL

strSQL= "SELECT DISTINCT " & campo & _
		" FROM   ((SELECT T1.cdcliente," & _
		"          T1.cdproducto,T1.dtcontable" & _
		"   FROM   (SELECT cd.cdcliente,ca.cdproducto,cd.dtcontable " & _
		"           FROM   dbo.hcamionesdescarga cd" & _
		"                  JOIN dbo.hcamiones ca" & _
		"                    ON ca.idcamion = cd.idcamion" & _
		"                       AND ca.dtcontable = cd.dtcontable" & _
		"           WHERE  cd.dtcontable = '" & pMyFecha & "'"
        if (pLstClientes <> "") then strSQL = strSQL & " and cd.CDCLIENTE in ("& pLstClientes & ")"
		strSQL = strSQL & "            AND (SELECT Max(sqpesada)" & _
		"                                   FROM   dbo.hpesadascamion" & _
		"                                   WHERE  dtcontable =cd.dtcontable" & _
		"                                          AND idcamion = cd.idcamion" & _
		"                                          AND cdpesada = 2) is not NULL" & _
		"                  AND ca.cdestado NOT IN ( 12, 7, 14 )) AS T1              " & _
		"   GROUP  BY t1.cdcliente,      " & _           
		"             T1.cdproducto,T1.dtcontable)" & _
		"    UNION" & _
		"    ( SELECT T1.cdcliente," & _
		"           T1.cdproducto, '"&Year(Now())& "-" & GF_nDigits(Month(Now()), 2) & "-" & GF_nDigits(Day(Now()), 2)&"' as dtcontable" & _
		"    FROM   (SELECT cd.cdcliente," & _
		"                  ca.cdproducto" & _
		"           FROM   dbo.camionesdescarga cd" & _
		"                  JOIN dbo.camiones ca" & _
		"                    ON ca.idcamion = cd.idcamion" & _
		"           WHERE  (SELECT Max(sqpesada)" & _
		"                                   FROM   dbo.pesadascamion" & _
		"                                   WHERE  idcamion = cd.idcamion" & _
		"                                          AND cdpesada = 2) is not NULL"
        if (pLstClientes <> "")  then strSQL = strSQL & " and cd.CDCLIENTE in ("& pLstClientes & ")"
		strSQL = strSQL & " AND ca.cdestado NOT IN ( 12, 7, 14 )) AS T1             " & _
		"   GROUP  BY t1.cdcliente,            " & _    
		"             T1.cdproducto)) AS TF" & _
		"      LEFT JOIN dbo.clientes c" & _
		"                 ON c.cdcliente = TF.cdcliente" & _
		"          LEFT JOIN dbo.productos p" & _
		"                 ON p.cdproducto = TF.cdproducto" & _
        "   WHERE  TF.dtcontable = '" & pMyFecha & "'" & _
		" order by " & campo
Call GF_BD_Puertos(ppto, rs, "OPEN", strSQL)		  
Set armarColumnaClientesCamiones = rs
End function 
'------------------------------------------------------------------------------------
Function armarColumnasProductos(pMyFecha,ppto,pLstClientes)
Dim strSQL
strSQL = "				select distinct e.cdproducto,										"
strSQL = strSQL & "						p.DSPRODUCTO as dsProducto							" 
strSQL = strSQL & "		from excuentcorrientes e 											"
strSQL = strSQL & " 		left join productos p on e.cdproducto = p.cdproducto			"
strSQL = strSQL & " 	where e.dtcontable = '"& pMyFecha &"'								"
if (pLstClientes <> "") then strSQL = strSQL & " and e.cdcliente in ("& pLstClientes & ")"
strSQL = strSQL & " 		AND ( e.vlsaldoinicial + e.vlcredito - e.vldebito ) <> 0		"
strSQL = strSQL & " 	group by e.cdproducto,p.DSPRODUCTO									"
Call GF_BD_Puertos(ppto, rs, "OPEN", strSQL)		  
Set armarColumnasProductos = rs
End function 
'------------------------------------------------------------------------------------
Function ObtenerSQLSaldoInicial(pMyFecha,ppto,pLstClientes)
Dim strSQL, Mywhere,rs
strSQL = " 				SELECT  cc.cdcliente,												"
strSQL = strSQL & " 			dscliente,													"
strSQL = strSQL & " 			cc.cdproducto,												"
strSQL = strSQL & " 			dsproducto,													"
strSQL = strSQL & " 			( cc.vlsaldoinicial + cc.vlcredito - cc.vldebito ) as Kilos "
strSQL = strSQL & " 	FROM  dbo.excuentcorrientes cc									"
strSQL = strSQL & " 		LEFT JOIN dbo.clientes C									"
strSQL = strSQL & " 			ON C.cdcliente = cc.cdcliente								"
strSQL = strSQL & " 		LEFT JOIN dbo.productos P									"
strSQL = strSQL & " 			ON P.cdproducto = cc.cdproducto								"
strSQL = strSQL & " 	WHERE  cc.dtcontable = '"&pMyFecha&"'								"
strSQL = strSQL & " 		AND ( cc.vlsaldoinicial + cc.vlcredito - cc.vldebito ) <> 0		"
if (pLstClientes <> "") then strSQL = strSQL & " and cc.cdcliente in ("& pLstClientes & ")"
strSQL = strSQL & " 	ORDER  BY cc.cdcliente, 											"
strSQL = strSQL & " 			  cc.cdproducto	"
Call GF_BD_Puertos(ppto, rs, "OPEN", strSQL)		  
Set ObtenerSQLSaldoInicial = rs
End function
'---------------------------------------------------------------------------------------
Function ObtenerSQLDescargaVagones(pMyFecha,ppto,pLstClientes, campo)
Dim strSQL, Mywhere,rs,myFechaActual

if (campo <> "") then
    strSQL = strSQL & "Select DISTINCT " & campo 
    strSQL = strSQL & " from ( "                     
end if
strSQL = strSQL & "		SELECT T1.cdempresa,"
strSQL = strSQL & "			   E.dsempresa,"
strSQL = strSQL & "			   T1.cdcliente,"
strSQL = strSQL & "			   c.dscliente,"
strSQL = strSQL & "			   T1.cdproducto,"
strSQL = strSQL & "			   p.dsproducto,"
strSQL = strSQL & "			   Sum(T1.kilos) AS Cant1,"
strSQL = strSQL & "			   Count(*)      AS Cant2"
strSQL = strSQL & "		FROM   (SELECT o.cdempresa,"
strSQL = strSQL & "					   o.cdcliente,"
strSQL = strSQL & "					   va.cdproducto,"
strSQL = strSQL & "					   va.cdoperativo,"
strSQL = strSQL & "					   va.cdvagon,"
strSQL = strSQL & "					   ( (SELECT pv.vlpesada"
strSQL = strSQL & "						  FROM   dbo.hpesadasvagon pv"
strSQL = strSQL & "						  WHERE  pv.dtcontable = va.dtcontable"
strSQL = strSQL & "								 AND pv.cdoperativo = va.cdoperativo"
strSQL = strSQL & "								 AND pv.nucartaporte = va.nucartaporte"
strSQL = strSQL & "								 AND pv.nucartaporteserie = va.nucartaporteserie"
strSQL = strSQL & "								 AND pv.cdvagon = va.cdvagon"
strSQL = strSQL & "								 AND pv.cdpesada = 1"
strSQL = strSQL & "								 AND pv.sqpesada = (SELECT Max(sqpesada)"
strSQL = strSQL & "													FROM   dbo.hpesadasvagon"
strSQL = strSQL & "													WHERE  dtcontable = pv.dtcontable"
strSQL = strSQL & "														   AND cdoperativo ="
strSQL = strSQL & "															   pv.cdoperativo"
strSQL = strSQL & "														   AND cdoperativoserie ="
strSQL = strSQL & "															   pv.cdoperativoserie"
strSQL = strSQL & "														   AND nucartaporte ="
strSQL = strSQL & "															   pv.nucartaporte"
strSQL = strSQL & "														   AND nucartaporteserie ="
strSQL = strSQL & "															   pv.nucartaporteserie"
strSQL = strSQL & "														   AND cdvagon = va.cdvagon"
strSQL = strSQL & "														   AND cdpesada = 1)) -"
strSQL = strSQL & "						   (SELECT pv.vlpesada"
strSQL = strSQL & "							FROM   dbo.hpesadasvagon pv"
strSQL = strSQL & "							WHERE  pv.dtcontable = va.dtcontable"
strSQL = strSQL & "								   AND pv.cdoperativo = va.cdoperativo"
strSQL = strSQL & "								   AND pv.cdoperativoserie = va.cdoperativoserie"
strSQL = strSQL & "								   AND pv.nucartaporte = va.nucartaporte"
strSQL = strSQL & "								   AND pv.nucartaporteserie = va.nucartaporteserie"
strSQL = strSQL & "								   AND pv.cdvagon = va.cdvagon"
strSQL = strSQL & "								   AND pv.cdpesada = 2"
strSQL = strSQL & "								   AND pv.sqpesada = (SELECT Max(sqpesada)"
strSQL = strSQL & "													  FROM   dbo.hpesadasvagon"
strSQL = strSQL & "													  WHERE  dtcontable = pv.dtcontable"
strSQL = strSQL & "															 AND cdoperativo ="
strSQL = strSQL & "																 pv.cdoperativo"
strSQL = strSQL & "															 AND cdoperativoserie ="
strSQL = strSQL & "																 pv.cdoperativoserie"
strSQL = strSQL & "															 AND nucartaporte ="
strSQL = strSQL & "																 pv.nucartaporte"
strSQL = strSQL & "															 AND nucartaporteserie ="
strSQL = strSQL & "																 pv.nucartaporteserie"
strSQL = strSQL & "															 AND cdvagon = va.cdvagon"
strSQL = strSQL & "															 AND cdpesada = 2)) -"
strSQL = strSQL & "						   (SELECT pv.vlmermakilos"
strSQL = strSQL & "							FROM   dbo.hmermasvagones pv"
strSQL = strSQL & "							WHERE  pv.dtcontable = va.dtcontable"
strSQL = strSQL & "								   AND pv.cdoperativo = va.cdoperativo"
strSQL = strSQL & "								   AND pv.cdoperativoserie = va.cdoperativoserie"
strSQL = strSQL & "								   AND pv.nucartaporte = va.nucartaporte"
strSQL = strSQL & "								   AND pv.nucartaporteserie = va.nucartaporteserie"
strSQL = strSQL & "								   AND pv.cdvagon = va.cdvagon"
strSQL = strSQL & "								   AND pv.sqpesada = (SELECT Max(sqpesada)"
strSQL = strSQL & "													  FROM   dbo.hpesadasvagon"
strSQL = strSQL & "													  WHERE  dtcontable = pv.dtcontable"
strSQL = strSQL & "															 AND cdoperativo ="
strSQL = strSQL & "																 pv.cdoperativo"
strSQL = strSQL & "															 AND cdoperativoserie ="
strSQL = strSQL & "																 pv.cdoperativoserie"
strSQL = strSQL & "															 AND nucartaporte ="
strSQL = strSQL & "																 pv.nucartaporte"
strSQL = strSQL & "															 AND nucartaporteserie ="
strSQL = strSQL & "																 pv.nucartaporteserie"
strSQL = strSQL & "															 AND cdvagon = va.cdvagon"
strSQL = strSQL & "															 AND cdpesada = 2)) ) AS"
strSQL = strSQL & "					   Kilos"
strSQL = strSQL & "				FROM   dbo.hvagones va"
strSQL = strSQL & "					   JOIN dbo.hoperativos o"
strSQL = strSQL & "						 ON o.cdoperativo = va.cdoperativo"
strSQL = strSQL & "							AND o.cdoperativoserie = va.cdoperativoserie"
strSQL = strSQL & "							AND o.nucartaporte = va.nucartaporte"
strSQL = strSQL & "							AND o.nucartaporteserie = va.nucartaporteserie"
strSQL = strSQL & "							AND o.dtcontable = va.dtcontable"
strSQL = strSQL & "				WHERE  va.dtcontablevagon = '"&pMyFecha&"'"
if (pLstClientes <> "") then strSQL = strSQL & "	and o.cdcliente in ("& pLstClientes & ")"
'strSQL = strSQL & "					   AND (SELECT pv.sqpesada"
'strSQL = strSQL & "							FROM   dbo.hpesadasvagon pv"
'strSQL = strSQL & "							WHERE  PV.dtcontable = va.dtcontable"
'strSQL = strSQL & "								   AND PV.cdoperativo = va.cdoperativo"
'strSQL = strSQL & "								   AND pv.nucartaporte = va.nucartaporte"
'strSQL = strSQL & "								   AND pv.nucartaporteserie = va.nucartaporteserie"
'strSQL = strSQL & "								   AND pv.cdvagon = va.cdvagon"
'strSQL = strSQL & "								   AND pv.cdpesada = 2"
'strSQL = strSQL & "								   AND pv.sqpesada = (SELECT Max(sqpesada)"
'strSQL = strSQL & "													  FROM   dbo.hpesadasvagon"
'strSQL = strSQL & "													  WHERE  dtcontable = pv.dtcontable"
'strSQL = strSQL & "															 AND cdoperativo ="
'strSQL = strSQL & "																 pv.cdoperativo"
'strSQL = strSQL & "															 AND cdoperativoserie ="
'strSQL = strSQL & "																 pv.cdoperativoserie"
'strSQL = strSQL & "															 AND nucartaporte ="
'strSQL = strSQL & "																 pv.nucartaporte"
'strSQL = strSQL & "															 AND nucartaporteserie ="
'strSQL = strSQL & "																 pv.nucartaporteserie"
'strSQL = strSQL & "															 AND cdvagon = pv.cdvagon"
'strSQL = strSQL & "															 AND cdpesada = 2)) <> 0"
strSQL = strSQL & "					   AND va.cdestado = 8"
strSQL = strSQL & "				UNION"
strSQL = strSQL & "				SELECT o.cdempresa,"
strSQL = strSQL & "					   o.cdcliente,"
strSQL = strSQL & "					   va.cdproducto,"
strSQL = strSQL & "					   va.cdoperativo,"
strSQL = strSQL & "					   va.cdvagon,"
strSQL = strSQL & "					   ( (SELECT pv.vlpesada"
strSQL = strSQL & "						  FROM   dbo.pesadasvagon pv"
strSQL = strSQL & "						  WHERE  pv.cdoperativo = va.cdoperativo"
strSQL = strSQL & "								 AND pv.cdoperativoserie = va.cdoperativoserie"
strSQL = strSQL & "								 AND pv.nucartaporte = va.nucartaporte"
strSQL = strSQL & "								 AND pv.nucartaporteserie = va.nucartaporteserie"
strSQL = strSQL & "								 AND pv.cdvagon = va.cdvagon"
strSQL = strSQL & "								 AND pv.cdpesada = 1"
strSQL = strSQL & "								 AND pv.sqpesada = (SELECT Max(sqpesada)"
strSQL = strSQL & "													FROM   dbo.pesadasvagon"
strSQL = strSQL & "													WHERE  cdoperativo = pv.cdoperativo"
strSQL = strSQL & "														   AND cdoperativoserie ="
strSQL = strSQL & "															   pv.cdoperativoserie"
strSQL = strSQL & "														   AND nucartaporte ="
strSQL = strSQL & "															   pv.nucartaporte"
strSQL = strSQL & "														   AND nucartaporteserie ="
strSQL = strSQL & "															   pv.nucartaporteserie"
strSQL = strSQL & "														   AND cdvagon = va.cdvagon"
strSQL = strSQL & "														   AND cdpesada = 1)) -"
strSQL = strSQL & "						   (SELECT pv.vlpesada"
strSQL = strSQL & "							FROM   dbo.pesadasvagon pv"
strSQL = strSQL & "							WHERE  pv.cdoperativo = va.cdoperativo"
strSQL = strSQL & "								   AND pv.cdoperativoserie = va.cdoperativoserie"
strSQL = strSQL & "								   AND pv.nucartaporte = va.nucartaporte"
strSQL = strSQL & "								   AND pv.nucartaporteserie = va.nucartaporteserie"
strSQL = strSQL & "								   AND pv.cdvagon = va.cdvagon"
strSQL = strSQL & "								   AND pv.cdpesada = 2"
strSQL = strSQL & "								   AND pv.sqpesada = (SELECT Max(sqpesada)"
strSQL = strSQL & "													  FROM   dbo.pesadasvagon"
strSQL = strSQL & "													  WHERE"
strSQL = strSQL & "									   cdoperativo = pv.cdoperativo"
strSQL = strSQL & "									   AND cdoperativoserie ="
strSQL = strSQL & "										   pv.cdoperativoserie"
strSQL = strSQL & "									   AND nucartaporte = pv.nucartaporte"
strSQL = strSQL & "									   AND nucartaporteserie ="
strSQL = strSQL & "										   pv.nucartaporteserie"
strSQL = strSQL & "									   AND cdvagon = va.cdvagon"
strSQL = strSQL & "									   AND cdpesada = 2)) -"
strSQL = strSQL & "						   (SELECT pv.vlmermakilos"
strSQL = strSQL & "							FROM   dbo.mermasvagones pv"
strSQL = strSQL & "							WHERE  pv.cdoperativo = va.cdoperativo"
strSQL = strSQL & "								   AND pv.cdoperativoserie = va.cdoperativoserie"
strSQL = strSQL & "								   AND pv.nucartaporte = va.nucartaporte"
strSQL = strSQL & "								   AND pv.nucartaporteserie = va.nucartaporteserie"
strSQL = strSQL & "								   AND pv.cdvagon = va.cdvagon"
strSQL = strSQL & "								   AND pv.sqpesada = (SELECT Max(sqpesada)"
strSQL = strSQL & "													  FROM   dbo.pesadasvagon"
strSQL = strSQL & "													  WHERE"
strSQL = strSQL & "									   cdoperativo = pv.cdoperativo"
strSQL = strSQL & "									   AND cdoperativoserie ="
strSQL = strSQL & "										   pv.cdoperativoserie"
strSQL = strSQL & "									   AND nucartaporte = pv.nucartaporte"
strSQL = strSQL & "									   AND nucartaporteserie ="
strSQL = strSQL & "										   pv.nucartaporteserie"
strSQL = strSQL & "									   AND cdvagon = va.cdvagon"
strSQL = strSQL & "									   AND cdpesada = 2)) ) AS Kilos"
strSQL = strSQL & "				FROM   dbo.vagones va"
strSQL = strSQL & "					   JOIN dbo.operativos o"
strSQL = strSQL & "						 ON o.cdoperativo = va.cdoperativo"
strSQL = strSQL & "							AND o.cdoperativoserie = va.cdoperativoserie"
strSQL = strSQL & "							AND o.nucartaporte = va.nucartaporte"
strSQL = strSQL & "							AND o.nucartaporteserie = va.nucartaporteserie"
strSQL = strSQL & "				WHERE  va.dtcontablevagon = '"&pMyFecha&"'"
if (pLstClientes <> "") then strSQL = strSQL & "	and o.cdcliente in ("& pLstClientes & ")"
'strSQL = strSQL & "					   AND (SELECT pv.sqpesada"
'strSQL = strSQL & "							FROM   dbo.pesadasvagon pv"
'strSQL = strSQL & "							WHERE  pv.cdoperativo = va.cdoperativo"
'strSQL = strSQL & "								   AND pv.nucartaporte = va.nucartaporte"
'strSQL = strSQL & "								   AND pv.cdoperativoserie = va.cdoperativoserie"
'strSQL = strSQL & "								   AND pv.nucartaporteserie = va.nucartaporteserie"
'strSQL = strSQL & "								   AND pv.cdvagon = va.cdvagon"
'strSQL = strSQL & "								   AND pv.cdpesada = 2"
'strSQL = strSQL & "								   AND pv.sqpesada = (SELECT Max(sqpesada)"
'strSQL = strSQL & "													  FROM   dbo.pesadasvagon"
'strSQL = strSQL & "													  WHERE"
'strSQL = strSQL & "									   cdoperativo = pv.cdoperativo"
'strSQL = strSQL & "									   AND cdoperativoserie ="
'strSQL = strSQL & "										   pv.cdoperativoserie"
'strSQL = strSQL & "									   AND nucartaporte = pv.nucartaporte"
'strSQL = strSQL & "									   AND nucartaporteserie ="
'strSQL = strSQL & "										   pv.nucartaporteserie"
'strSQL = strSQL & "									   AND cdvagon = pv.cdvagon"
'strSQL = strSQL & "									   AND cdpesada = 2)) <> 0"
strSQL = strSQL & "                    and not exists(Select * from HVAGONES CTV where CTV.CDOPERATIVO=va.CDOPERATIVO and CTV.CDVAGON=va.CDVAGON)"
strSQL = strSQL & "					   AND va.cdestado = 8) AS t1"
strSQL = strSQL & "			   LEFT JOIN dbo.empresas e"
strSQL = strSQL & "					  ON e.cdempresa = T1.cdempresa"
strSQL = strSQL & "			   LEFT JOIN dbo.clientes c"
strSQL = strSQL & "					  ON c.cdcliente = T1.cdcliente"
strSQL = strSQL & "			   LEFT JOIN dbo.productos p"
strSQL = strSQL & "					  ON p.cdproducto = T1.cdproducto"
strSQL = strSQL & "		GROUP  BY t1.cdempresa,"
strSQL = strSQL & "				  E.dsempresa,"
strSQL = strSQL & "				  t1.cdcliente,"
strSQL = strSQL & "				  c.dscliente,"
strSQL = strSQL & "				  T1.cdproducto,"
strSQL = strSQL & "				  p.dsproducto  "
if (campo <> "") then    
    strSQL = strSQL & " ) T1 "                     
end if
'response.Write strSQL & "<br>"
Call GF_BD_Puertos(ppto, rs, "OPEN", strSQL)
set ObtenerSQLDescargaVagones = rs		  
End function
'----------------------------------------------------------------------------------------
Function ObtenerSQLDescargaCamiones(pMyFecha, ppto, pLstClientes)
Dim strSQL, Mywhere,rs
		
strSQL =" 	SELECT t2.cdcliente,t2.dscliente,t2.cdproducto, t2.dsproducto,	                "&_
		"	        sum(t2.cant)  AS Cant2, sum(t2.Cant1) as Cant1 		                    "&_
		" FROM   (SELECT 	T1.cdcliente,													"&_
		"					c.dscliente,													"&_
		"					T1.cdproducto,													"&_
		"					p.dsproducto,													"&_
        "           ( 																		"&_
		"				(SELECT Sum(p.vlpesada)												"&_
		"				 FROM   hcamiones ca,hcamionesdescarga cd,hpesadascamion p			"&_
		"				 WHERE  ca.dtcontable = cd.dtcontable								"&_
		"					 AND ca.idcamion = cd.idcamion									"&_
		"					 AND ca.dtcontable = p.dtcontable								"&_
		"					 AND ca.idcamion = p.idcamion									"&_
		"					 AND p.cdpesada = 1												"&_
		"					 AND p.sqpesada = (SELECT Max(sqpesada)							"&_
		"					 					FROM   hpesadascamion						"&_
		"										WHERE  dtcontable = p.dtcontable			"&_
		"											 AND idcamion = p.idcamion				"&_
		"											 AND cdpesada = 1)						"&_
		"					 AND ca.cdproducto = t1.cdproducto								"&_
		"					 AND cd.cdcliente = t1.cdcliente								"&_
		"					 AND cd.dtcontable = '"&pMyFecha&"'								"&_
		"					 AND ca.cdestado NOT IN ( 12, 7, 14 ))							"&_
		"				-																 	"&_
		"				(SELECT Sum(p.vlpesada)												"&_
		"				 FROM   hcamiones ca,hcamionesdescarga cd,hpesadascamion p			"&_
		"				 WHERE	ca.idcamion = cd.idcamion									"&_
		"					AND ca.dtcontable = cd.dtcontable								"&_
		"					AND ca.dtcontable = p.dtcontable								"&_
		"					AND ca.idcamion = p.idcamion									"&_
		"					AND p.cdpesada = 2												"&_
		"					AND p.sqpesada = (SELECT Max(sqpesada)							"&_
		"									   FROM   hpesadascamion						"&_	
		"									   WHERE  dtcontable = p.dtcontable				"&_
		"										  AND idcamion = p.idcamion					"&_
		"										  AND cdpesada = 2)							"&_
		"					AND ca.cdproducto = t1.cdproducto								"&_
		" 					AND cdcliente = t1.cdcliente									"&_
		"					AND cd.dtcontable = '"&pMyFecha&"'								"&_
		"					AND ca.cdestado NOT IN ( 12, 7, 14 ))							"&_
		"				-																	"&_
		"				(SELECT Sum(m.vlmermakilos)											"&_
		"				 FROM   hcamiones ca, hcamionesdescarga cd, hmermascamiones m		"&_
		"				 WHERE ca.dtcontable = cd.dtcontable								"&_
		"					 AND ca.idcamion = cd.idcamion									"&_
		"			 		 AND ca.dtcontable = m.dtcontable								"&_
		"					 AND ca.idcamion = m.idcamion									"&_
		"					 AND m.sqpesada =  (SELECT Max(sqpesada)						"&_
		" 										FROM   hpesadascamion						"&_
		"										WHERE  dtcontable = m.dtcontable			"&_
		"											 AND idcamion = m.idcamion				"&_
		"											 AND cdpesada = 2)						"&_
		"					 AND ca.cdproducto = t1.cdproducto								"&_		
		"					 AND cdcliente = t1.cdcliente									"&_
		"					 AND cd.dtcontable = '"&pMyFecha&"'								"&_
		"					 AND ca.cdestado NOT IN ( 12, 7, 14 )) 							"&_
		"			) AS Cant1,																"&_
		"					(SELECT Count(*)												"&_
		"					 FROM   dbo.hcamionesdescarga cd,dbo.hcamiones ca		"&_
		"					 WHERE  cd.dtcontable = ca.dtcontable							"&_
		"						AND cd.idcamion = ca.idcamion								"&_
		"						AND ca.cdproducto = t1.cdproducto							"&_
		"						AND cd.cdcliente = t1.cdcliente								"&_
		"						AND cd.dtcontable = '"&pMyFecha&"'							"&_
		"					 	AND ca.cdestado NOT IN ( 12, 7, 14 )) AS Cant				"&_	
		"		  FROM   (SELECT cd.cdcliente,ca.cdproducto									"&_
		"				  FROM   dbo.hcamionesdescarga cd								"&_
		"				  JOIN dbo.hcamiones ca										"&_
		"					 ON ca.idcamion = cd.idcamion									"&_
		"				 	 AND ca.dtcontable = cd.dtcontable								"&_
		"				  WHERE  cd.dtcontable = '"&pMyFecha&"'								"
        if (pLstClientes <> "") then strSQL = strSQL & " and cd.cdcliente in ("& pLstClientes & ")"
        strSQL = strSQL &"	 AND (SELECT pc.sqpesada										"&_
		"						  FROM   dbo.hpesadascamion pc							"&_
		"						  WHERE  pc.dtcontable = cd.dtcontable						"&_
		"							 AND pc.idcamion = cd.idcamion							"&_
		"							 AND pc.cdpesada = 2									"&_
		"							 AND pc.sqpesada = (SELECT Max(sqpesada)				"&_
		"												FROM   dbo.hpesadascamion		"&_
		"												WHERE  dtcontable = pc.dtcontable	"&_
		"													 AND idcamion = pc.idcamion		"&_
		"													 AND cdpesada = 2)				"&_
		"							) > 0													"&_
		"					 AND ca.cdestado NOT IN ( 12, 7, 14 )) AS T1					"&_
		"		  LEFT JOIN dbo.clientes c												"&_
		"			 ON c.cdcliente = T1.cdcliente											"&_
		"		 LEFT JOIN dbo.productos p												"&_
		"			 ON p.cdproducto = T1.cdproducto										"&_
		"  		GROUP  BY t1.cdcliente,c.dscliente,T1.cdproducto,							"&_
		"		 p.dsproducto) AS t2  GROUP BY t2.cdcliente,t2.dscliente,t2.cdproducto, t2.dsproducto"&_
        "       ORDER  BY T2.cdcliente, T2.cdproducto	    			"
'Response.Write strSQL &"<BR>"
'Response.End
Call GF_BD_Puertos(ppto, rs, "OPEN", strSQL)
set ObtenerSQLDescargaCamiones = rs
End function
'-------------------------------------------------------------------------------------------
Function obtenerSQLTransferencia(pMyFecha,pto,pLstClientes)
Dim strSQL, rs
strSQL ="			 SELECT T1.*,														"
'						Recibidos
strSQL = strSQL &" 			( SELECT Sum(m.vlkilos)										"
strSQL = strSQL &" 			  FROM dbo.exmovimientos m								"
strSQL = strSQL &" 			  WHERE  m.dtcontable = '"&pMyFecha&"'						"
strSQL = strSQL &" 				 AND m.vlkilos <> 0										"
if (pLstClientes <> "") then strSQL = strSQL &" AND m.cdcliente in ("& pLstClientes & ")"
strSQL = strSQL &" 				 AND m.cdcliente = T1.cdcliente							"
strSQL = strSQL &" 				 AND m.cdproducto = T1.cdproducto						"
strSQL = strSQL &" 				 AND m.cdtransaccion IN( 2, 3, 4, 5,6, 7 )				"
strSQL = strSQL &"			 ) AS Cant1,												"
'						Transferidos
strSQL = strSQL &" 			( SELECT Sum(mv.vlkilos)									"
strSQL = strSQL &" 			  FROM dbo.exmovimientos mv							"
strSQL = strSQL &" 			  WHERE  mv.dtcontable = '"& pMyFecha &"'					"
if (pLstClientes <> "") then strSQL = strSQL &" AND mv.cdcliente in ("& pLstClientes & ")"
strSQL = strSQL &" 				 AND mv.vlkilos <> 0									"		
strSQL = strSQL &" 				 AND mv.cdcliente = T1.cdcliente						"
strSQL = strSQL &" 				 AND mv.cdproducto = T1.cdproducto						"
strSQL = strSQL &" 				 AND mv.cdtransaccion IN( 102, 103, 104, 105,106, 107 )	"	
strSQL = strSQL &"			 ) AS Cant2													"
strSQL = strSQL &" 	  FROM   (SELECT DISTINCT m.cdcliente,								"			
strSQL = strSQL &" 							  dscliente,								"
strSQL = strSQL &" 							  m.cdproducto,								"
strSQL = strSQL &" 						      dsproducto								"
strSQL = strSQL &" 			  FROM dbo.exmovimientos m								"
strSQL = strSQL &" 			  LEFT JOIN dbo.clientes c								"
strSQL = strSQL &"				 ON c.cdcliente = m.cdcliente							"
strSQL = strSQL &" 			  LEFT JOIN dbo.productos p							"
strSQL = strSQL &" 				 ON p.cdproducto = m.cdproducto							"
strSQL = strSQL &" 			  WHERE  m.dtcontable = '"& pMyFecha &"'					"		
if (pLstClientes <> "") then strSQL = strSQL &" AND m.cdcliente in ("& pLstClientes & ")"
strSQL = strSQL &" 					AND m.vlkilos <> 0									"
strSQL = strSQL &" 					AND ( ( m.cdtransaccion >= 102						"
strSQL = strSQL &" 							AND m.cdtransaccion <= 107 )				"
strSQL = strSQL &" 						  OR ( m.cdtransaccion >= 2						"
strSQL = strSQL &" 							AND m.cdtransaccion <= 7 ) )				"
strSQL = strSQL &" 			  GROUP  BY m.cdcliente,									"
strSQL = strSQL &"						dscliente,										"
strSQL = strSQL &" 						m.cdproducto,									"
strSQL = strSQL &" 						dsproducto 	) AS T1								"
strSQL = strSQL &" 	ORDER BY  T1.cdcliente,									"
strSQL = strSQL &" 			  T1.cdproducto 									"		 
'response.write strSQL
'response.end
Call GF_BD_Puertos(pto, rs, "OPEN", strSQL)
set obtenerSQLTransferencia = rs
End function
'--------------------------------------------------------------------------
Function obtenerSQLOleoducto(pMyFecha,pto,pLstClientes)
Dim strSQL, rs

strSQL = " 			SELECT  m.cdcliente,					"
strSQL = strSQL &" 			c.dscliente,					"
strSQL = strSQL &" 			m.cdproducto,					"
strSQL = strSQL &" 			p.dsproducto,					"
strSQL = strSQL &" 			Count(*) AS Cant2,				"
strSQL = strSQL &" 			Sum(m.vlkilos) AS Cant1			"
strSQL = strSQL &"  FROM dbo.extransfoleo m			"
strSQL = strSQL &" 	LEFT JOIN dbo.clientes c			"
strSQL = strSQL &" 		ON c.cdcliente = m.cdcliente		"
strSQL = strSQL &" 	LEFT JOIN dbo.productos p			"
strSQL = strSQL &" 		ON p.cdproducto = m.cdproducto		"
strSQL = strSQL &" 	WHERE  m.dtcontable = '" & pMyFecha &"' "
strSQL = strSQL &" 		AND m.vlkilos <> 0					"
if (pLstClientes <> "") then strSQL = strSQL &" AND m.cdcliente in ("& pLstClientes & ")"
strSQL = strSQL &" 	GROUP  BY m.cdcliente,					"
strSQL = strSQL &" 			  c.dscliente,					"
strSQL = strSQL &" 			  m.cdproducto,					"
strSQL = strSQL &" 			  p.dsproducto					"
Call GF_BD_Puertos(pto, rs, "OPEN", strSQL)
Set obtenerSQLOleoducto = rs
End function 
'------------------------------------------------------------------------------------------
Function obtenerSQLRecargas(pMyFecha,pto,pLstClientes)
Dim strSQL, rs
strSQL = strSQL &"	  SELECT   TN.cdproducto,"
strSQL = strSQL &"			   P.dsproducto, "
strSQL = strSQL &"			   TN.cdcliente,"
strSQL = strSQL &"			   C.dscliente,"
strSQL = strSQL &"			   TN.NETO as Cant1,"
strSQL = strSQL &"			   TN.cant as Cant2 "               
strSQL = strSQL &"		FROM("
strSQL = strSQL &"			SELECT BRUTO.cdproducto,"
strSQL = strSQL &"				   BRUTO.cdcliente,"
strSQL = strSQL &"				   ( BRUTO.kilos-tara.kilos ) NETO,"
strSQL = strSQL &"				   BRUTO.cant  CANT"
strSQL = strSQL &"			FROM"
'					 Totaliza Kilos Bruto por Product/Cliente
strSQL = strSQL &"			(SELECT HC.cdproducto,"
strSQL = strSQL &"					HCC.cdcliente,"
strSQL = strSQL &"					Sum(HCP.vlpesada) KILOS,"
strSQL = strSQL &"					COUNT(*)          CANT"
strSQL = strSQL &"			FROM   (SELECT *"
strSQL = strSQL &"					 FROM   hcamiones"
strSQL = strSQL &"				 WHERE  dtcontable = '" & pMyFecha & "'"
strSQL = strSQL &"						   AND cdestado NOT IN ( 12, 7, 14 )) HC"
strSQL = strSQL &"					INNER JOIN hcamionescarga HCC"
strSQL = strSQL &"							ON HC.dtcontable = HCC.dtcontable"
strSQL = strSQL &"							   AND HC.idcamion = HCC.idcamion"
strSQL = strSQL &"					INNER JOIN hpesadascamion HCP"
strSQL = strSQL &"							ON HC.dtcontable = HCP.dtcontable"
strSQL = strSQL &"							   AND HC.idcamion = HCP.idcamion"
strSQL = strSQL &"			WHERE  HCP.sqpesada = (SELECT Max(sqpesada)"
strSQL = strSQL &"									FROM   hpesadascamion"
strSQL = strSQL &"									WHERE  dtcontable = HCP.dtcontable"
strSQL = strSQL &"										   AND idcamion = HCP.idcamion"
strSQL = strSQL &"										   AND cdpesada = 1)"
if (pLstClientes <> "") then strSQL = strSQL &" and HCC.cdcliente = " & pLstClientes 
strSQL = strSQL &"			GROUP  BY HC.cdproducto,"
strSQL = strSQL &"					   HCC.cdcliente) BRUTO"
strSQL = strSQL &"			INNER JOIN			"
'					 Totaliza Kilos Tara por Product/Cliente
strSQL = strSQL &"			(SELECT HC.cdproducto,"
strSQL = strSQL &"					HCC.cdcliente,"
strSQL = strSQL &"					Sum(HCP.vlpesada) KILOS"
strSQL = strSQL &"			FROM   (SELECT *"
strSQL = strSQL &"					 FROM   hcamiones"
strSQL = strSQL &"					 WHERE  dtcontable = '" & pMyFecha & "'"
strSQL = strSQL &"							AND cdestado NOT IN ( 12, 7, 14 )) HC"
strSQL = strSQL &"					INNER JOIN hcamionescarga HCC"
strSQL = strSQL &"							ON HC.dtcontable = HCC.dtcontable"
strSQL = strSQL &"							   AND HC.idcamion = HCC.idcamion"
strSQL = strSQL &"					INNER JOIN hpesadascamion HCP"
strSQL = strSQL &"							ON HC.dtcontable = HCP.dtcontable"
strSQL = strSQL &"							   AND HC.idcamion = HCP.idcamion"
strSQL = strSQL &"			WHERE  HCP.sqpesada = (SELECT Max(sqpesada)"
strSQL = strSQL &"									FROM   hpesadascamion"
strSQL = strSQL &"									WHERE  dtcontable = HCP.dtcontable"
strSQL = strSQL &"										   AND idcamion = HCP.idcamion"
strSQL = strSQL &"										   AND cdpesada = 2)"
if (pLstClientes <> "") then strSQL = strSQL &" and HCC.cdcliente = " & pLstClientes 
strSQL = strSQL &"			GROUP  BY HC.cdproducto,"
strSQL = strSQL &"					   HCC.cdcliente) TARA"
strSQL = strSQL &"					ON BRUTO.cdproducto = TARA.cdproducto"
strSQL = strSQL &"					   AND BRUTO.cdcliente = TARA.cdcliente) as TN"
strSQL = strSQL &"		INNER JOIN dbo.Clientes C"
strSQL = strSQL &"			ON C.cdcliente = TN.cdcliente"
strSQL = strSQL &"		INNER JOIN dbo.productos P	"
strSQL = strSQL &"			ON P.cdproducto = TN.cdproducto"
strSQL = strSQL &"		ORDER  BY TN.cdcliente, TN.cdproducto "

'response.write strSQL
'response.end
Call GF_BD_Puertos(pto, rs, "OPEN", strSQL)
Set obtenerSQLRecargas = rs		
End function
'---------------------------------------------------------------------------
Function obtenerSQLAjusteCredDebi(pMyFecha,pto,pLstClientes)
Dim strSQL, rs
strSQL ="Select T.cdcliente, T.cdproducto, c.dscliente, p.dsproducto,   " & _
		"       Sum(Cant1) as Cant1, Sum(Cant2) as Cant2                " & _
        "FROM															" & _
		"((SELECT cdcliente,											" & _
		" 		cdproducto,												" & _
		" 		vlkilosajuste AS Cant1,									" & _
		" 		0 AS Cant2												" & _
		"FROM dbo.excreditosdebitos 								" & _
		"WHERE  dtcontable = '" & pMyFecha & "'							"
        if (pLstClientes <> "") then strSQL = strSQL & " and cdcliente in ("& pLstClientes & ")"
		strSQL = strSQL & " AND vlkilosajuste < 0)									" & _
		"UNION															" & _
		"(SELECT cdcliente,												" & _
		" 		cdproducto,												" & _
		" 		0 AS Cant1,												" & _
		" 		vlkilosajuste AS Cant2									" & _
		"FROM dbo.excreditosdebitos 								" & _
		"WHERE  dtcontable = '" & pMyFecha & "'							"
        if (pLstClientes <> "") then strSQL = strSQL & " and cdcliente in ("& pLstClientes & ")"
		strSQL = strSQL & " AND vlkilosajuste >= 0)) T								" & _
		"LEFT JOIN dbo.clientes c ON c.cdcliente = T.cdcliente		" & _
		"LEFT JOIN dbo.productos p ON p.cdproducto = T.cdproducto	" & _
		"GROUP BY T.cdcliente, T.cdproducto, c.dscliente, p.dsproducto  " & _
        "ORDER BY T.cdcliente,	T.cdproducto                            "
'Response.Write strSQL
'Response.End
Call GF_BD_Puertos(pto, rs, "OPEN", strSQL)
Set obtenerSQLAjusteCredDebi = rs			
End function 
'---------------------------------------------------------------------------
Function obtenerSQLAjusteStock(pMyFecha,pto,pLstClientes)
Dim strSQL, rs
strSQL =" 			SELECT  a.cdcliente,				"
strSQL = strSQL &" 			dscliente,					"
strSQL = strSQL &" 			a.cdproducto,				"
strSQL = strSQL &" 			dsproducto,					"
strSQL = strSQL &" 			Sum(a.vlkilos) AS Kilos 	"
strSQL = strSQL &"  FROM dbo.exajstkopecli a		"
strSQL = strSQL &" 	LEFT JOIN dbo.clientes c		"
strSQL = strSQL &" 		ON c.cdcliente = a.cdcliente	"
strSQL = strSQL &" 	LEFT JOIN dbo.productos p		"
strSQL = strSQL &" 		ON p.cdproducto = a.cdproducto	"
strSQL = strSQL &" 	GROUP  BY a.dtcontable,				"
strSQL = strSQL &" 			  a.cdcliente,				"
strSQL = strSQL &" 			  dscliente,				"
strSQL = strSQL &" 			  a.cdproducto,				"			
strSQL = strSQL &" 			  dsproducto				"
strSQL = strSQL &" HAVING a.dtcontable = '"&pMyFecha&"'"
if (pLstClientes <> "") then strSQL = strSQL & " and a.cdcliente in ("& pLstClientes & ")"
'response.write strSQL
'response.end
Call GF_BD_Puertos(pto, rs, "OPEN", strSQL)
Set obtenerSQLAjusteStock = rs			
End function
'---------------------------------------------------------------------------
Function ArmarColumnaAjusteStock(pMyFecha,ppto,campo,pLstClientes)
Dim strSQL
strSQL = "			 SELECT DISTINCT "&campo 
strSQL = strSQL &" 	 FROM   dbo.exajstkopecli a		"
strSQL = strSQL &" 	 LEFT JOIN dbo.clientes c			"
strSQL = strSQL &" 		ON c.cdcliente = a.cdcliente		"
strSQL = strSQL &" 	 LEFT JOIN dbo.productos p			"
strSQL = strSQL &" 		ON p.cdproducto = a.cdproducto		"
strSQL = strSQL &" 	 WHERE a.dtcontable = '"&pMyFecha&"'	"
if (pLstClientes <> "") then  strSQL = strSQL & " and a.cdcliente in ("& pLstClientes & ")"
strSQL = strSQL &" 	 GROUP  BY "&campo
Call GF_BD_Puertos(pto, rs, "OPEN", strSQL)
Set ArmarColumnaAjusteStock = rs			
End function
'---------------------------------------------------------------------------
Function obtenerSQLEmbarcados(pMyFecha,pto,pLstClientes)
Dim strSQL, rs, myWhere, myTablaA,myTablaB
strSQL = strSQL &"	SELECT DISTINCT ce.cdaviso,					"
strSQL = strSQL &"					ce.cdproducto,				"
strSQL = strSQL &"                p.dsproducto,					"
strSQL = strSQL &"                ce.cdcliente,					"
strSQL = strSQL &"                c.dscliente,					"
strSQL = strSQL &"                em.cdbuque,					"
strSQL = strSQL &"                b.dsbuque,					"
strSQL = strSQL &"                Sum(ce.vlkilos) AS Kilos		"
strSQL = strSQL &"FROM   dbo.cargasembarque ce				"
strSQL = strSQL &"       LEFT JOIN dbo.clientes c			"
strSQL = strSQL &"              ON c.cdcliente = ce.cdcliente	"
strSQL = strSQL &"       LEFT JOIN dbo.productos p			"
strSQL = strSQL &"              ON p.cdproducto = ce.cdproducto,"
strSQL = strSQL &"       dbo.embarques em					"
strSQL = strSQL &"       LEFT JOIN dbo.buques b			"
strSQL = strSQL &"              ON b.cdbuque = em.cdbuque		"
strSQL = strSQL &"WHERE  ce.cdaviso = em.cdaviso				"
strSQL = strSQL &"       AND ce.dtcarga = '"&pMyFecha&"'		"
if (pLstClientes <> "") then strSQL = strSQL &" AND ce.cdcliente in ("& pLstClientes & ")"
strSQL = strSQL &"GROUP  BY ce.cdaviso,							"
strSQL = strSQL &"          ce.cdproducto,						"
strSQL = strSQL &"          p.dsproducto,						"
strSQL = strSQL &"          ce.cdcliente,						"
strSQL = strSQL &"          c.dscliente,						"
strSQL = strSQL &"          em.cdbuque,							"
strSQL = strSQL &"          b.dsbuque							"
strSQL = strSQL &"	UNION										"	
strSQL = strSQL &"SELECT DISTINCT ce.cdaviso,					"
strSQL = strSQL &"                ce.cdproducto,				"
strSQL = strSQL &"                p.dsproducto,					"
strSQL = strSQL &"                ce.cdcliente,					"	
strSQL = strSQL &"                c.dscliente,					"
strSQL = strSQL &"                em.cdbuque,					"
strSQL = strSQL &"                b.dsbuque,					"
strSQL = strSQL &"                Sum(ce.vlkilos) AS Kilos		"
strSQL = strSQL &"FROM   dbo.hcargasembarque ce			"
strSQL = strSQL &"       LEFT JOIN dbo.clientes c			"
strSQL = strSQL &"              ON c.cdcliente = ce.cdcliente	"
strSQL = strSQL &"       LEFT JOIN dbo.productos p			"
strSQL = strSQL &"              ON p.cdproducto = ce.cdproducto,"
strSQL = strSQL &"       dbo.hembarques em					"
strSQL = strSQL &"       LEFT JOIN dbo.buques b			"
strSQL = strSQL &"              ON b.cdbuque = em.cdbuque		"
strSQL = strSQL &"WHERE  ce.dtcontable = em.dtcontable			"
strSQL = strSQL &"       AND ce.cdaviso = em.cdaviso			"
strSQL = strSQL &"       AND ce.dtcarga = '"&pMyFecha&"'		"
if (pLstClientes <> "") then strSQL = strSQL &" AND ce.cdcliente in ("& pLstClientes & ")"
strSQL = strSQL &"GROUP  BY ce.cdaviso,							"
strSQL = strSQL &"          ce.cdproducto,						"
strSQL = strSQL &"          p.dsproducto,						"
strSQL = strSQL &"          ce.cdcliente,						"
strSQL = strSQL &"          c.dscliente,						"
strSQL = strSQL &"          em.cdbuque,							"
strSQL = strSQL &"          b.dsbuque  							" 
'response.write strSQL
'response.end
Call GF_BD_Puertos(pto, rs, "OPEN", strSQL)
Set obtenerSQLEmbarcados = rs			
End function
'---------------------------------------------------------------------------
Function obtenerTopeKilos(pCdAviso, pProducto, pcliente, pto)
	Dim strSQL,rs,rtrn
	rtrn = 0  
	strSQL = " Select sum(t.Kilos) as kilos from (select sum(cc.vlkilos) as kilos from dbo.HCoordinadosCoordinadorasEmb cc"
	strSQL = strSQL & " Where cc.cdaviso = " & pCdAviso
	strSQL = strSQL & " and cc.cdcoordinado = " & pcliente & " and cc.cdproducto = " & pProducto
	strSQL = strSQL & " Union "
	strSQL = strSQL & " select sum(cc.vlkilos) as kilos from dbo.CoordinadosCoordinadorasEmb cc"
	strSQL = strSQL & " Where cc.cdaviso = " & pCdAviso   
	strSQL = strSQL & " and cc.cdcoordinado = " & pcliente & " and cc.cdproducto = " & pProducto
	strSQL = strSQL & " and cc.cdaviso not in (select cdaviso from dbo.HCoordinadosCoordinadorasEmb "
	strSQL = strSQL & " Where cdaviso = cc.cdaviso And cdempresacoordinadora = cc.cdempresacoordinadora  "
	strSQL = strSQL & " and cdcoordinado = cc.cdcoordinado and cdProducto = cc.cdproducto)) as t"

	Call GF_BD_Puertos(pto, rs, "OPEN", strSQL)
	if not isNull(rs("Kilos")) then rtrn = rs("Kilos")	
	obtenerTopeKilos = CDbl(rtrn)
End function
'--------------------------------------------------------------------------
Function obtenerKilosAcumulados(cdAviso, pProducto, pto, pCodCliente, pFecha)
	Dim strSQL,rs,rtrn,campoAnioMesDia,campoDiaMes,campoMes,campoDia,v_diferencia
	rtrn = 0 
	
	strSQL = "Select sum(t.Kilos) as KILOS from "
	strSQL = strSQL & " (select sum(ce.vlkilos) as kilos from dbo.HCargasEmbarque ce "
	strSQL = strSQL & " Where ce.cdaviso = " & cdAviso & " And ce.cdproducto = " & pProducto
	strSQL = strSQL & " and ce.cdcliente = " & pCodCliente
	strSQL = strSQL & " and ce.dtcarga between '2000-01-01' and '" & pFecha & "'"
	strSQL = strSQL & " Union "
	strSQL = strSQL & " select sum(ce.vlkilos) as kilos from dbo.CargasEmbarque ce"
	strSQL = strSQL & " Where ce.cdaviso = " & cdAviso & " And ce.cdproducto = " & pProducto
	strSQL = strSQL & " and ce.cdcliente = " & pCodCliente
	strSQL = strSQL & " and ce.dtcarga between '2000-01-01' and '" & pFecha & "'"
	strSQL = strSQL & " and ce.cdaviso not in (select cdaviso from dbo.HCargasEmbarque "
	strSQL = strSQL & " Where cdaviso  = ce.cdaviso And cdproducto = ce.cdproducto "
	strSQL = strSQL & " and cdempresa = ce.cdempresa and cdcliente = ce.cdcliente)) as t"	
	Call GF_BD_Puertos(pto, rs, "OPEN", strSQL)

	if not isNull(rs("KILOS")) then rtrn = rs("KILOS")	
	
	obtenerKilosAcumulados = CDbl(rtrn)
End function
'--------------------------------------------------------------------------
Function armarSeccionDescargas(pto, rsDatos, rsColumnas, rsFilas, fecha, titulo, col1, col2)		
	Dim vectCol, vecTotalesCant1, vecTotalesCant2, z

    if (not rsColumnas.eof) then     	
        call writeXLS("<table> <tr> <td  colspan='3' class='xls_align_left'>"& titulo &"</td> </tr> </table>")
        call writeXLS("<table style='width:150%; font-size:9;' class='xls_border_left' >")
        call writeXLS(" <tr style='background-color:#BABABA;' whidt='200%'>")
        call writeXLS("     <td width='40px' colspan='1' rowSpan='2' class='xls_align_center'>CLIENTE</td>")
        z =0			 
        
        Redim vectCol(rsColumnas.recordcount)
	    Redim  vecTotalesCant1(rsColumnas.recordcount)
	    Redim  vecTotalesCant2(rsColumnas.recordcount)				
    	
	    while(not rsColumnas.eof)    
            call writeXLS("<td  colspan='2' class='xls_align_center'>"& rsColumnas("dsproducto") & "</td>")
            vectCol(z)= rsColumnas("cdproducto")
		    z = z + 1
		    rsColumnas.movenext
	    wend
        call writeXLS("</tr>")
        call writeXLS("<tr>")
        rsColumnas.MoveFirst()
        while(not rsColumnas.eof)    
            call writeXLS("<td class='xls_align_center'>"& col1 &"</td>")
            call writeXLS("<td  class='xls_align_center'>"& col2 &"</td>")
            rsColumnas.MoveNext()
	    wend
        call writeXLS("</tr>")
        while(not rsFilas.eof)
            call writeXLS("<tr><td align='left'>"& rsFilas("cdcliente") & " - " & rsFilas("dscliente") &"</td>")
            seguir= true	
		    i = 0
		    'Recorro las descargas mientras el cliente sea el indicado por la fila en proceso.
		    while((not rsDatos.eof) and seguir)
                if(cdbl(rsFilas("cdcliente")) = cdbl(rsDatos("cdcliente")))then
			        if (cdbl(vectCol(i)) = cdbl(rsDatos("cdproducto")))then        
                        call writeXLS("<td align='right'>")
                        if not isNull(rsDatos("Cant1")) then 
					        call writeXLS(cdbl(rsDatos("Cant1")))
						    vecTotalesCant1(i) = cdbl(vecTotalesCant1(i)) + cdbl(rsDatos("Cant1"))
					    end if
                        call writeXLS("</td>")
                        call writeXLS("<td align='right'>")
                        if not isNull(rsDatos("Cant2")) then
                            call writeXLS(cdbl(rsDatos("Cant2")))
						    vecTotalesCant2(i) = cdbl(vecTotalesCant2(i)) + cdbl(rsDatos("Cant2"))							
					    end if
                        call writeXLS("</td>")
                        rsDatos.MoveNext()
                	    i = i + 1
				    else	
                        'Si el producto no coincide, solo queda que el producto en la tabla de descargas sea
					    'mayor al de la columna dado que NUNCA puede ocurrir que un cliente tenga descrgas de un producto sin columna.
                        call writeXLS("<td align='right'></td>")
                        call writeXLS("<td align='right' ></td>")
                        i = i+1
                        'No deberia ser necesario dado que siempre los clientes tienen al menos una descarga para mostrar, pero por seguridad se agrega el control.
					    if (i > UBound(vectCol)) then	seguir = false
                    end if
			    else					
			        seguir = false
			    end if					
		    wend
		    rsFilas.MoveNext()
            call writeXLS("</tr>")
        wend
        call writeXLS("<tr style='background-color:#EDEDED;' class='xls_align_center'>")
        call writeXLS("<td colspan='1' class='xls_align_center'>TOTAL</td>")
        for i = LBound(vecTotalesCant1) to UBound(vecTotalesCant1)-1
            call writeXLS("<td class='xls_align_right'>"& vecTotalesCant1(i) &"</td>")
            call writeXLS("<td class='xls_align_right'>"& vecTotalesCant2(i) &"</td>")
        next
        call writeXLS("</tr>")
        call writeXLS("</table><br><br><br>")
    end if    
End Function
'--------------------------------------------------------------------------
Function armarSaldos(pto, rsDatos, fecha, pLstClientes)
	Dim vectCol, vecTotalesProducto, vecCamionesProducto, z
    call writeXLS("<table style='width:150%; font-size:9;' class='xls_border_left'>")
    call writeXLS("<tr style='background-color:#BABABA;' whidt='200%'>")
    call writeXLS("<td whidt='40px' colspan='1' class='xls_align_center'>CLIENTE</td>")
    z =0
    Set rsColumnas = armarColumnasProductos(fecha, pto, pLstClientes)
	if (not rsColumnas.eof) then
	    Redim vectCol(rsColumnas.recordcount)
		Redim  vecTotalesProducto(rsColumnas.recordcount)
	end if
	while(not rsColumnas.eof)
        call writeXLS("<td  colspan='1' class='xls_align_center'>"& rsColumnas("dsproducto") &"</td>")
        vectCol(z)= rsColumnas("cdproducto")
		z = z + 1
		rsColumnas.movenext
	wend

    call writeXLS("</tr>")
    Set rsFilas = armarFilasClientes(fecha,pto,pLstClientes)
    while(not rsFilas.eof)
        call writeXLS("<tr><td align='left'>"& rsFilas("cdcliente") & " - " & rsFilas("dscliente") &"</td>")
        seguir= true	
        i = 0
	    'Recorro las descargas mientras el cliente sea el indicado por la fila en proceso.
	    while((not rsDatos.eof) and seguir)
	        'Si es el cliente y es el producto, muestro los kilos y sumo para el total de la columna.
		    if(cdbl(rsFilas("cdcliente")) = cdbl(rsDatos("cdcliente")))then 				
		        if (cdbl(vectCol(i)) = cdbl(rsDatos("cdproducto")))then
                    call writeXLS("<td align='right' >"& cdbl(rsDatos("Kilos")) &"</td>")
                    vecTotalesProducto(i) = cdbl(vecTotalesProducto(i)) + cdbl(rsDatos("Kilos"))
				    rsDatos.MoveNext()
				    i = i+1
			    else	
			        'Si el producto no coincide, solo queda que el producto en la tabla de descargas sea 
				    'mayor al de la columna dado que NUNCA puede ocurrir que un cliente tenga descrgas de un producto sin columna.					
                    call writeXLS("<td align='right' ></td>")
                    i = i+1	
				    'No deberia ser necesario dado que siempre los clientes tienen al menos una descarga para mostrar, pero por seguridad se agrega el control.
				    if (i > UBound(vectCol)) then	seguir = false
			    end if
		    else					
		        seguir = false
		    end if					
	    wend
	    rsFilas.MoveNext()
        call writeXLS("</tr>")
    wend

    call writeXLS("<tr style='background-color:#EDEDED;' class='xls_align_center'>")
    call writeXLS("<td colspan='1' class='xls_align_center'>TOTAL</td>")
    for i = LBound(vecTotalesProducto) to UBound(vecTotalesProducto)-1
        call writeXLS("<td class='xls_align_right'>"& vecTotalesProducto(i) &"</td>")
    next
    call writeXLS("</tr></table><br><br><br>")
End Function
'-----------------------------------------------------------------------------------------------------------------------------
Function obtenerFechaInicial(pMyFecha,pto,pLstClientes)
    Dim strSQL,rs
    strSQL = "Select ((Year(T.DTCONTABLE) * 10000) + (Month(T.DTCONTABLE) * 100) + Day(T.DTCONTABLE)) AS FECHA "&_
             "from ( Select MAX(DTCONTABLE) AS DTCONTABLE from dbo.excuentcorrientes where DTCONTABLE < '"& pMyFecha &"'"
    if (pLstClientes <> "") then strSQL = strSQL & " and CDCLIENTE in ("& pLstClientes & ")"
    strSQL = strSQL & " ) T"
   
    Call GF_BD_Puertos(pto, rs, "OPEN", strSQL)
    if (not rs.Eof) then obtenerFechaInicial = rs("FECHA")
End function
'--------------------------------------------------------------------------------------------------------
Function armarReporteTerminalXLS(pPto, fname, fnActual, pLstClientes,pMode)
    Dim db2FechaActual, db2FechaMenos,fecMenosD,fecMenosM,fecMenosA,fechaMenos, fechaActual
    g_strPuerto = pPto    
    db2FechaActual = fnActual
    fechaActual = GF_FN2DTCONTABLE(fnActual)
    db2FechaMenos = obtenerFechaInicial(fechaActual, pPto, pLstClientes)
    fecMenosD = Right(db2FechaMenos, 2)
    fecMenosM = Mid(db2FechaMenos, 5, 2)
    fecMenosA = Left(db2FechaMenos, 4)
    fechaMenos  = fecMenosA &"-"& fecMenosM &"-"& fecMenosD    
    xls_mode = pMode
    armarReporteTerminalXLS = GF_createXLS(fname)
    call writeXLS("<html>")
    call writeXLS("<head>")
    call writeXLS("<style type='text/css'>")
    call writeXLS(".xls_border_left {border-color:#666666;border-style:solid;border-width:thin;}")
    call writeXLS(".xls_align_center {border-color:#666666;border-style:solid; border-width:thin;text-align: center;}")
    call writeXLS(".xls_align_right { border-color:#666666; border-style:solid; border-width:thin;text-align: right;}")
    call writeXLS(".xls_precioUC_tablaArticulos{BACKGROUND-COLOR: #ffff80;border-color:#666666; border-style:solid; border-width:thin;}")
    call writeXLS("</style> </head>")
    call writeXLS("<body>")
    call writeXLS("<table class=xls_border_left style=background-color:#EDEDED; font-weight:bold>")
    call writeXLS(" <tr>")
    call writeXLS("     <td>" & "ACTI AR" &"</td>")
    call writeXLS("     <td></td><td></td>")
    call writeXLS(" </tr>")
    call writeXLS(" <tr>")
    call writeXLS(" <td>"& "PTO:"& getNombrePuerto(pPto) &"</td>")
    call writeXLS(" <td align=right colspan=3>Fecha : " & GF_FN2DTE(session("MmtoSistema")) &"</td>")
    call writeXLS(" </tr></table><br><br>")
    Call armarSaldoInicial(fechaMenos, pPto, pLstClientes,db2FechaActual)
    
    Call armarDescargaCamiones(fechaActual, pPto, pLstClientes, db2FechaActual)

    Call armarDescargaVagones(fechaActual, pPto, pLstClientes, db2FechaActual)

    Call armarDescargaOleoducto(fechaActual, pPto, pLstClientes, db2FechaActual)

    Call armarEmbarcados(fechaActual,pPto, pLstClientes, db2FechaActual)

    Call armarTransferencia(fechaActual,pPto, pLstClientes, db2FechaActual)

    Call armarRecargas(fechaActual,pPto, pLstClientes, db2FechaActual)

    Call armarCreditoDebito(fechaActual,pPto, pLstClientes, db2FechaActual)

    Call armarAjusteStock(fechaActual,pPto, pLstClientes, db2FechaActual)

    Call armarSaldoFinal(fechaActual,pPto, pLstClientes, db2FechaActual)

    call writeXLS("</body></html>")
    Call closeXLS()       
End Function
'--------------------------------------------------------------------------------------------------------
Function armarSaldoInicial(pFechaMenos, pPto, pLstClientes,pDb2FechaActual)
    Dim rsSalIni
    Set rsSalIni = ObtenerSQLSaldoInicial(pFechaMenos, pPto, pLstClientes)
    if(not rsSalIni.eof)then
        call writeXLS("<table > <tr> <td colspan=3>Movimiento de la Terminal del "& GF_FN2DTE(pDb2FechaActual)&"</td> </tr>	")
        call writeXLS("     <tr><td whidt=40px colspan=3 class=xls_align_left>Posicion del "& GF_FN2DTE(pDb2FechaActual) &" 00:00 Horas</td></tr> ")
        call writeXLS("</table> ")
        Call armarSaldos(pPto, rsSalIni, pFechaMenos, pLstClientes)
    end if
End function
'--------------------------------------------------------------------------------------------------------
Function armarDescargaCamiones(pFechaActual, pPto, pLstClientes, pDb2FechaActual)
    Dim rsCamDes,rsColumnas,rsFilas
    Set rsCamDes = ObtenerSQLDescargaCamiones(pFechaActual, pPto, pLstClientes)
    if(not rsCamDes.eof)then
	    Set rsColumnas	= armarColumnaClientesCamiones(pFechaActual,pPto,"TF.cdproducto,dsproducto,TF.dtcontable",pLstClientes)
	    Set rsFilas		= armarColumnaClientesCamiones(pFechaActual,pPto, "TF.cdcliente,dscliente,TF.dtcontable",pLstClientes)
	    Call armarSeccionDescargas(pPto, rsCamDes, rsColumnas, rsFilas, "", "Descargas de Camiones del " & GF_FN2DTE(pDb2FechaActual), "Kilos", "Camiones")
    end if
End Function
'--------------------------------------------------------------------------------------------------------
Function armarDescargaVagones(pFechaActual, pPto, pLstClientes, pDb2FechaActual)
    Dim rsVagDes,rsColumnas,rsFilas
    Set rsVagDes  = ObtenerSQLDescargaVagones(pFechaActual, pPto, pLstClientes, "")
    if(not rsVagDes.eof)then
	    Set rsColumnas	= ObtenerSQLDescargaVagones(pFechaActual, pPto, pLstClientes, "T1.cdproducto, T1.dsproducto")
	    Set rsFilas	= ObtenerSQLDescargaVagones(pFechaActual, pPto, pLstClientes, " T1.cdcliente, T1.dscliente ")	    
	    Call armarSeccionDescargas(pPto, rsVagDes, rsColumnas, rsFilas, "", "Descargas de Vagones del " & GF_FN2DTE(pDb2FechaActual), "Kilos", "Vagones")
    end if
End function
'--------------------------------------------------------------------------------------------------------
Function armarDescargaOleoducto(pFechaActual, pPto, pLstClientes, pDb2FechaActual)
    Dim rsOleo,rsColumnas,rsFilas
    Set rsOleo = obtenerSQLOleoducto(pFechaActual,pPto, pLstClientes)
    if(not rsOleo.eof)then
	    Set rsColumnas	= ArmarColumnaOleoducto(pFechaActual,pPto," m.cdproducto, p.dsproducto ",pLstClientes)
	    Set rsFilas		= ArmarColumnaOleoducto(pFechaActual,pPto," m.cdcliente, c.dscliente ",pLstClientes)
	    Call armarSeccionDescargas(pPto, rsOleo, rsColumnas, rsFilas, "", "Oleoductos del " & GF_FN2DTE(pDb2FechaActual), "Kilos", "Cantidad")
    end if
End function
'--------------------------------------------------------------------------------------------------------
Function armarEmbarcados(pFechaActual,pPto, pLstClientes, pDb2FechaActual)
    Dim rsBuques,rsColumnas,rsFilas
    Dim myCargaSaldo,myCargaAcumulada
    
    set rsBuques = obtenerSQLEmbarcados(pFechaActual,pPto,pLstClientes)
    if not rsBuques.eof then 	
	
	While (not rsBuques.eof) 	
		cdavisoOld = rsBuques("cdaviso")	
		sumTotalKilosDia = 0
		sumTotalKilosSaldo = 0
		sumTotalKilosAcumulado = 0
		sumTotalKilosEstimado = 0

        call writeXLS(" <table> <tr> <td  colspan=3 class=xls_align_left>Buque: "&rsBuques("cdbuque")&"-"&rsBuques("dsbuque")&"</td> </tr></table>")
        call writeXLS(" <table class=xls_border_left style='width:100%; font-size:9;' ><tr style='background-color:#BABABA;'>")
        call writeXLS(" <td width='40%' colspan='1' class='xls_align_center'>PRODUCTO</td>")
        call writeXLS(" <td width='40%' colspan='1' class='xls_align_center'>CLIENTE</td>")
        call writeXLS(" <td width='10%' colspan='1' class='xls_align_center'>CARGA ESTIMADA</td>")
        call writeXLS(" <td width='10%' colspan='1' class='xls_align_center'>CARGA DIA</td>")
        call writeXLS(" <td width='10%' colspan='1' class='xls_align_center'>CARGA ACUMULADA</td>")
        call writeXLS(" <td width='10%' colspan='1' class='xls_align_center'>SALDO CARGA</td></tr>")
        cdavisoAct = cdavisoOld
		while ((not rsBuques.eof) and (cdavisoOld = cdavisoAct))
			productoOld = rsBuques("cdproducto")
			sumProductoKilosDia = 0
			sumProductoKilosSaldo = 0
			sumProductoKilosAcumulado = 0
			sumProductoKilosEstimado = 0
            call writeXLS("<tr><td>"& rsBuques("cdproducto") & "-" & rsBuques("dsproducto") &"</td></tr> ")
            productoAct = productoOld
			while ((not rsBuques.eof) and (cdavisoOld = cdavisoAct) and (productoOld = productoAct))

                call writeXLS("<tr><td></td><td width='40%' colspan='1' align='left'>"& rsBuques("cdcliente")&"-"&rsBuques("dscliente")&"</td> ")
                myCargaEstimada = obtenerTopeKilos(rsBuques("cdaviso"),rsBuques("cdproducto"),rsBuques("cdcliente"),pPto)
                call writeXLS("<td width='10%' colspan='1' align='right'>"& myCargaEstimada &"</td> ")
                myCargaDia = CDbl(rsBuques("kilos"))
			    if  (isNull(rsBuques("kilos"))) then myCargaDia=0
                call writeXLS("<td width='10%' colspan='1' align='right'>"& myCargaDia & "</td>")
                myCargaAcumulada = obtenerKilosAcumulados(rsBuques("cdaviso"),rsBuques("cdproducto"),pPto,rsBuques("cdcliente"), pDb2FechaActual)
                call writeXLS("<td width='10%' colspan='1' align='right'>"&myCargaAcumulada&"</td>")
                myCargaSaldo = myCargaEstimada - myCargaAcumulada
                call writeXLS("<td width='10%' colspan='1' align='right'>"& myCargaSaldo &"</td></tr>")
                sumProductoKilosEstimado	= sumProductoKilosEstimado + myCargaEstimada
			    sumProductoKilosDia			= sumProductoKilosDia +	myCargaDia			 
			    sumProductoKilosAcumulado	= sumProductoKilosAcumulado	+ myCargaAcumulada
			    sumProductoKilosSaldo		= sumProductoKilosSaldo + myCargaSaldo
			    rsBuques.MoveNext()
			    if (not rsBuques.eof) then
			        cdavisoAct = rsBuques("cdaviso")
				    productoAct = rsBuques("cdproducto")
			    end if
			wend	
            'Sumo los totales generales del buque
			sumTotalKilosEstimado	= sumTotalKilosEstimado + sumProductoKilosEstimado
			sumTotalKilosDia		= sumTotalKilosDia + sumProductoKilosDia			 
			sumTotalKilosAcumulado	= sumTotalKilosAcumulado + sumProductoKilosAcumulado
			sumTotalKilosSaldo		= sumTotalKilosSaldo + sumProductoKilosSaldo
				
            call writeXLS("<tr style='background-color:#EDEDED;' class='xls_align_center'>")
            call writeXLS(" <td></td> <td style='background-color:#EDEDED;'  class='xls_align_center'>Subtotal</td>")
            call writeXLS("<td class='xls_align_right'>"& sumProductoKilosEstimado &"</td>")
            call writeXLS("<td class='xls_align_right'>"& sumProductoKilosDia & "</td>")
            call writeXLS("<td class='xls_align_right'>"& sumProductoKilosAcumulado &"</td>")
            call writeXLS("<td class='xls_align_right'>"& sumProductoKilosSaldo & "</td></tr>	")
        wend

        call writeXLS("<tr><td colspan='6'></td></tr>")
        call writeXLS("<tr style='background-color:#EDEDED;' class='xls_align_center'>")
        call writeXLS("<td colspan='2' width='40%' style='background-color:#EDEDED;' class='xls_align_center'>TOTAL</td>")
        call writeXLS("<td width='10%' class='xls_align_right'>"& sumTotalKilosEstimado & "</td>")
        call writeXLS("<td width='10%' class='xls_align_right'>"& sumTotalKilosDia & "</td>")
        call writeXLS("<td width='10%' class='xls_align_right'>"& sumTotalKilosAcumulado & "</td>")
        call writeXLS("<td width='10%' class='xls_align_right'>"& sumTotalKilosSaldo & "</td> ")
        call writeXLS("</tr></table><br>")
    wend	
end if
End function
'--------------------------------------------------------------------------------------------------------
Function armarTransferencia(pFechaActual,pPto, pLstClientes, pDb2FechaActual)
    Dim rsTransf,rsColumnas,rsFilas
    set rsTransf = obtenerSQLTransferencia(pFechaActual,pPto,pLstClientes) 
    if(not rsTransf.eof)then
	    Set rsColumnas	= ArmarColumnaTransferencia(pFechaActual,pPto,"m.cdproducto, dsproducto",pLstClientes)
	    Set rsFilas		= ArmarColumnaTransferencia(pFechaActual, pPto, "m.cdcliente, dscliente",pLstClientes)		
	    Call armarSeccionDescargas(pPto, rsTransf, rsColumnas, rsFilas, "", "Transferencias del " & GF_FN2DTE(pDb2FechaActual), "Recibidos", "Transferidos")
    end if
End function
'--------------------------------------------------------------------------------------------------------
Function armarRecargas(pFechaActual,pPto, pLstClientes, pDb2FechaActual)
    Dim rsCamRec,rsColumnas,rsFilas
    Set rsCamRec  = obtenerSQLRecargas(pFechaActual, pPto,pLstClientes)
    if(not rsCamRec.eof)then
    	Set rsColumnas	= ArmarColumnaRecargas(pFechaActual,pPto,"T2.cdproducto, T2.dsproducto",pLstClientes)
    	Set rsFilas		= ArmarColumnaRecargas(pFechaActual, pPto, "T2.cdcliente, T2.dscliente",pLstClientes)
    	Call armarSeccionDescargas(pPto, rsCamRec, rsColumnas, rsFilas, "", "Recargas de Camiones del " & GF_FN2DTE(pDb2FechaActual), "Kilos", "Camiones")
    end if
End function
'--------------------------------------------------------------------------------------------------------
Function armarCreditoDebito(pFechaActual,pPto, pLstClientes, pDb2FechaActual)
    Dim rsCtoDto,rsColumnas,rsFilas
    set rsCtoDto = obtenerSQLAjusteCredDebi(pFechaActual, pPto, pLstClientes)
    if(not rsCtoDto.eof)then
    	Set rsColumnas	= ArmarColumnaAjustes(pFechaActual,pPto," cd.cdproducto,p.dsproducto ",pLstClientes)
    	Set rsFilas		= ArmarColumnaAjustes(pFechaActual,pPto," cd.cdcliente,c.dscliente ",pLstClientes	)
    	Call armarSeccionDescargas(pPto, rsCtoDto, rsColumnas, rsFilas, "", "Ajustes Creditos Debitos del " & GF_FN2DTE(pDb2FechaActual), "Debito", "Credito")
    end if
End function
'--------------------------------------------------------------------------------------------------------
Function armarAjusteStock(pFechaActual,pPto, pLstClientes, pDb2FechaActual)
    Dim rsAjuStock,rsColumnas,rsFilas
    set rsAjuStock = obtenerSQLAjusteStock(pFechaActual, pPto,pLstClientes)
    if(not rsAjuStock.eof)then

    call writeXLS("<table><tr><td colspan='2' class='xls_align_left'>Ajustes de Stock del "& GF_FN2DTE(pDb2FechaActual) &"</td></tr></table>")
    call writeXLS("<table style='width:80%; font-size:9;' class='xls_border_left' >")
    call writeXLS(" <tr style='background-color:#BABABA;' whidt='200%'>")
    call writeXLS("     <td whidt='60px' colspan='1' class='xls_align_center'>CLIENTE</td>")
    z =0
	filtroProducto = " a.cdproducto,dsproducto "
	filtroCliente =  " a.cdcliente,dscliente "
	Set rsColumnas = ArmarColumnaAjusteStock(pFechaActual,pPto,filtroProducto,pLstClientes)
	if (not rsColumnas.eof) then
	    Redim vectCol(rsColumnas.recordcount)
		Redim  vecTotalesProducto(rsColumnas.recordcount)
	end if
	while(not rsColumnas.eof)
	    call writeXLS("<td  colspan='1' class='xls_align_center'>"& rsColumnas("dsproducto") & "</td>")
        vectCol(z)= rsColumnas("cdproducto")
		z = z + 1
		rsColumnas.movenext
	wend
    call writeXLS("</tr>")
    Set rsFilas = ArmarColumnaAjusteStock(pFechaActual,pPto,filtroCliente,pLstClientes)		
    while(not rsFilas.eof)
        call writeXLS("<tr><td  align='left'>"& rsFilas("cdcliente") & " - " & rsFilas("dscliente") &"</td>")
        seguir= true	
		i = 0				 
		while((not rsAjuStock.eof) and seguir)					
		    if(cdbl(rsFilas("cdcliente")) = cdbl(rsAjuStock("cdcliente")))then 				
			    if (cdbl(vectCol(i)) = cdbl(rsAjuStock("cdproducto")))then
	                call writeXLS("<td align='right' >"& cdbl(rsAjuStock("Kilos")) &"</td>")
					vecTotalesProducto(i) = cdbl(vecTotalesProducto(i)) + cdbl(rsAjuStock("Kilos"))
					rsAjuStock.MoveNext()
					i = i + 1
				else
				    call writeXLS("<td align='right' ></td>")
					i = i+1								
					if (i > UBound(vectCol)) then	seguir = false
				end if
			else					
			    seguir = false
			end if					
		wend
		rsFilas.MoveNext()
        call writeXLS("</tr>")
    wend
    call writeXLS("<tr style='background-color:#EDEDED;' class='xls_align_center'>")
    call writeXLS("<td colspan='1' class='xls_align_center'>TOTAL</td>")
    for i = LBound(vecTotalesProducto) to UBound(vecTotalesProducto)-1
        call writeXLS("<td class='xls_align_right'>"& vecTotalesProducto(i) & "</td>")
    next
    call writeXLS("</tr></table> <br><br><br>")

end if
End function
'--------------------------------------------------------------------------------------------------------+
Function armarSaldoFinal(pFechaActual,pPto, pLstClientes, pDb2FechaActual)
    Dim rsSalIni,rsColumnas,rsFilas
    Set rsSalIni = ObtenerSQLSaldoInicial(pFechaActual, pPto, pLstClientes)
    if(not rsSalIni.eof)then
        call writeXLS("<table><tr><td colspan='2' class='xls_align_left'>Posicion del "& GF_FN2DTE(pDb2FechaActual) &" 24:00 Horas</td></tr></table>")
        Call armarSaldos(pPto, rsSalIni, pFechaActual, pLstClientes)
    end if
End Function
%>
