<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<%
dim idAlmacenes, fechaCierre, strSQL, rs, con, idDivision, fechaCierreAnterior
dim myCtaDebe, myCtaHaber, myCCostos, myCtaInterOrigen, myCtaInterDestino, myIdDivisionExpo, myCantidad
flagHeader = false
idDivision = GF_Parametros7("idDivision", 0, 6)
idAlmacenes = GF_Parametros7("idAlmacen", "", 6)
fechaCierre = GF_Parametros7("fechaCierre", "", 6)
fechaCierreAnterior = GF_Parametros7("fechaCierreAnt", "", 6)

''''REVERSION DE PROVISIONES''''
Response.Write "<hr>REVERSION DE PROVISIONES<hr>"
strSQL ="INSERT INTO TBLARTCTACTE SELECT " & fechaCierre & ", IDDIVISION, IDVALE, IDARTICULO, CANTIDAD, VLUPESOS, VLUDOLARES, '" & REVERSION_PROVISION & "', CUENTAGASTOS, CUENTAINVENTARIO, CCOSTOS, " & session("mmtodato") & " FROM TBLARTCTACTE WHERE FECHACIERRE=" & fechaCierreAnterior & " AND IDDIVISION=" & idDivision & " AND TIPOVALUACION='" & PROVISION & "'"
Response.Write "<HR>" & strSQL & "<HR>"
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
''''FIN REVERSION DE PROVISIONES''''

''''GASTOS''''
Response.Write "<hr>GASTOS<hr>"
'ARMADO DE LA SQL
'	Se une la cabecera con el detalle de los vales.
'	Se une TBLVALESCONTABLE para que solo se traigan aquellos vales-articulos valorizados en el mes
'	Se une TBLARTICULOS y TBLARTCATEGORIAS para obtener los datos de las cuentas de inverntario
'	Se unen con LEFT TBLBUDGETOBRAS y TBLDATOSOBRAS para qeu si tiene traiga el budgetr asociado al vale, cuentas de gastos
'	Se une con LEFT TBLSECTORES para qeu si tiene traiga el sector asociado al vale, cuentas de gastos
'	Se ordenan de acuerdo a la fecha del vale
strSQL = "SELECT VC.IDVALE, VC.CDVALE, VC.IDSECTOR, VC.IDBUDGETAREA, VC.IDBUDGETDETALLE, DOB.ESINVERSION, CO.IDARTICULO, CO.CANTIDAD AS CANTIDAD_VALUADA, CO.VLUPESOS AS PRECIO_CONTABLE, CO.VLUDOLARES AS PRECIO_CONTABLE_D, " & _
		 "	ART.CDCUENTA AS ART_CTAINVENTARIO, ART.CDCUENTAGASTOS AS ART_CTAGASTOS, ART.CCOSTOS AS ART_CCOSTOS, " & _
 		 "	CAT.CDCUENTA AS CAT_CTAINVENTARIO, CAT.CDCUENTAGASTOS AS CAT_CTAGASTOS, CAT.CCOSTOS AS CAT_CCOSTOS, CAT.ESMANTENIMIENTO, " & _
 		 "	BO.CDCUENTA AS BUD_CTAGASTOS, BO.CCOSTOS AS BUD_CCOSTOS, " & _ 		 
 		 "	SEC.CDCUENTA AS SEC_CTAGASTOS, SEC.CCOSTOS AS SEC_CCOSTOS " & _ 		  		 
		 "	FROM TBLVALESCABECERA VC " & _ 
		 "	    INNER JOIN TBLVALESDETALLE VD ON VC.IDVALE = VD.IDVALE " & _ 
		 "	    INNER JOIN TBLVALESCONTABLE CO ON VD.IDVALE=CO.IDVALE AND VD.IDARTICULO=CO.IDARTICULO " & _ 
		 "	    INNER JOIN TBLARTICULOS ART ON ART.IDARTICULO=VD.IDARTICULO " & _
		 "	    INNER JOIN TBLARTCATEGORIAS CAT ON CAT.IDCATEGORIA=ART.IDCATEGORIA " & _
		 "	    LEFT JOIN TBLBUDGETOBRAS BO ON BO.IDOBRA=VC.IDOBRA AND BO.IDAREA=VC.IDBUDGETAREA AND BO.IDDETALLE=VC.IDBUDGETDETALLE " & _  
		 "	    LEFT JOIN TBLDATOSOBRAS DOB ON DOB.IDOBRA=VC.IDOBRA " & _ 
		 "	    LEFT JOIN TBLSECTORES SEC ON VC.IDSECTOR=SEC.IDSECTOR " & _ 
		 "	    WHERE VC.ESTADO=" & ESTADO_ACTIVO & " AND VC.IDALMACEN IN (" & idAlmacenes  & ") AND CO.FECHACIERRE=" & fechaCierre & _
		 "		AND VC.FECHA>'20121231' " & _ 
		 "		ORDER BY VC.FECHA "
Response.Write "<HR>" & strSQL & "<HR>"
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
while not rs.eof
	myCtaDebe = "FALTA_GTO"
	'Generales para cualquiera de los vales
	'Tomar cuenta de inventario de articuloc o de ategoria
	myCtaHaber = rs("ART_CTAINVENTARIO")
	if ((isnull(myCtaHaber)) or (trim(myCtaHaber) = "")) then 
		myCtaHaber = trim(rs("CAT_CTAINVENTARIO"))
		if ((isnull(myCtaHaber)) or (trim(myCtaHaber) = "")) then 
			myCtaHaber = "FALTA_INV"
		end if
	end if	

	select case rs("CDVALE")
		case CODIGO_VS_SALIDA, CODIGO_VS_AJUSTE_VALE
			'Es inversion?
			'Si es una inversion y posee area y detalle de partida es una obra en curso
			if trim(rs("ESINVERSION")) = "S" and (rs("IDBUDGETAREA") <> 0 AND rs("IDBUDGETDETALLE") <> 0) then
				myCtaDebe = CUENTA_OBRAENCURSO
				myCCostos = ""
			else
				'Si posee area y detalle de partida y no es una inversion, la cuenta de gastos sera la asociada a la partida
				if rs("IDBUDGETAREA") <> 0 AND rs("IDBUDGETDETALLE") <> 0 then
					myCtaDebe = rs("BUD_CTAGASTOS")
					myCCostos = rs("BUD_CCOSTOS")
				'No tiene area y detalle pero si articulos de mantenimiento por lo tanto la cta gastos es la del sector	
				elseif rs("ESMANTENIMIENTO") = "S" then
 					myCtaDebe = rs("SEC_CTAGASTOS")
					myCCostos = rs("SEC_CCOSTOS")
				else
					'No tiene area ni detalle y tampoco art de mantenimiento, la cta gastos esta en elpropio art o categoria
					myCtaDebe = rs("ART_CTAGASTOS")
					myCCostos = rs("ART_CCOSTOS")
					if ((isnull(myCtaDebe)) or (trim(myCtaDebe) = "")) then 
						myCtaDebe = trim(rs("CAT_CTAGASTOS"))
						myCCostos = rs("CAT_CCOSTOS")
					end if	
				end if	
			end if
			if ((isnull(myCtaDebe)) or (trim(myCtaDebe) = "")) then myCtaDebe = "FALTA_GTO"
			call insertCtaCte(idDivision, rs("IDVALE"), rs("IDARTICULO"), rs("CANTIDAD_VALUADA"), rs("PRECIO_CONTABLE"), rs("PRECIO_CONTABLE_D"), GASTO, myCtaDebe, myCtaHaber, myCCostos)
		case CODIGO_VS_AJUSTE_TRANSFERENCIA
			myCtaDebe = CUENTA_AJUSTE_STOCK
			myCCostos = ""
			call insertCtaCte(idDivision, rs("IDVALE"), rs("IDARTICULO"), cdbl(rs("CANTIDAD_VALUADA")), rs("PRECIO_CONTABLE"), rs("PRECIO_CONTABLE_D"), GASTO, myCtaDebe, myCtaHaber, myCCostos)
		case CODIGO_VS_AJUSTE_STOCK
			myCtaDebe = CUENTA_AJUSTE_STOCK
			myCCostos = ""
			call insertCtaCte(idDivision, rs("IDVALE"), rs("IDARTICULO"), cdbl(rs("CANTIDAD_VALUADA")) * -1, rs("PRECIO_CONTABLE"), rs("PRECIO_CONTABLE_D"), GASTO, myCtaDebe, myCtaHaber, myCCostos)
		case CODIGO_VS_RECLASIFICACION_STOCK
			myCtaDebe = "VRS"
			call insertCtaCte(idDivision, rs("IDVALE"), rs("IDARTICULO"), cdbl(rs("CANTIDAD_VALUADA"))*-1, rs("PRECIO_CONTABLE"), rs("PRECIO_CONTABLE_D"), GASTO, myCtaDebe, myCtaHaber, myCCostos)
		case CODIGO_VS_TRANSFERENCIA	
			'Se registra el gasto del origen de la transferencia.
			myCtaInterOrigen = getCuentaInterdivisional(idDivision)	
			call insertCtaCte(idDivision, rs("IDVALE"), rs("IDARTICULO"), rs("CANTIDAD_VALUADA"), rs("PRECIO_CONTABLE"), rs("PRECIO_CONTABLE_D"), GASTO, myCtaInterOrigen, myCtaHaber, myCCostos)			
		case CODIGO_VS_RECEPCION
			'Se debe hacer el pase entre cuentas interdivisionales
			myCtaInterDestino = getCuentaInterdivisional(idDivision)
			myCtaInterOrigen = getCtaInterOrigen(rs("IDVALE"))
			myIdDivisionExpo = getDivisionID(CODIGO_EXPORTACION)
			call insertCtaCte(myIdDivisionExpo, rs("IDVALE"), rs("IDARTICULO"), rs("CANTIDAD_VALUADA"), rs("PRECIO_CONTABLE"), rs("PRECIO_CONTABLE_D"), GASTO, myCtaInterDestino, myCtaInterOrigen, myCCostos)						
			'Se incrementa el inventario del receptor de la mercaderia
			call insertCtaCte(idDivision, rs("IDVALE"), rs("IDARTICULO"), rs("CANTIDAD_VALUADA"), rs("PRECIO_CONTABLE"), rs("PRECIO_CONTABLE_D"), GASTO, myCtaHaber, myCtaInterDestino, myCCostos)									
		case else
		
			'Nada
		end select	

	rs.movenext
wend	
''''FIN GASTOS''''

''''''''''''''''''''''''''''''''''''''''''''''''''''PROVISIONES''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Response.Write "<hr>PROVISIONES<hr>"
'ARMADO DE LA SQL
'	Se une la cabecera con el detalle de los vales.
'	Se hace el left TBLVALESCONTABLE y luego en la clausula where se quitan los que ya estan para que solo se traigan aquellos vales-articulos que no estan valorizados en el mes
'	Se une TBLARTICULOS y TBLARTCATEGORIAS para obtener los datos de las cuentas de inverntario
'	La cuenta de gastos sera PROVISIONES!
'	Se ordenan de acuerdo a la fecha del vale
strSQL = "SELECT VC.IDVALE, VC.CDVALE, VD.IDARTICULO, VC.IDBUDGETAREA, VC.IDBUDGETDETALLE, VD.EXISTENCIA AS EXISTENCIA, T1.CANTIDAD_VALUADA, VD.VLUPESOS AS PRECIO, VD.VLUDOLARES AS PRECIO_D, " & _
         "          DOB.ESINVERSION, BO.CDCUENTA AS BUD_CTAGASTOS, BO.CCOSTOS AS BUD_CCOSTOS,  " & _
         "       	SEC.CDCUENTA AS SEC_CTAGASTOS, SEC.CCOSTOS AS SEC_CCOSTOS, " & _
		 "			ART.CDCUENTAGASTOS AS ART_CTAGASTOS, ART.CCOSTOS AS ART_CCOSTOS, " & _
 		 "			CAT.CDCUENTAGASTOS AS CAT_CTAGASTOS, CAT.CCOSTOS AS CAT_CCOSTOS, CAT.ESMANTENIMIENTO " & _
		 "	FROM TBLVALESCABECERA VC  " & _
		 "	    INNER JOIN TBLVALESDETALLE VD ON VC.IDVALE = VD.IDVALE  " & _
		 "      LEFT JOIN (SELECT IDVALE, IDARTICULO, SUM(CANTIDAD) AS CANTIDAD_VALUADA FROM TBLVALESCONTABLE GROUP BY IDVALE, IDARTICULO) T1  " & _
		 "			ON VD.IDVALE= T1.IDVALE AND VD.IDARTICULO=T1.IDARTICULO  " & _
		 "	    INNER JOIN TBLARTICULOS ART ON ART.IDARTICULO=VD.IDARTICULO " & _
		 "	    INNER JOIN TBLARTCATEGORIAS CAT ON CAT.IDCATEGORIA=ART.IDCATEGORIA " & _
		 "	    LEFT JOIN TBLBUDGETOBRAS BO ON BO.IDOBRA=VC.IDOBRA AND BO.IDAREA=VC.IDBUDGETAREA AND BO.IDDETALLE=VC.IDBUDGETDETALLE " & _
		 "	    LEFT JOIN TBLDATOSOBRAS DOB ON DOB.IDOBRA=VC.IDOBRA  " & _
		 "	    LEFT JOIN TBLSECTORES SEC ON VC.IDSECTOR=SEC.IDSECTOR  " & _
		 "	    WHERE (T1.IDARTICULO IS NULL OR T1.CANTIDAD_VALUADA < VD.EXISTENCIA)  " & _
		 "			AND VC.IDALMACEN IN (" & idAlmacenes  & ") AND VC.FECHA BETWEEN '" & FECHA_INICIO_CONTABLE & "' AND '" & fechaCierre & "'" & _
		 "			AND VC.CDVALE IN ('" & CODIGO_VS_SALIDA & "','" & CODIGO_VS_AJUSTE_VALE & "','" & CODIGO_VS_AJUSTE_TRANSFERENCIA & "','" & CODIGO_VS_AJUSTE_STOCK & "')" & _
		 "			AND EXISTENCIA <> 0 AND VC.ESTADO=" & ESTADO_ACTIVO & _
		 "		ORDER BY VC.FECHA "
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)		 
Response.Write "<HR>" & strSQL & "<HR>"
while not rs.eof
	myCtaHaber = CUENTA_PROVISIONES
	
	select case rs("CDVALE")
		case CODIGO_VS_SALIDA, CODIGO_VS_AJUSTE_VALE
			'Es inversion?
			'Si es una inversion y posee area y detalle de partida es una obra en curso
			if trim(rs("ESINVERSION")) = "S" and (rs("IDBUDGETAREA") <> 0 AND rs("IDBUDGETDETALLE") <> 0) then
				myCtaDebe = CUENTA_OBRAENCURSO
				myCCostos = ""
			else
				'Si posee area y detalle de partida y no es una inversion, la cuenta de gastos sera la asociada a la partida
				if rs("IDBUDGETAREA") <> 0 AND rs("IDBUDGETDETALLE") <> 0 then
					myCtaDebe = rs("BUD_CTAGASTOS")
					myCCostos = rs("BUD_CCOSTOS")
				'No tiene area y detalle pero si articulos de mantenimiento por lo tanto la cta gastos es la del sector	
				elseif rs("ESMANTENIMIENTO") = "S" then
 					myCtaDebe = rs("SEC_CTAGASTOS")
					myCCostos = rs("SEC_CCOSTOS")
				else
					'No tiene area ni detalle y tampoco art de mantenimiento, la cta gastos esta en elpropio art o categoria
					myCtaDebe = rs("ART_CTAGASTOS")
					myCCostos = rs("ART_CCOSTOS")					
					if ((isnull(myCtaDebe)) or (trim(myCtaDebe) = "")) then 
						myCtaDebe = trim(rs("CAT_CTAGASTOS"))
						myCCostos = rs("CAT_CCOSTOS")						
					end if	
				end if	
			end if
			if ((isnull(myCtaDebe)) or (trim(myCtaDebe) = "")) then myCtaDebe = "FALTA_GTO"
			'call insertCtaCte(idDivision, rs("IDVALE"), rs("IDARTICULO"), rs("CANTIDAD_VALUADA"), rs("PRECIO_CONTABLE"), GASTO, myCtaDebe, myCtaHaber, myCCostos)
		case CODIGO_VS_AJUSTE_TRANSFERENCIA
		    myCtaDebe = CUENTA_PROVISION_AJT
			myCCostos = CCOSTO_PROVISION_AJT
		case CODIGO_VS_AJUSTE_STOCK
			myCtaDebe = CUENTA_AJUSTE_STOCK
			myCCostos = ""
		case else
			'Nada
		end select		
	
	myCantidad = cdbl(rs("EXISTENCIA"))
	Response.Write "1-CANTIDAD(" & myCantidad & ")"
	if not isnull(rs("CANTIDAD_VALUADA")) then myCantidad = myCantidad - cdbl(rs("CANTIDAD_VALUADA"))
	Response.Write "2-CANTIDAD(" & myCantidad & ")"
	Call insertCtaCte(idDivision, rs("IDVALE"), rs("IDARTICULO"), myCantidad, cdbl(rs("PRECIO"))*100, cdbl(rs("PRECIO_D"))*100, PROVISION, myCtaDebe, myCtaHaber, myCCostos)
	rs.movenext
wend	

'Transferecnias que puede ser parte de la provision
'VMT que fueron recibidos pero no pudo valorizarse contablemente el vale
'Primero se obtienen todos los VMR que fueron valorizados en la primera etapa (valorizacionContableVMRAjax)
'Se debe obtener la diferencia entra la cantidad que recibio el vale y la cantidad que se pudo valorizar
'Se hara la provision, tanto para el VMT como para el VMR, por la diferencia
Response.Write "<hr>VMT COMO PROVISION<hr>"
strSQL = "SELECT VC.IDVALE, VD.IDARTICULO, VD.EXISTENCIA, VD.VLUPESOS, VD.VLUDOLARES, CO.CANTIDAD, " & _
		 "	ART.CDCUENTA AS ART_CTAINVENTARIO, ART.CCOSTOS AS ART_CCOSTOS, " & _
 		 "	CAT.CDCUENTA AS CAT_CTAINVENTARIO, CAT.CCOSTOS AS CAT_CCOSTOS " & _
		 " FROM	" & _
		 "	(SELECT IDVALE, IDARTICULO, SUM(CANTIDAD) AS CANTIDAD FROM TBLVALESCONTABLE WHERE FECHACIERRE BETWEEN " & FECHA_INICIO_CONTABLE & " AND " & fechaCierre & " GROUP BY IDVALE, IDARTICULO) AS CO " & _ 
		 "	INNER JOIN TBLVALESCABECERA VC ON CO.IDVALE=VC.IDVALE AND VC.CDVALE='" & CODIGO_VS_RECEPCION & "'" & _
		 "	INNER JOIN TBLVALESDETALLE VD ON CO.IDVALE=VD.IDVALE AND CO.IDARTICULO=VD.IDARTICULO AND CO.CANTIDAD<>VD.EXISTENCIA " & _
		 "	    INNER JOIN TBLARTICULOS ART ON ART.IDARTICULO=VD.IDARTICULO " & _
		 "	    INNER JOIN TBLARTCATEGORIAS CAT ON CAT.IDCATEGORIA=ART.IDCATEGORIA " & _		 
		 "	WHERE VC.ESTADO=" & ESTADO_ACTIVO & " AND VC.IDALMACEN IN (" & idAlmacenes  & ")" & _
		 "		AND VC.FECHA>'20121231' " & _ 		 
		 "  ORDER BY VC.IDVALE ASC "
Response.Write "<HR>" & strSQL & "<HR>"	
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
while not rs.eof 
		'Tomar cuenta de inventario de articuloc o de ategoria
		myCtaHaber = rs("ART_CTAINVENTARIO")
		if ((isnull(myCtaHaber)) or (trim(myCtaHaber) = "")) then 
			myCtaHaber = trim(rs("CAT_CTAINVENTARIO"))
			if ((isnull(myCtaHaber)) or (trim(myCtaHaber) = "")) then 
				myCtaHaber = "FALTA_INV"
			end if
		end if	
		myCtaDebe = CUENTA_PROVISIONES
		myCantidad = cdbl(rs("EXISTENCIA"))
		if not IsNull(rs("CANTIDAD")) then myCantidad = myCantidad - cdbl(rs("CANTIDAD"))
		call insertCtaCte(idDivision, rs("IDVALE"), rs("IDARTICULO"), myCantidad, cdbl(rs("VLUPESOS"))*100, cdbl(rs("VLUDOLARES"))*100, PROVISION, myCtaDebe, myCtaHaber, myCCostos)
		call obtenerVMT(rs("IDVALE"), myIdVMT, myAlmacenVMT)
		if myIdVMT <> 0 then
			call insertCtaCte(getDivisionAlmacen(myAlmacenVMT), myIdVMT, rs("IDARTICULO"), myCantidad, cdbl(rs("VLUPESOS"))*100, cdbl(rs("VLUDOLARES"))*100, PROVISION, myCtaDebe, myCtaHaber, myCCostos)
		end if	
	rs.movenext
wend	
Response.Write "<HR>" & strSQL & "<HR>"


''''''''''''''''''''''''''''''''''''''''''''''''''''FIN PROVISIONES''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''REVERSION DE MERCADERIA EN TRANSITO''''
Response.Write "<hr>REVERSION DE MERCADERIA EN TRANSITO<hr>"
strSQL ="INSERT INTO TBLARTCTACTE SELECT " & fechaCierre & ", IDDIVISION, IDVALE, IDARTICULO, CANTIDAD, VLUPESOS, VLUDOLARES, '" & REVERSION_MERCADERIA_TRANSITO & "', CUENTAINVENTARIO, CUENTAGASTOS, CCOSTOS, " & session("mmtodato") & " FROM TBLARTCTACTE WHERE FECHACIERRE=" & fechaCierreAnterior & " AND IDDIVISION=" & idDivision & " AND TIPOVALUACION='" & MERCADERIA_TRANSITO & "'"
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
Response.Write "<HR>" & strSQL & "<HR>"
''''FIN REVERSION DE MERCADERIA EN TRANSITO''''

''''MERCADERIA EN TRANSITO''''
'Se obtienen todos los PM de transferencias que no estan saldados
'
Response.Write "<hr>MERCADERIA EN TRANSITO<hr>"
strSQL = "SELECT PARTIDAPENDIENTE, IDARTICULO, SUM(TOTAL_VALES) AS SALDO_VALES FROM " & _
		 "	( " & _
		"	SELECT VC.PARTIDAPENDIENTE, VD.IDARTICULO, CASE VC.CDVALE WHEN '" & CODIGO_VS_AJUSTE_TRANSFERENCIA & "' THEN (-VD.EXISTENCIA) WHEN '" & CODIGO_VS_RECEPCION & "' THEN (-VD.EXISTENCIA) WHEN '" & CODIGO_VS_TRANSFERENCIA & "' THEN (VD.EXISTENCIA) END AS TOTAL_VALES  " & _
		"	    FROM TBLVALESCABECERA VC  " & _
		"	        INNER JOIN TBLVALESDETALLE VD ON VC.IDVALE=VD.IDVALE  " & _
		"	        INNER JOIN TBLPMCABECERA PM ON PM.IDPEDIDO=VC.PARTIDAPENDIENTE  " & _
		"	        INNER JOIN TBLALMACENES ALM_ORI ON PM.IDALMACEN=ALM_ORI.IDALMACEN  " & _
		"	        INNER JOIN TBLALMACENES ALM_DES ON PM.IDALMACENDEST=ALM_DES.IDALMACEN " & _
		"	        WHERE ALM_ORI.IDDIVISION<>ALM_DES.IDDIVISION  " & _
		"	            AND VC.CDVALE IN ('" & CODIGO_VS_RECEPCION & "','" & CODIGO_VS_TRANSFERENCIA & "','" & CODIGO_VS_AJUSTE_TRANSFERENCIA & "') " & _
		"				AND VC.ESTADO=" & ESTADO_ACTIVO & " AND VD.EXISTENCIA<>0 AND VC.FECHA BETWEEN " & FECHA_INICIO_CONTABLE & " AND " & fechaCierre & _
		"				AND ALM_ORI.IDDIVISION = " & idDivision & _
		"	) T1 " & _
		"	GROUP BY PARTIDAPENDIENTE, IDARTICULO " & _
		"	HAVING SUM(TOTAL_VALES) > 0 "
Response.Write "<HR>" & strSQL & "<HR>"	
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
while not rs.eof 
		call insertCtaCteMercaderiaEnTransito(rs("PARTIDAPENDIENTE"), rs("IDARTICULO"), rs("SALDO_VALES"))
	rs.movenext
wend	
Response.Write "<HR>" & strSQL & "<HR>"
''''FIN MERCADERIA EN TRANSITO''''
'--------------------------------------------------------------------------------------------------------------------------
sub insertCtaCteMercaderiaEnTransito(pPM, pIdArticulo, pSaldo)
'Esta funcion recibe el PM, el artiuclo y la cantidad a saldar.
'Se debe asignar el o los VMT que tengan ese articulo y la cantidad que se debe
dim strSQL, rs, mySaldo
strSQL = "SELECT * FROM " & _
		 "	TBLVALESCABECERA VC " & _ 
		 "		INNER JOIN TBLVALESDETALLE VD ON VC.IDVALE=VD.IDVALE AND VC.CDVALE='" & CODIGO_VS_TRANSFERENCIA & "' AND VC.ESTADO=" & ESTADO_ACTIVO & " AND VC.PARTIDAPENDIENTE=" & pPM & " AND EXISTENCIA<>0 AND IDARTICULO=" & pIdArticulo
Response.Write "<HR>" & strSQL & "<HR>"
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
while not rs.eof and cdbl(pSaldo) > 0
		if CDbl(rs("EXISTENCIA")) >= cdbl(pSaldo) then
			myUnidades = pSaldo
			pSaldo = 0
		else
			myUnidades = CDbl(rs("EXISTENCIA"))
			pSaldo = pSaldo - myUnidades
		end if	
		call insertCtaCte(idDivision, rs("IDVALE"), rs("IDARTICULO"), myUnidades, CDbl(rs("VLUPESOS"))*100, CDbl(rs("VLUDOLARES"))*100, MERCADERIA_TRANSITO, "","","")
	rs.movenext
wend		

end sub
'-------------------------------------------------------------------------------------------------------------------------
sub insertCtaCte(pIdDivision, pIdVale, pIdArticulo, pCantidad, pVluPesos, pVluDolares, pTipoValuacion, pCtaDebe, pCtaHaber, pCCostos)
dim myCtaDebe, myCtaHaber, myCentroCostos, rs, strSQL
if pTipoValuacion = PROVISION and (pCtaDebe=CUENTA_OBRAENCURSO or pCtaHaber=CUENTA_OBRAENCURSO) then
	'No se provisionan las inversiones
else	
	strSQL = "INSERT INTO TBLARTCTACTE VALUES (" & fechaCierre & "," & pIdDivision & "," & pIdVale & "," & pIdArticulo & "," & pCantidad & "," & pVluPesos & "," & pVluDolares & ",'" & pTipoValuacion & "','" & trim(pCtaHaber) & "','" & trim(pCtaDebe) & "','" & trim(pCCostos) & "','" & session("mmtodato") & "')"
	'Response.Write "<HR>" & strSQL & "<HR>"
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
end if
end sub
'-------------------------------------------------------------------------------------------------------------------------
function getCtaInterOrigen(pIdVale)
'Esta funcion devuelve la cuenta interdivisional del origen de una transferencia. o sea la cuenta de la division que genero el VMT
'Se recibe por parametro el VMR, con lo cual hay que buscar el nro de VMT y a partir de la division a la que corresponda
'obtener la cuenta
'Un VMR va a tener siempre un unico VMT, no asi al reves, un VMT puede tener varios VMR
dim rs, strSQL, rtrn
rtrn = "ERR INT-O"
strSQL = "SELECT IDDIVISION FROM TBLVALESCABECERA VC " & _
		 " INNER JOIN TBLVALESCABECERA VC1 ON VC.PARTIDAPENDIENTE=VC1.PARTIDAPENDIENTE AND VC1.CDVALE='" & CODIGO_VS_TRANSFERENCIA & "'" & _  
		 " INNER JOIN TBLALMACENES AL ON AL.IDALMACEN=VC1.IDALMACEN " & _
		 " WHERE VC.IDVALE=" & pIdVale 
Response.Write "<HR>" & strSQL & "<HR>"
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
if not rs.eof then rtrn = getCuentaInterdivisional(rs("IDDIVISION"))
getCtaInterOrigen = rtrn	
end function
'-------------------------------------------------------------------------------------------------------------------------
Response.Write "Hecho..."
'----------------------------------------------------------------------------------------
%>