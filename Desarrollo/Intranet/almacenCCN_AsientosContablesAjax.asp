  <!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<% 
dim myNroAsiento, myDescripcion, myDescripcionSAF
dim idCierre, rs, oConn, strSQL, contadorDetalleSAF, contadorDetalle, impDolares, impPesos, impDbCr
dim idAlmacen, idDivision, fechaAsiento, fechaCierre, myAIA2, myNombreMiembro, myDivisionAux
idAlmacen = GF_Parametros7("idAlmacen", "", 6)  
idCierre = GF_Parametros7("idCierre", 0, 6)
fechaAsiento = GF_Parametros7("fechaAsiento", "", 6) 
fechaCierre = GF_Parametros7("fechaCierre", "", 6)
idDivision = GF_Parametros7("idDivision", "", 6)  
ciaDivision = getCIADivision(idDivision)
	strSQL ="Select IDCIERRE, CDCUENTA, CCOSTOS, DBCR, TIPOCAMBIO, round(SUM(IMPORTEPESOS)/10000, 2) IMPORTEPESOS, round(SUM(IMPORTEDOLARES)/10000, 2) IMPORTEDOLARES from TBLCIERRESASIENTOS2 WHERE IDCIERRE=" & idCierre & " group by IDCIERRE, CDCUENTA, CCOSTOS, DBCR, TIPOCAMBIO"
	'Response.Write strsqL
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if not rs.eof then
		myNroAsiento = getProximoNroCabecera(ciaDivision,CODIGO_CPTE_CIERRES,left(fechaCierre,4),mid(fechaCierre,5,2))
		'GUARDAR CABECERA PARA CONTABILIDAD
		call loadDescripcion(rs("IDCIERRE"), myDescripcion, myDescripcionSAF)
		call addCabeceraContable(ciaDivision, CODIGO_CPTE_CIERRES, myNroAsiento, fechaAsiento, myDescripcion)
		while not rs.eof 
				if trim(rs("CDCUENTA")) = CUENTA_OBRAENCURSO then
						strSQL ="	SELECT TG.IDOBRA, DOB.CDOBRA, BOB.CDCUENTA,  TG.IDBUDGETAREA, TG.IDBUDGETDETALLE, TG.TOTALPESOS " & _
								"	FROM " & _
								"	( " & _
								"	SELECT T1.IDOBRA, T1.IDBUDGETAREA, T1.IDBUDGETDETALLE, SUM(T1.TOTALPESOS) AS TOTALPESOS  FROM " & _
								"		( " & _
								"		SELECT VD.IDARTICULO, VC.IDOBRA, VC.IDBUDGETAREA, VC.IDBUDGETDETALLE, SUM(CO.CANTIDAD*CO.VLUPESOS) AS TOTALPESOS " & _
								"			FROM TBLVALESCABECERA VC  " & _
								"		 	    INNER JOIN TBLVALESDETALLE VD ON VC.IDVALE = VD.IDVALE " & _
					 			"				INNER JOIN TBLVALESCONTABLE CO ON VD.IDVALE=CO.IDVALE AND VD.IDARTICULO=CO.IDARTICULO " & _
								"				INNER JOIN TBLCIERRESARTICULOS2 CIE ON VD.IDARTICULO=CIE.IDARTICULO AND CIE.IDALMACEN IN (" & idAlmacen & ") AND CIE.FECHACIERRE=" & fechaCierre & _
								"				WHERE VC.IDALMACEN IN (" & idAlmacen & ") AND VC.FECHA LIKE '" & left(fechaCierre,6) & "%' AND VC.ESTADO=" & ESTADO_ACTIVO & _
								"					AND VC.CDVALE IN ('" & CODIGO_VS_SALIDA & "','" & CODIGO_VS_AJUSTE_VALE & "') AND VC.IDBUDGETAREA <> 0 " & _
								"					AND VC.IDBUDGETDETALLE <> 0 " & _
								"				GROUP BY VD.IDARTICULO, VC.IDOBRA, VC.IDBUDGETAREA, VC.IDBUDGETDETALLE " & _
								"		)T1    " & _
								"	GROUP BY T1.IDOBRA, T1.IDBUDGETAREA, T1.IDBUDGETDETALLE " & _
								"	)TG   " & _
								"	INNER JOIN TBLDATOSOBRAS DOB ON TG.IDOBRA=DOB.IDOBRA AND DOB.ESINVERSION='S' " & _
								"	INNER JOIN TBLBUDGETOBRAS BOB ON TG.IDOBRA=BOB.IDOBRA AND TG.IDBUDGETAREA=BOB.IDAREA AND TG.IDBUDGETDETALLE=BOB.IDDETALLE"
						'Response.Write "<HR>" & strSQL
						'Response.End 
						Call executeQueryDB(DBSITE_SQL_INTRA, rsA, "OPEN", strSQL)
						while not rsA.eof 
							if trim(rsA("CDCUENTA")) <> "" then
								Response.Write "<br>1 - " & cdbl(rsA("TOTALPESOS"))
								Response.Write "<br>1 - " & cdbl(rs("TIPOCAMBIO"))								
								'Response.Write "<br>1 - " & round(cdbl(rsA("TOTALPESOS"))/cdbl(rs("TIPOCAMBIO")),0)
								if cdbl(rs("TIPOCAMBIO")) = 0 then 
									impDolares = 0
								else
									impDolares = round(cdbl(rsA("TOTALPESOS"))/cdbl(rs("TIPOCAMBIO")),0)
								end if	
								call addDetalleContableSAF(TRIM(rsA("CDOBRA")), TRIM(rsA("CDCUENTA")), myDescripcionSAF, fechaAsiento,rs("DBCR"), rsA("TOTALPESOS"), impDolares)
							end if
							rsA.movenext
						wend
				end if
				contadorDetalle = contadorDetalle + 1
				contadorDetalleSAF = contadorDetalleSAF + 1
				impPesos = CDbl(rs("IMPORTEPESOS"))
				impDolares = CDbl(rs("IMPORTEDOLARES"))
				impDbCr = CInt(rs("DBCR"))
				if (impPesos < 0) then
					impPesos = impPesos * -1
					impDolares = impDolares * -1
					if (impDbCr = 1) then 
						impDbCr = 2
					else 
						impDbCr = 1
					end if
				end if
				call addDetalleContable(ciaDivision, CODIGO_CPTE_CIERRES, fechaAsiento, myNroAsiento, rs("CDCUENTA"), rs("CCOSTOS"), myDescripcion, impPesos, impDolares, impDbCr)
			rs.MoveNext	
		wend
	    'Sanear Diferencia entre Debe-Haber
		Call eliminarDiferencias(ciaDivision, CODIGO_CPTE_CIERRES, fechaAsiento, myNroAsiento)    
		
	end if	
	call actualizarEstadoCierre(idCierre, TIPO_CIERRE_DEFINITIVO)
	'Response.Write "1 Cabecera, " & contadorDetalle & " Detalles, " & contadorDetalleSAF & " Detalles SAF - Hecho..."
'----------------------------------------------------------------------------------------
function loadDescripcion(pIdCierre, byRef pDescripcionCAB, byRef pDescripcionSAF)
dim strSQL, rs, oConn, rtrn, mes
pDescripcionCAB = ""
pDescripcionSAF = ""
strSQL ="SELECT * FROM TBLCIERRESCABECERA2 WHERE IDCIERRE=" & pIdCierre 
'Response.Write "<HR>" & strSQL
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if not rs.eof then 
		mes = rs("MES")
		if len(rs("MES")) = 1 then mes = "0" & rs("MES")
		pDescripcionCAB = "CONS. DIV. " & getCIADivision(rs("IDDIVISION")) & " " & mes & "-" & rs("ANIO") 
		pDescripcionSAF = "DIVISION " & getCIADivision(rs("IDDIVISION"))
	end if	
end function
'----------------------------------------------------------------------------------------
function getProximoNroCabecera(pCiaDivision, pCodigo, pAnio, pMes)
dim strSQL, rs, oConn
strSQL ="SELECT NUMERO FROM CGT014A WHERE CIA= " & pCiaDivision & " AND TIPOCPTE='" & CODIGO_CPTE_CIERRES & "' AND ANIO=" & pAnio & " AND MES=" & pMes 
'Response.Write "<HR>" & strSQL
Call executeQueryDB(DBSITE_SQL_MAGIC, rs, "OPEN", strSQL)
if not rs.eof then 
	rtrn = CLng(rs("NUMERO")) + 1
	strSQL ="UPDATE CGT014A Set Numero=" & rtrn & " WHERE CIA= " & pCiaDivision & " AND TIPOCPTE='" & CODIGO_CPTE_CIERRES & "' AND ANIO=" & pAnio & " AND MES=" & pMes 
	Call executeQueryDB(DBSITE_SQL_MAGIC, rs, "EXEC", strSQL)
else
	rtrn = CInt(pMes)*10000
	rtrn = CLng(rtrn) + 1
	'CARGAR NUEVA CABECERA
	strSQL ="INSERT INTO CGT014A VALUES(" & pCiaDivision & ",'" & pCodigo & "'," & pAnio & "," & pMes & "," & rtrn & ")" 
	Call executeQueryDB(DBSITE_SQL_MAGIC, rs, "EXEC", strSQL)
end if
getProximoNroCabecera = rtrn
'Response.Write " -- PROX(" & rtrn & ")"
end function
'----------------------------------------------------------------------------------------
sub addCabeceraContable(pCia, pTipoCbte, pNro, pFechaAsiento, pDescripcion)
dim strSQL, rs, oConn
strSQL ="INSERT INTO CGT040A (cia, tipocpte, numero, feccpte, reng1, usuario, fecdia, hora, erro, act, cia1, sucursal, nnn) " & _
		" VALUES('"& pCia &"','" & pTipoCbte & "', " & pNro & ",'" & GF_DT2DTCONTABLE(pFechaAsiento) & "', '" & pDescripcion & "','" & session("Usuario") & "'" & _
		" , '" & GF_DT2DTCONTABLE(date()) & "', '" & hour(Time()) & ":" & minute(Time()) & ":" & second(Time()) & "', '', '', '000', null, 0)"
'Response.Write "<HR>Add en Cabecera CTG040A<br>" & strSQL
Call executeQueryDB(DBSITE_SQL_MAGIC, rs, "EXEC", strSQL)
end sub
'----------------------------------------------------------------------------------------
sub addDetalleContable(pCia, pTipoCbte, pFechaAsiento, pNro, pCdCuenta, pCCostos, pDescripcion, pImportePesos, pImporteDolares, pDbCr)
dim myCCostos
dim strSQL, rs, oConn
if pDbCr = 1 then
	pDbCr = "D"
else
	pDbCr = "H"
end if
myCCostos = pCCostos
if trim(myCCostos)="" then myCCostos="0"
strSQL ="INSERT INTO CGT040B (cia,tipocpte,feccpte,numero,cuenta,aux,descrip,partida,centro,impomL,dh,unidades,impom1,impom2,leyerr,act,cia1,sucursal,contrato,cheque,codvap) " & _
		" VALUES('"& pCia & "','" & pTipoCbte & "', '" & GF_DT2DTCONTABLE(pFechaAsiento) & "', " & pNro & ", '" & pCdCuenta & "', " & myCCostos & ", '" & pDescripcion & "'" & _
		" , '', '', " & round(cdbl(pImportePesos), 2) & ",'" & pDbCr & "',0, " & round(cdbl(pImportePesos), 2) & ", " & round(cdbl(pImporteDolares), 2) & _
		" ,'', '', '', '', 0, 0, 0)"
'Response.Write "<HR>Add en Detalle CTG040B<br>" & strSQL
Call executeQueryDB(DBSITE_SQL_MAGIC, rs, "EXEC", strSQL)
end sub



'----------------------------------------------------------------------------------------
sub eliminarDiferencias(pCia, pTipoCbte, pFechaAsiento, pNroAsiento)
dim strSQL, rsDebe, rsHaber, oConn, rsC
Response.Write "<hr>Eliminando Diferencias"
'Calculo totales Debito y Credito
strSQL= "Select sum(IMPOML) ImportePesos, sum(IMPOM2) ImporteDolares from CGT040B where CIA='" & pCia & "' AND tipocpte='" & pTipoCbte & "' AND NUMERO=" & pNroAsiento & " and FECCPTE='" & GF_DT2DTCONTABLE(pFechaAsiento) & "' and DH='D'"
'response.Write "<br>" & strSQL
Call executeQueryDB(DBSITE_SQL_MAGIC, rsDebe, "OPEN", strSQL)
strSQL= "Select sum(IMPOML) ImportePesos, sum(IMPOM2) ImporteDolares from CGT040B where CIA='" & pCia & "' AND tipocpte='" & pTipoCbte & "' AND NUMERO=" & pNroAsiento & " and FECCPTE='" & GF_DT2DTCONTABLE(pFechaAsiento) & "' and DH='H'"
'response.Write "<br>" & strSQL
Call executeQueryDB(DBSITE_SQL_MAGIC, rsHaber, "OPEN", strSQL)
if (not rsDebe.eof) then
	'response.write "<br>D-Pesos: " & CDbl(rsDebe("ImportePesos")) 
	'response.write "<br>H-Pesos: " & CDbl(rsHaber("ImportePesos"))
	if not isnull(rsDebe("ImportePesos")) then diffPesos = round(CDbl(rsDebe("ImportePesos")) - CDbl(rsHaber("ImportePesos")), 2)
	'response.Write "<br>Dif D-H: " & diffPesos
	'response.write "<br>D-Dolar: " & CDbl(rsDebe("ImporteDolares")) 
	'response.write "<br>H-Dolar: " & CDbl(rsHaber("ImporteDolares"))		
	if not isnull(rsDebe("ImporteDolares")) then diffDolares = round(CDbl(rsDebe("ImporteDolares")) - CDbl(rsHaber("ImporteDolares")), 2)	
	'response.Write "<br>Dif D-H: " & diffDolares
    if ((diffPesos <> 0) or (diffDolares <> 0)) then
        'Elimino la diferencia en la cuenta de mayor importe.
        strSQL = "Select * from CGT040B where CIA='" & pCia & "' AND tipocpte='" & pTipoCbte & "' AND NUMERO=" & pNroAsiento & " and FECCPTE='" & GF_DT2DTCONTABLE(pFechaAsiento) & "' and DH='H' order by IMPOML desc"
        'response.Write "<br>" & strSQL
        Call executeQueryDB(DBSITE_SQL_MAGIC, rsC, "OPEN", strSQL)                
        if (not rsC.eof) then
            strSQL="Update CGT040B SET IMPOML=IMPOML+" & diffPesos & ", IMPOM1=IMPOM1+" & diffPesos & ", IMPOM2=IMPOM2+" & diffDolares & " where recno=" & rsC("recno")
            'response.Write "<br>" & strSQL
            Call executeQueryDB(DBSITE_SQL_MAGIC, rsX, "EXEC", strSQL)
        end if
    end if
end if
Response.Write "<hr>"
end sub
'----------------------------------------------------------------------------------------
function getNextCodigo(pCdObra, pCdCuenta)
dim strSQL, rs, oConn, rtrn, mes, nextNumber
strSQL ="SELECT * FROM [SAF].[dbo].FXAMPRGO WHERE IDPROY='" & pCdObra & "' AND IDOTRB='" & pCdCuenta & "' ORDER BY IDCODI DESC"
'Response.Write "<HR>" & strSQL
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if not rs.eof then 		
		nextNumber = cdbl(rs("IDCODI")) + 1		
		rtrn = string(6-len(nextNumber),"0") & nextNumber
	else
		rtrn = "000001"
	end if	
'Response.Write "<BR>" & RTRN
getNextCodigo = rtrn
end function
'----------------------------------------------------------------------------------------
sub addDetalleContableSAF(pObra, pCdCuenta, pDescripcion, pFechaAsiento, pDbCr, pImportePesos, pImporteDolares)
'dim strSQL, rs, oConn
if pDbCr = 2 then
	pImportePesos = cDbl(pImportePesos) * -1
	pImporteDolares = cDbl(pImporteDolares) * -1
end if
strSQL ="INSERT INTO [SAF].[dbo].FXAMPRGO VALUES('" & pObra & "', '" & pCdCuenta & "','" & getNextCodigo(pObra,pCdCuenta) & "','000',1,'CU','CONSUMOS DE ALMACEN','" & GF_DT2DTCONTABLE(pFechaAsiento) & "','" & pDescripcion & "','99999','',''," & cdbl(pImportePesos)/10000 & "," & cdbl(pImportePesos)/10000 & ",'01',1," & cdbl(pImporteDolares)/10000 & "," & cdbl(pImportePesos)/10000 & ",'','','0','0')"
Response.Write "<HR>Add en SAF<br>" & strSQL
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
end sub
%>