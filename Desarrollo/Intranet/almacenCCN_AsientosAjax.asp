<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<%
dim fechaCierre, tipoCierre, strSQL
dim rs, con, rtrn, anio, mes, idDivision, cdCuentaDEBE, cdCuentaHABER, etapa
dim centroCosto, idCierre, flagHayAsientos
'Response.Write "Hecho..."
'Response.End 
idDivision = GF_Parametros7("idDivision", 0, 6)
fechaCierre = GF_Parametros7("fechaCierre", "", 6)
tipoCierre = GF_Parametros7("tipoCierre", "", 6) 
if tipoCierre = "" then tipoCierre = "P"
anio = left(fechaCierre,4)
mes = mid(fechaCierre,5,2)
idCierre = getIdCierre2(anio, mes, idDivision, tipoCierre)
'LIMPIAR ASIENTOS ANTERIORES
'Response.Write "CIERRE(" & idCierre & ")"
strSQL = "DELETE FROM TBLCIERRESASIENTOS2 WHERE IDCIERRE=" & idCierre 
Call executeQueryDB(DBSITE_SQL_INTRA, rsINS, "EXEC", strSQL)

flagHayAsientos=false

'Leer desde la cuenta corriente 
'DEBE
	strSQL = "SELECT CUENTAINVENTARIO, CCOSTOS, SUM(VLUPESOS*CANTIDAD) AS TOTAL_CUENTA, SUM(VLUDOLARES*CANTIDAD) AS TOTAL_CUENTA_D FROM TBLARTCTACTE " & _
			 "	WHERE FECHACIERRE=" & fechaCierre & " AND IDDIVISION=" & idDivision & _
			 "		AND TIPOVALUACION IN('" & GASTO & "','" & PROVISION & "','" & REVERSION_PROVISION & "') " & _
			 "	GROUP BY CUENTAINVENTARIO, CCOSTOS"
	Response.Write "<BR>" & strSQL
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	while not rs.eof
	        flagHayAsientos=true
			Call insertAsiento(TIPO_CUENTA_HABER, rs("CUENTAINVENTARIO"), rs("CCOSTOS"), CDBL(rs("TOTAL_CUENTA")), CDBL(rs("TOTAL_CUENTA_D")))
		rs.movenext
	wend	
'HABER	
	strSQL = "SELECT CUENTAGASTOS, CCOSTOS, SUM(VLUPESOS*CANTIDAD) AS TOTAL_CUENTA, SUM(VLUDOLARES*CANTIDAD) AS TOTAL_CUENTA_D FROM TBLARTCTACTE " & _
			 "	WHERE FECHACIERRE=" & fechaCierre & " AND IDDIVISION=" & idDivision & _
			 "		AND TIPOVALUACION IN('" & GASTO & "','" & PROVISION & "','" & REVERSION_PROVISION & "') AND CUENTAGASTOS<>'VRS' " & _
			 "	GROUP BY CUENTAGASTOS, CCOSTOS"
	Response.Write "<BR>" & strSQL
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	while not rs.eof
	        flagHayAsientos=true
			call insertAsiento(TIPO_CUENTA_DEBE, rs("CUENTAGASTOS"), rs("CCOSTOS"), CDBL(rs("TOTAL_CUENTA")), CDBL(rs("TOTAL_CUENTA_D")))
		rs.movenext
	wend	
Response.Write "<br>Hecho..."		
'----------------------------------------------------------------------------------------
sub insertAsiento(pDebeHaber, pCdCuenta, pCentroCosto, pImportePesos, pImporteDolares)
dim strSQL, rs, con, sqlINS, rsINS, conINS
if isNull(pImporte) then pImporte = 0

if cDbl(pImportePesos) <> 0 then
	strSQL = "SELECT * FROM TBLCIERRESASIENTOS2 WHERE IDCIERRE=" & idCierre & " AND CDCUENTA='" & pCdCuenta & "' AND CCOSTOS='" & pCentroCosto & "' AND DBCR=" & pDebeHaber
	Response.Write "<BR>" & strSQL
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if rs.eof then
		sqlINS = "INSERT INTO TBLCIERRESASIENTOS2 VALUES(" & idCierre & ",'" & pCdCuenta & _
			"'," & pDebeHaber & "," & pImportePesos & ", " & pImporteDolares & ",0,'" & pCentroCosto & "')"
	else
		sqlINS ="UPDATE TBLCIERRESASIENTOS2 SET IMPORTEPESOS=IMPORTEPESOS + " & pImportePesos & ", IMPORTEDOLARES=IMPORTEDOLARES + " & pImporteDolares &_ 
				" WHERE IDCIERRE=" & idCierre & " AND CDCUENTA='" & pCdCuenta & "' AND CCOSTOS='" & pCentroCosto & "' AND DBCR=" & pDebeHaber
	end if
	Response.Write "<BR>" & sqlINS
	Call executeQueryDB(DBSITE_SQL_INTRA, rsINS, "EXEC", sqlINS)
end if	
end sub
'----------------------------------------------------------------------------------------

%>