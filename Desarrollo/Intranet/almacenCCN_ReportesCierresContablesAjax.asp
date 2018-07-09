<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosHKEY.asp"-->
<%
dim fecha, mesAux, myDescription, idDivision
idDivision = GF_Parametros7("idDivision", 0, 6)
fecha = GF_Parametros7("fecha", "", 6)
'------------------------------------------------------------------------------------------
%>
<table width='40%' class="reg_Header">
<% if fecha = "" then 
	'call imprimirTitulo("Cierres Realizados")
	%>
		<tr class="reg_Header_nav">
			<td align="center"><%=GF_TRADUCIR("CIERRES REALIZADOS")%></td>
			<td width="20%" align="center"><%=GF_TRADUCIR("ACCIONES")%></td>
			<td width="15%" align="center"><%=GF_TRADUCIR("STS")%></td>
		</tr>
	<%
	strSQL =" SELECT CAB.IDCIERRE, CAB.ANIO as ANIO, CAB.MES AS MES, CAB.ESTADO, sum(IMPORTEPESOS) as TOTALPESOS, sum(IMPORTEDOLARES) as TOTALDOLARES " & _
			" FROM TBLCIERRESCABECERA2 CAB INNER JOIN TBLCIERRESASIENTOS2 ASI  " & _
			"    ON CAB.IDCIERRE=ASI.IDCIERRE  " & _
			"    WHERE CAB.IDDIVISION=" & idDivision & " AND DBCR=" & TIPO_CIERRE_DEBE & _
			" GROUP BY CAB.IDCIERRE, CAB.ANIO, CAB.MES, CAB.IDDIVISION, CAB.ESTADO " & _
			" order by CAB.ANIO DESC, CAB.MES DESC, CAB.IDDIVISION DESC, CAB.ESTADO ASC"
	'Response.Write strSQL
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	while not rs.eof
		mesAux = rs("MES")
		if len(mesAux) = 1 then mesAux = "0" & mesAux
		
			%>
				<tr class="reg_Header_navdos" onMouseOver="javascript:lightOn(this)" onMouseOut="javascript:lightOff(this)">
					<td valign="center" ALIGN="CENTER">
						<%=rs("ANIO") & "/" & mesAux%>
					</td>
					<td valign="top" ALIGN="CENTER">
						<img src="images/almacenes/see_all-16x16.png" onclick="realizarConsulta('<%=idDivision%>', '<%=rs("ANIO") & mesAux & getLastDayOfMonth(rs("ANIO"),mesAux)%>');" TITLE="<%=GF_Traducir("Ver Cuentas")%>">
						&nbsp;
						<img src="images/almacenes/printer-16x16.png" onclick="armarPDF('<%=idDivision%>', '<%=rs("ANIO") & mesAux & getLastDayOfMonth(rs("ANIO"),mesAux)%>');" TITLE="<%=GF_Traducir("Imprimir Cuentas")%>">
					</td>
					<TD valign="top" WIDTH="3%" ALIGN="CENTER">
					<% if rs("ESTADO") = TIPO_CIERRE_PROVISORIO then %>
						<img src="images/almacenes/edit-16x16.png" TITLE="PROVISORIO">
					<% else %>
						<img src="images/almacenes/lock-16x16.png" TITLE="DEFINITIVO">
					<% end if %>
					</TD>
				</tr>
			<%
		rs.movenext
	wend	
	%>
		<tr class="reg_Header_navdos" onMouseOver="javascript:lightOn(this)" onMouseOut="javascript:lightOff(this)">
			<td valign="center" ALIGN="CENTER">
				<%="2010/12"%>
			</td>
			<td valign="center" ALIGN="CENTER">
				<img src="images/almacenes/see_all-16x16.png" onclick="realizarConsulta('<%=idDivision%>', '20101231');" TITLE="<%=GF_Traducir("Ver Cuentas")%>">
				&nbsp;
				<img src="images/almacenes/printer-16x16.png" onclick="armarPDF('<%=idDivision%>', '20101231');" TITLE="<%=GF_Traducir("Imprimir Cuentas")%>">				
			</td>
			<td valign="center" WIDTH="3%" ALIGN="CENTER">
				<img src="images/almacenes/lock-16x16.png" TITLE="DEFINITIVO">
			</td>
		</tr>	
	<input type="hidden" id="idCierreAFirmar" value="<%=idCierreAFirmar%>">
	<%
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
else
	%>

	<table width='70%' class="reg_Header">
		<tr class="reg_Header_naav">
			<td colspan="4" align="left"><b>Cierre realizado el dia <%=right(fecha,2)%> del mes de <%=getNameOfMonth(mid(fecha,5,2))%> del año <%=left(fecha,4)%></b></td>
		</tr>
		<tr class="reg_Header_nav">
			<td width="25%" align="center"><%=GF_TRADUCIR("Cuenta")%>	</td>
			<td width="25%" align="center"><%=GF_TRADUCIR("Total Pesos")%>	</td>
		</tr>
	<%	
	strSQL= "SELECT TG.CUENTA AS CDCUENTA, NOMCUE AS DSCUENTA, SUM(IMPORTE) AS TOTAL_CUENTA FROM " & _
			"	( " & _
			"	SELECT ART.IDARTICULO, ART.DSARTICULO, CASE WHEN (ART.CDCUENTA='' OR ART.CDCUENTA IS NULL) THEN CAT.CDCUENTA ELSE ART.CDCUENTA END AS CUENTA, (STOCKDISPONIBLE*VLUPESOS) AS IMPORTE " & _
			"		FROM TBLARTVALUACION VAL " & _ 
			"			INNER JOIN TBLARTICULOS ART ON VAL.IDARTICULO=ART.IDARTICULO AND (ART.CDCUENTA LIKE '1141%' OR  ART.CDCUENTA = '') " & _
			"			INNER JOIN TBLARTCATEGORIAS CAT ON ART.IDCATEGORIA=CAT.IDCATEGORIA AND (CAT.CDCUENTA LIKE '1141%' OR CAT.CDCUENTA = '')" & _
			"				WHERE FECHACIERRE=" & fecha & " AND IDDIVISION=" & idDivision & " AND VAL.STOCKDISPONIBLE<>0 " & _
			"	) TG " & _			
			"	LEFT JOIN [Database].[dbo].[CGT020A] CGT ON TG.CUENTA = CONVERT(VARCHAR(12),CGT.CUENTA) COLLATE Modern_Spanish_CI_AS " & _
			"	and CIA = '" & getCIADivision(idDivision) & "'" & _
			"	WHERE TG.CUENTA <> '' " & _
			"	GROUP BY TG.CUENTA, NOMCUE " & _
			"ORDER BY TG.CUENTA"

	'Response.Write strSQL
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if rs.eof then
		%><td class="TDERROR" colspan="3"><%=GF_Traducir("No se encontraron resultados")%></td><%
	else
		while not rs.eof	
			%>
			<tr class='reg_Header_navdos' onclick="document.location.href='almacenCCN_ReportesCierresContablesXLS.asp?idDivision=<%=idDivision%>&cuentacontable=<%=rs("CDCUENTA")%>&fecCierre=<%=fecha%>'" onMouseOver='javascript:lightOn(this)' onMouseOut='javascript:lightOff(this)'>	
				<td align="LEFT"><%=rs("CDCUENTA") & "-" & rs("DSCUENTA")%></td>
				<td align="right"><% if rs("TOTAL_CUENTA") <> "" then Response.write GF_EDIT_DECIMALS(cDbl(rs("TOTAL_CUENTA"))/100,2) & "&nbsp;" & getSimboloMoneda(MONEDA_PESOS)%>	</td>
			</tr>	
			<%	
				totalCuentas = totalCuentas + cDbl(rs("TOTAL_CUENTA"))/100
			rs.movenext
		wend	
	end if
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
	%>
			<tr class='reg_Header_navdos'>	
				<td align="right"><%=GF_Traducir("Total en Inventario Pañol:")%></td>
				<td style="cursor:pointer;" align="right" title="Todas" onclick="document.location.href='almacenCCN_ReportesCierresContablesXLS1.asp?idDivision=<%=idDivision%>&fecCierre=<%=fecha%>'"><% if totalCuentas <> 0 then Response.write GF_EDIT_DECIMALS(cDbl(totalCuentas),2) & "&nbsp;" & getSimboloMoneda(MONEDA_PESOS)%>	</td>
			</tr>	
	</TABLE>
	<table align="center" align="center" width="70%" border="0">
	<tr>
		<td align="right">
			<A style="color:blue;cursor:pointer;" ONCLICK="javascript:window.location.href='almacenCCN_ReportesCierresContables.asp?idDivision=<%=idDivision%>';">[Volver]</A>
		</td>
	</tr>
	</table>	
	<%
end if
%>
</table>
