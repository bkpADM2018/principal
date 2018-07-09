<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosHKEY.asp"-->
<%
dim fecha, mesAux, myDescription, idAlmacen, idCierre
idDivision = GF_Parametros7("idDivision", 0, 6)
idCierre = GF_Parametros7("idCierre", 0, 6)
idAlmacen = GF_Parametros7("idAlmacen", "", 6)
fecha = GF_Parametros7("fecha", "", 6)
'------------------------------------------------------------------------------------------
%>
<table width='50%' class="reg_Header">
<% if fecha = "" then 
	'call imprimirTitulo("Cierres Realizados")
	%>
		<tr class="reg_Header_nav">
			<td align="center"><%=GF_TRADUCIR("Cierres Realizados")%></td>
			<td align="center"><%=GF_TRADUCIR("Control Cruzado")%></td>
			<td align="center"><%=GF_TRADUCIR("STS")%></td>
		</tr>
	<%
	strSQL =" SELECT CAB.IDCIERRE, CAB.ANIO as ANIO, CAB.MES AS MES, CAB.ESTADO, sum(IMPORTEPESOS) as TOTALPESOS, sum(IMPORTEDOLARES) as TOTALDOLARES " & _
			" FROM TBLCIERRESCABECERA CAB INNER JOIN TBLCIERRESASIENTOS ASI  " & _
			"    ON CAB.IDCIERRE=ASI.IDCIERRE  " & _
			"    WHERE CAB.IDDIVISION=" & idDivision & " AND DBCR=" & TIPO_CIERRE_DEBE & _
			" GROUP BY CAB.IDCIERRE, CAB.ANIO, CAB.MES, CAB.IDDIVISION, CAB.ESTADO " & _
			" order by CAB.ANIO, CAB.MES, CAB.IDDIVISION, CAB.ESTADO ASC"
	'Response.Write strSQL
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	while not rs.eof
		mesAux = rs("MES")
		if len(mesAux) = 1 then mesAux = "0" & mesAux
			%>
				<tr class="reg_Header_navdos" onMouseOver="javascript:lightOn(this)" onMouseOut="javascript:lightOff(this)">
					<td valign="top">
						<%=rs("ANIO") & "/" & mesAux%>
					</td>
					<td onclick="verCruzado('<%=rs("ANIO") & mesAux%>',<%=rs("IDCIERRE")%>);" valign="top" ALIGN="CENTER">
						<img src="images/almacenes/see_all-16x16.png" TITLE="<%=GF_Traducir("Ver articulos")%>">
					</td>
					<td valign="top" WIDTH="3%" ALIGN="CENTER">
						<img src="images/almacenes/edit-16x16.png" TITLE="PROVISORIO">
					</td>
				</tr>
			<%
		rs.movenext
	wend	
	%>
	<input type="hidden" id="idCierreAFirmar" value="<%=idCierreAFirmar%>">
	<%
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
else
	%>

	<table width='70%' class="reg_Header">
		<tr class="reg_Header_naav">
			<td colspan="4" align="center">Cierre realizado el mes <%=right(fecha,2)%> del año <%=left(fecha,4)%></td>
		</tr>
		<tr class="reg_Header_nav">
			<td width="25%" align="center"><%=GF_TRADUCIR("Cuenta - Vales")%>	</td>
			<td width="25%" align="center"><%=GF_TRADUCIR("Cuenta - Asientos")%></td>
			<td width="25%" align="center"><%=GF_TRADUCIR("Pesos - Vales")%>	</td>
			<td width="25%" align="center"><%=GF_TRADUCIR("Pesos - Asientos")%>	</td>
		</tr>
	<%	
	strSQL= "SELECT CDCUENTACAT, CDCUENTA, TOTAL_VAL, TOTAL_ASI, TOTAL_VAL-TOTAL_ASI AS DIF FROM " & _
			"	( " & _
			"	SELECT TG.CDCUENTACAT, SUM(TG.EXISTENCIA) AS TOTAL_VAL FROM " & _
			"	    ( " & _
			"	        SELECT CAT.CDCUENTA AS CDCUENTACAT, SUM(DET.EXISTENCIA*DET.VLUPESOS) as EXISTENCIA " & _ 
			"	            FROM TBLVALESCABECERA CAB " & _ 
			"	                INNER JOIN TBLVALESDETALLE DET ON CAB.IDVALE = DET.IDVALE " & _
			"	                INNER JOIN TBLARTICULOS ART ON ART.IDARTICULO = DET.IDARTICULO " & _
			"	                INNER JOIN TBLARTCATEGORIAS CAT ON ART.IDCATEGORIA = CAT.IDCATEGORIA " & _ 
			"	                WHERE CAB.ESTADO=" & ESTADO_ACTIVO & " AND CAB.CDVALE IN('" & CODIGO_VS_SALIDA & "', '" & CODIGO_VS_AJUSTE_VALE & "','" & CODIGO_VS_AJUSTE_TRANSFERENCIA & "') AND CAB.IDALMACEN IN (" & idAlmacen & ")" & _
			"	                AND CAB.FECHA like '" & fecha & "%' " & _
			"	            GROUP BY CAT.CDCUENTA " & _ 
			"	    UNION " & _
			"	        SELECT CAT.CDCUENTA AS CDCUENTACAT, SUM(DET.EXISTENCIA*DET.VLUPESOS) as EXISTENCIA " & _
			"	            FROM TBLVALESCABECERA CAB " & _ 
			"	                INNER JOIN TBLVALESDETALLE DET ON CAB.IDVALE = DET.IDVALE " & _
			"	                INNER JOIN TBLARTICULOS ART ON ART.IDARTICULO = DET.IDARTICULO " & _
			"	                INNER JOIN TBLARTCATEGORIAS CAT ON ART.IDCATEGORIA = CAT.IDCATEGORIA " & _ 
			"	                WHERE CAB.ESTADO=" & ESTADO_ACTIVO & " AND CAB.CDVALE = '" & CODIGO_VS_AJUSTE_STOCK & "' AND DET.EXISTENCIA<>0 AND CAB.IDALMACEN IN (" & idAlmacen & ") AND CAB.FECHA LIKE '" & fecha & "%' " & _
			"	            GROUP BY CAT.CDCUENTA " & _
			"	    ) TG " & _
			"	    GROUP BY  TG.CDCUENTACAT " & _
			"	) T1 " & _
			"	LEFT JOIN " & _ 
			"	( " & _
			"	    SELECT CDCUENTA , SUM(IMPORTEPESOS) AS TOTAL_ASI FROM TBLCIERRESASIENTOS WHERE IDCIERRE=" & idCierre & " GROUP BY CDCUENTA " & _
			"	) T2 " & _
			"	ON T1.CDCUENTACAT=T2.CDCUENTA "

	'Response.Write strSQL
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if rs.eof then
		%><td class="TDERROR" colspan="3"><%=GF_Traducir("No se encontraron resultados")%></td><%
	else
		while not rs.eof	
			myStyle=" class='reg_Header_navdos' onMouseOver='javascript:lightOn(this)' onMouseOut='javascript:lightOff(this)'"
			if not isnull(rs("DIF")) then
				'Response.Write "(" & clng(rs("DIF")) & ")"
				if clng(rs("DIF")) <> 0 then myStyle = " class='TDERROR'"
			end if	
			%>
			<tr <%=myStyle%>>	
				<td align="center"><%=rs("CDCUENTACAT")%></td>
				<td align="center"><%=rs("CDCUENTA")%></td>
				<td align="right"><% if rs("TOTAL_VAL") <> "" then Response.write GF_EDIT_DECIMALS(clng(rs("TOTAL_VAL")),2) & "&nbsp;" & getSimboloMoneda(MONEDA_PESOS)%>	</td>
				<td align="right"><% if rs("TOTAL_ASI") <> "" then Response.Write GF_EDIT_DECIMALS(clng(rs("TOTAL_ASI")),2) & "&nbsp;" & getSimboloMoneda(MONEDA_PESOS)%>	</td>				
			</tr>	
			<%	
			rs.movenext
		wend	
	end if
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
	%>
	</TABLE>
	<table align="center" align="center" width="70%" border="0">
	<tr>
		<td align="right">
			<A style="color:blue;cursor:pointer;" ONCLICK="javascript:window.location.href='almacenControlCruzadoContable.asp?idDivision=<%=idDivision%>';">[Volver]</A>
		</td>
	</tr>
	</table>	
	<%
end if
%>
</table>
