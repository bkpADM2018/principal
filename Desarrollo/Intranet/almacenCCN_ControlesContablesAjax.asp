<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosHKEY.asp"-->
<%
dim fecha, mesAux, myDescription, idAlmacen
idDivision = GF_Parametros7("idDivision", 0, 6)
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
			<td align="center"><%=GF_TRADUCIR("Art Sin Precio")%></td>
			<td align="center"><%=GF_TRADUCIR("STS")%></td>
		</tr>
	<%
	strSQL =" SELECT CAB.IDCIERRE, CAB.ANIO as ANIO, CAB.MES AS MES, CAB.ESTADO, sum(IMPORTEPESOS) as TOTALPESOS, sum(IMPORTEDOLARES) as TOTALDOLARES " & _
			" FROM TBLCIERRESCABECERA CAB INNER JOIN TBLCIERRESASIENTOS ASI  " & _
			"    ON CAB.IDCIERRE=ASI.IDCIERRE  " & _
			"    WHERE CAB.ESTADO='" & TIPO_CIERRE_PROVISORIO & "' AND CAB.IDDIVISION=" & idDivision & " AND DBCR=" & TIPO_CIERRE_DEBE & _
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
					<td onclick="verArtSinPrecio('<%=rs("ANIO") & mesAux%>',<%=rs("IDCIERRE")%>);" valign="top" ALIGN="CENTER">
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
	<table align="center" align="center" width="70%" border="0">
		<tr>
			<td align="right">
				<img style="cursor:pointer;" src='images/back-16x16.png' ONCLICK="javascript:window.location.href='almacenControlesContables.asp?idDivision=<%=idDivision%>';">
			</td>
		</tr>
	</table>	
	<table width='70%' class="reg_Header">
		<tr class="reg_Header_nav">
			<td align="center"><%=GF_TRADUCIR("Id")%></td>
			<td align="center"><%=GF_TRADUCIR("Descripcion")%></td>
			<td align="center"><%=GF_TRADUCIR("Existencia")%></td>
		</tr>
	<%	
	strSQL= "SELECT * FROM TBLCIERRESARTICULOS CIE " & _
			 "   INNER JOIN " & _ 
			 "       ( " & _
			 "       SELECT IDARTICULO FROM TBLVALESCABECERA CAB " & _ 
			 "           INNER JOIN TBLVALESDETALLE DET ON CAB.IDVALE=DET.IDVALE " & _
			 "           WHERE CAB.CDVALE IN ('VMS','AJU','AJS') AND CAB.FECHA LIKE '" & fecha & "%'" & _
			 "               AND IDALMACEN IN (" & idAlmacen & ") AND DET.EXISTENCIA<>0 AND ESTADO=" & ESTADO_ACTIVO & _
			 "               group by IDARTICULO " & _
			 "        ) T1 " & _
			 "   ON CIE.IDARTICULO=T1.IDARTICULO " & _
			 "   INNER JOIN TBLARTICULOS ART ON ART.IDARTICULO=CIE.IDARTICULO " & _
			" WHERE CIE.VLUPESOS=0 AND " & _ 
			" CIE.FECHACIERRE like '" & fecha & "%' AND CIE.IDALMACEN IN (" & idAlmacen & ")"
	'Response.Write strSQL
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if rs.eof then
		%><td class="TDERROR" colspan="3"><%=GF_Traducir("No se encontraron resultados")%></td><%
	else
		while not rs.eof	
			%>
			<TR class="reg_Header_navdos" onMouseOver="javascript:lightOn(this)" onMouseOut="javascript:lightOff(this)">	
				<td align="center"><%=rs("IDARTICULO")%></td>
				<td align="left"><%=rs("DSARTICULO")%></td>
				<td align="center"><%=rs("EXISTENCIA")%></td>				
			</TR>	
			<%	
			rs.movenext
		wend	
	end if
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
end if
%>
</table>
