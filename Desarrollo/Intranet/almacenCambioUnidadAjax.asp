<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<%
dim idArticulo, factorConversion, strSQL, rs, conn, myVluPesos, myVluDolares
idArticulo 	= GF_PARAMETROS7("idArticulo",0,6)
factorConversion 	= GF_PARAMETROS7("factorConversion",4,6)
strSQL =" SELECT * FROM TBLARTICULOSDATOS AD " & _
		"	 INNER JOIN TBLALMACENES AL ON AD.IDALMACEN=AL.IDALMACEN  " & _
		"    INNER JOIN TBLDIVISIONES DI ON AL.IDDIVISION=DI.IDDIVISION  " & _
		"    INNER JOIN  " & _
		"        ( " & _
		"         SELECT TEMP2.* FROM  " & _
		"            (SELECT MAX(MMTOPRECIO) AS MAXIMAFECHA, IDDIVISION FROM TBLARTICULOSPRECIOS WHERE IDARTICULO=" & idArticulo & " GROUP BY IDDIVISION) TEMP  " & _
		"             INNER JOIN TBLARTICULOSPRECIOS TEMP2 ON TEMP.MAXIMAFECHA=TEMP2.MMTOPRECIO AND TEMP.IDDIVISION=TEMP2.IDDIVISION WHERE IDARTICULO=" & idArticulo & _
		"         )  " & _
		"         AP ON AD.IDARTICULO=AP.IDARTICULO AND AL.IDDIVISION=AP.IDDIVISION " & _
		" WHERE AD.IDARTICULO=" & idArticulo & " AND (AD.EXISTENCIA<>0 OR AD.SOBRANTE<>0) " & _
		" ORDER BY DI.DSDIVISION" 
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
'------------------------------------------------------------
function getNuevosPrecios(pIdArticulo, pIdDivision, pExistencia, pFactorConversion, byref pVluPesos, byref pVluDolares)
dim strSQL, rsPrecios, conn, myVluPesos, myVluDolares
strSQL = "select * from tblarticulosprecios where idarticulo = " & pIdArticulo & _
		 " and mmtoprecio = (select max(mmtoprecio) from tblarticulosprecios where idarticulo = "&pIdArticulo&")" & _
		 " and iddivision = " & pIdDivision
Call executeQueryDB(DBSITE_SQL_INTRA, rsPrecios, "OPEN", strSQL)
myVluPesos = cdbl(rsPrecios("VLUPESOS"))
myVluDolares = cdbl(rsPrecios("VLUDOLARES")) 
stockOri = pExistencia
stockNuevo = cdbl(stockOri) * cdbl(pFactorConversion)
pVluPesos = cdbl(myVluPesos) * (cdbl(stockOri) / cdbl(stockNuevo))
pVluDolares = cdbl(myVluDolares) * (cdbl(stockOri) / cdbl(stockNuevo))
end function
%>
		<table BORDER=0 class="ui-widget-content ui-corner-all " align="center" width="60%">
			<tr>
				<td WIDTH="20%" colspan=1 rowspan="2" align="center" class="reg_header_nav">
					División
				</td>
				<td WIDTH="20%" colspan=1 rowspan="2" align="center" class="reg_header_nav">
					Almacén
				</td>
				<td WIDTH="10%" colspan=2 align="center" class="reg_header_nav">
					Stk Actual
				</td>
				<td WIDTH="10%" colspan=2 align="center" class="reg_header_nav">
					Stk Nuevo
				</td>
				<td WIDTH="20%" colspan=2 align="center" class="reg_header_nav">
					Precio Actual
				</td>
				<td WIDTH="20%" colspan=2 align="center" class="reg_header_nav">
					Nuevo Precio
				</td>
			</tr>
			<tr>
				<td width="5%" align="center" class="reg_header_nav">
					EXI
				</td>
				<td width="5%" align="center" class="reg_header_nav">
					SOB
				</td>
				<td width="5%" align="center" class="reg_header_nav">
					EXI
				</td>
				<td width="5%" align="center" class="reg_header_nav">
					SOB
				</td>
				<td width="10%" align="center" class="reg_header_nav">
					$
				</td>
				<td width="10%" align="center" class="reg_header_nav">
					U$S
				</td>
				<td width="10%" align="center" class="reg_header_nav">
					$
				</td>
				<td width="10%" align="center" class="reg_header_nav">
					U$S
				</td>
			</tr>
			<%
			while not rs.eof%>
				<tr>
					<td align="center" class="reg_header_navdos"><%=rs("DSDIVISION")%></td>
					<td align="center" class="reg_header_navdos"><%=rs("DSALMACEN")%></td>
					<td align="center" class="reg_header_navdos"><%=GF_EDIT_DECIMALS(CDbl(rs("EXISTENCIA"))*100,2)%></td>
					<td align="center" class="reg_header_navdos"><%=GF_EDIT_DECIMALS(CDbl(rs("SOBRANTE"))*100,2)%></td>
					<td align="center" class="reg_header_navdos"><%=GF_EDIT_DECIMALS(CDbl(rs("EXISTENCIA"))*cdbl(factorConversion)*100,2)%></td>
					<td align="center" class="reg_header_navdos"><%=GF_EDIT_DECIMALS(CDbl(rs("SOBRANTE"))*CDbl(factorConversion)*100,2)%></td>
					<td align="center" class="reg_header_navdos"><%=getSimboloMoneda(MONEDA_PESO) & " " & GF_EDIT_DECIMALS(CDbl(rs("VLUPESOS")),2)	%></td>
					<td align="center" class="reg_header_navdos"><%=getSimboloMoneda(MONEDA_DOLAR) & " " & GF_EDIT_DECIMALS(CDbl(rs("VLUDOLARES")),2)	%></td>
					<%
					if cdbl(rs("EXISTENCIA"))>0 then call getNuevosPrecios(idArticulo, rs("IDDIVISION"), rs("EXISTENCIA"), factorConversion, myVluPesos, myVluDolares)
					%>
					<td align="center" class="reg_header_navdos"><%=getSimboloMoneda(MONEDA_PESOS) & " " & GF_EDIT_DECIMALS(CDbl(myVluPesos),2)%></td>
					<td align="center" class="reg_header_navdos"><%=getSimboloMoneda(MONEDA_DOLAR) & " " & GF_EDIT_DECIMALS(CDbl(myVluDolares),2)%></td>
				</tr>	
				<%
				rs.movenext
			wend%>
		</table>
