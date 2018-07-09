<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<%
dim fecha, cdCuenta, idVale, idArticulo, idCategoria, idAlmacen
fecha = GF_Parametros7("fecha", "", 6)
cdCuenta = GF_Parametros7("cdCuenta", "", 6)
cdVale = GF_Parametros7("cdVale", "", 6)
cdCategoria = GF_Parametros7("cdCategoria", "", 6)
idArticulo = GF_Parametros7("idArticulo", 0, 6)
idAlmacen = GF_Parametros7("idAlmacen", "", 6)

%>
<table width='100%' class="reg_Header">
<% if fecha = "" then %>
	<tr class="reg_Header_nav">
		<td rowspan="2" align="center"><%=GF_TRADUCIR("Cierres Realizados")%></td>
		<td colspan="2" align="center"><%=GF_TRADUCIR("Valorización")%></td>
	</tr>
	<tr class="reg_Header_nav">
		<td align="center" width="20%"><%=GF_TRADUCIR("Pesos")%></td>
		<td align="center" width="20%"><%=GF_TRADUCIR("Dólares")%></td>
	</tr>
	<%
	strSQL = "SELECT FECHACIERRE, sum(VLUPESOS) as TOTALPESOS, sum(VLUDOLARES) as TOTALDOLARES FROM TBLCIERRESARTICULOS WHERE IDALMACEN IN (" & idAlmacen & ") GROUP BY FECHACIERRE order by FECHACIERRE desc"
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	while not rs.eof
			%>
				<tr class="reg_Header_navdos" onMouseOver="javascript:lightOn(this)" onMouseOut="javascript:lightOff(this)">
					<td>
						<a style="cursor:pointer;color:blue;text-decoration:underline;" onclick="addFecha('<%=rs("FECHACIERRE")%>');"><%=rs("FECHACIERRE")%></a>
					</td>
					<td align="right"><%=GF_EDIT_DECIMALS(rs("TOTALPESOS"),2)%></td>
					<td align="right"><%=GF_EDIT_DECIMALS(rs("TOTALDOLARES"),2)%></td>
				</tr>
			<%
		rs.movenext
	wend	
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
else
	if cdCuenta <> "" then
		if cdCategoria <> "" then
			if idArticulo <> 0 then
				if cdVale <> "" or 1=1 then
					strSQL = "  select VC.IDVALE, VC.CDVALE, (VD.CANTIDAD*TAC.VLUPESOS) AS TOTALPESOS, (VD.CANTIDAD*TAC.VLUDOLARES) AS TOTALDOLARES from " & _
							 " TBLVALESCABECERA VC " & _ 
							 " INNER JOIN TBLVALESDETALLE VD on VC.IDVALE = VD.IDVALE " & _ 
							 " INNER JOIN TBLCIERRESARTICULOS TAC ON VD.IDARTICULO=TAC.IDARTICULO " & _ 
						     " WHERE TAC.IDALMACEN IN (" & idAlmacen & ") AND TAC.IDARTICULO=" & idArticulo & " AND TAC.FECHACIERRE LIKE '" & fecha & "%' AND VC.FECHA BETWEEN '" & fecha & "00' AND '" & fecha & "99' " & _
							 " order by VC.IDVALE asc"


					'Response.Write strSQL
					Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
					if rs.eof then
						%>
						<tr class="reg_Header_nav">
							<td colspan="2" align="center"><%=GF_TRADUCIR("No se encontraron resultados")%></td>
						</tr>
						<%
					else
						%>
						<tr class="reg_Header_nav">
							<td rowspan="2" align="center"><%=GF_TRADUCIR("Vales")%></td>
							<td colspan="2" align="center"><%=GF_TRADUCIR("Valorización")%></td>
						</tr>
						<tr class="reg_Header_nav">
							<td align="center" width="20%"><%=GF_TRADUCIR("Pesos")%></td>
							<td align="center" width="20%"><%=GF_TRADUCIR("Dólares")%></td>
						</tr>
						<%
					end if
					while not rs.eof
							%>
								<tr class="reg_Header_navdos" onMouseOver="javascript:lightOn(this)" onMouseOut="javascript:lightOff(this)">
									<td>
										<a class="link" onclick="openVale('<%=rs("IDVALE")%>');"><%=rs("iDVALE") & " - " & getLeyendaCdVale(rs("CDVALE")) & " (" & rs("CDVALE") & ")"%></a>
									</td>
									<td align="right"><%=GF_EDIT_DECIMALS(rs("TOTALPESOS"),2) & "&nbsp;" & getSimboloMoneda(MONEDA_PESOS)%></td>
									<td align="right"><%=GF_EDIT_DECIMALS(rs("TOTALDOLARES"),2) & "&nbsp;" & getSimboloMoneda(MONEDA_DOLAR)%></td>
								</tr>
							<%
						rs.movenext
					wend	
					Call executeQueryDB(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
				else 'No tengo un vale
					strSQL = "  select VC.CDVALE, SUM(VD.CANTIDAD*TAC.VLUPESOS) AS TOTALPESOS, SUM(VD.CANTIDAD*TAC.VLUDOLARES) AS TOTALDOLARES from " & _
							" TBLVALESCABECERA VC " & _ 
							" INNER JOIN TBLVALESDETALLE VD on VC.IDVALE = VD.IDVALE " & _ 
							" INNER JOIN TBLCIERRESARTICULOS TAC ON VD.IDARTICULO=TAC.IDARTICULO " & _ 
						    " WHERE TAC.IDALMACEN IN (" & idAlmacen & ") AND TAC.IDARTICULO=" & idArticulo & " AND VC.FECHA BETWEEN '" & fecha & "00' AND '" & fecha & "99' " & _
							" group by VC.CDVALE"
					Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
					if rs.eof then
						%>
						<tr class="TDERROR">
							<td colspan="2" align="center"><%=GF_TRADUCIR("No se encontraron resultados")%></td>
						</tr>
						<%
					else
						%>
						<tr class="reg_Header_nav">
							<td rowspan="2" align="center"><%=GF_TRADUCIR("Vales")%></td>
							<td colspan="2" align="center"><%=GF_TRADUCIR("Valorización")%></td>
						</tr>
						<tr class="reg_Header_nav">
							<td align="center" width="20%"><%=GF_TRADUCIR("Pesos")%></td>
							<td align="center" width="20%"><%=GF_TRADUCIR("Dólares")%></td>
						</tr>
						<%
					end if
					while not rs.eof
							%>
								<tr class="reg_Header_navdos" onMouseOver="javascript:lightOn(this)" onMouseOut="javascript:lightOff(this)">
									<td>
										<a class="link" onclick="addVale('<%=rs("CDVALE")%>');"><%=rs("CDVALE")%></a>
									</td>
									<td align="right"><%=GF_EDIT_DECIMALS(rs("TOTALPESOS"),2) & "&nbsp;" & getSimboloMoneda(MONEDA_PESOS)%></td>
									<td align="right"><%=GF_EDIT_DECIMALS(rs("TOTALDOLARES"),2) & "&nbsp;" & getSimboloMoneda(MONEDA_DOLAR)%></td>
								</tr>
							<%
						rs.movenext
					wend	
					Call executeQueryDB(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
				end if	
			else 'No tengo un articulo
				%>
				<tr class="reg_Header_nav">
					<td rowspan="2" align="center"><%=GF_TRADUCIR("Articulos")%></td>
					<td colspan="2" align="center"><%=GF_TRADUCIR("Valorización")%></td>
				</tr>
				<tr class="reg_Header_nav">
					<td align="center" width="20%"><%=GF_TRADUCIR("Pesos")%></td>
					<td align="center" width="20%"><%=GF_TRADUCIR("Dólares")%></td>
				</tr>
				<%
				strSQL = "select ART.IDARTICULO, ART.DSARTICULO, sum(VLUPESOS) as TOTALPESOS, sum(VLUDOLARES) as TOTALDOLARES from TBLCIERRESARTICULOS ARTC inner join tblarticulos ART " & _
						 " on ARTC.idarticulo=ART.idarticulo inner join tblartcategorias CAT  " & _
						 " on art.idcategoria = CAT.idcategoria where ARTC.IDALMACEN IN (" & idAlmacen & ") AND ARTC.FECHACIERRE like '" & fecha & "%' and CAT.cdcuenta='" & cdCuenta & "' and CAT.CDCATEGORIA=" & cdCategoria & " group by ART.IDARTICULO, ART.DSARTICULO"
				'Response.Write strSQL	
				Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
				while not rs.eof
						%>
							<tr class="reg_Header_navdos" onMouseOver="javascript:lightOn(this)" onMouseOut="javascript:lightOff(this)">
								<td>
									<a class="link" onclick="addArticulo('<%=rs("IDARTICULO")%>');"><%=rs("IDARTICULO") & "-" & rs("DSARTICULO")%></a>
								</td>
								<td align="right"><%=GF_EDIT_DECIMALS(rs("TOTALPESOS"),2) & "&nbsp;" & getSimboloMoneda(MONEDA_PESOS)%></td>
								<td align="right"><%=GF_EDIT_DECIMALS(rs("TOTALDOLARES"),2) & "&nbsp;" & getSimboloMoneda(MONEDA_DOLAR)%></td>
							</tr>
						<%
					rs.movenext
				wend	
				Call executeQueryDB(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
			end if
		else 'No tengo una categoria
			%>
			<tr class="reg_Header_nav">
				<td rowspan="2" align="center"><%=GF_TRADUCIR("Categorias")%></td>
				<td colspan="2" align="center"><%=GF_TRADUCIR("Valorización")%></td>
			</tr>
			<tr class="reg_Header_nav">
				<td align="center" width="20%"><%=GF_TRADUCIR("Pesos")%></td>
				<td align="center" width="20%"><%=GF_TRADUCIR("Dólares")%></td>
			</tr>
			<%
			strSQL = "select CAT.IDCATEGORIA, CAT.CDCATEGORIA, CAT.DSCATEGORIA, sum(VLUPESOS) as TOTALPESOS, sum(VLUDOLARES) as TOTALDOLARES from TBLCIERRESARTICULOS ARTC inner join tblarticulos ART " & _
					 " on ARTC.idarticulo=ART.idarticulo inner join tblartcategorias CAT  " & _
					 " on art.idcategoria = CAT.idcategoria where ARTC.IDALMACEN IN (" & idAlmacen & ") AND ARTC.FECHACIERRE like '" & fecha & "%' and CAT.cdcuenta='" & cdCuenta & "' group by CAT.IDCATEGORIA, CAT.CDCATEGORIA, CAT.DSCATEGORIA"
		    Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
			while not rs.eof
					%>
						<tr class="reg_Header_navdos" onMouseOver="javascript:lightOn(this)" onMouseOut="javascript:lightOff(this)">
							<td>
								<a class="link" onclick="addCategoria('<%=rs("CDCATEGORIA")%>');"><%=rs("CDCATEGORIA") & "-" & rs("DSCATEGORIA")%></a>
							</td>
							<td align="right"><%=GF_EDIT_DECIMALS(rs("TOTALPESOS"),2) & "&nbsp;" & getSimboloMoneda(MONEDA_PESOS)%></td>
							<td align="right"><%=GF_EDIT_DECIMALS(rs("TOTALDOLARES"),2) & "&nbsp;" & getSimboloMoneda(MONEDA_DOLAR)%></td>
						</tr>
					<%
				rs.movenext
			wend	
			Call executeQueryDB(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
		end if	
	else 'No tengo una cuenta
		%>
		<tr class="reg_Header_nav">
			<td rowspan="2" align="center"><%=GF_TRADUCIR("Cuentas")%></td>
			<td colspan="2" align="center"><%=GF_TRADUCIR("Valorización")%></td>
		</tr>
		<tr class="reg_Header_nav">
			<td align="center" width="20%"><%=GF_TRADUCIR("Pesos")%></td>
			<td align="center" width="20%"><%=GF_TRADUCIR("Dólares")%></td>
		</tr>
		<%
		strSQL =	" SELECT CDCUENTA, TG.TOTALPESOS, TG.TOTALDOLARES FROM TBLBUDGETOBRAS BO RIGHT JOIN " & _
					"	( " & _
					"	SELECT IDOBRA, IDBUDGETAREA, IDBUDGETDETALLE, SUM(TOTALPESOS) AS TOTALPESOS, SUM(TOTALDOLARES) AS TOTALDOLARES FROM " & _
					"	    ( " & _
					"	    SELECT IDOBRA, IDBUDGETAREA, IDBUDGETDETALLE, VD.IDARTICULO, SUM (VD.EXISTENCIA*VLUPESOS) AS TOTALPESOS, SUM (VD.EXISTENCIA*VLUDOLARES) AS TOTALDOLARES " & _
					"	        FROM TBLVALESCABECERA VC INNER JOIN TBLVALESDETALLE VD " & _
					"	            ON VC.IDVALE=VD.IDVALE  " & _
					"	                INNER JOIN TBLCIERRESARTICULOS CA  " & _
					"	                    ON VD.IDARTICULO = CA.IDARTICULO " & _
					"	                        WHERE VC.FECHA LIKE '201003%'" & _
					"	                            AND ((VC.CDVALE = '" & CODIGO_VS_SALIDA & "') OR (VC.CDVALE='" & CODIGO_VS_AJUSTE_STOCK & "' AND VD.EXISTENCIA<0)) " & _
					"	                            AND VC.ESTADO=" & ESTADO_ACTIVO & " AND VC.IDALMACEN IN (" & idAlmacen & ")" & _
					"	                            AND CA.FECHACIERRE LIKE '" & fecha & "%'" & _
					"	                            AND CA.IDALMACEN IN (" & idAlmacen & ")" & _
					"	                            GROUP BY IDOBRA, IDBUDGETAREA, IDBUDGETDETALLE, VD.IDARTICULO " & _
					"	    )T1 GROUP BY T1.IDOBRA, T1.IDBUDGETAREA, T1.IDBUDGETDETALLE " & _
					"	)TG " & _
					"	ON BO.IDOBRA=TG.IDOBRA AND BO.IDAREA=TG.IDBUDGETAREA AND BO.IDDETALLE=TG.IDBUDGETDETALLE"
							'"	                        WHERE VC.FECHA LIKE '" & fecha & "%'" & _
		
		'Response.Write strSQL
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		while not rs.eof
				%>
					<tr class="reg_Header_navdos" onMouseOver="javascript:lightOn(this)" onMouseOut="javascript:lightOff(this)">
						<td>
							<a class="link" onclick="addCuenta('<%=rs("cdcuenta")%>');"><%=rs("cdcuenta")%></a>
						</td>
						<td align="right"><%=GF_EDIT_DECIMALS(rs("TOTALPESOS"),2) & "&nbsp;" & getSimboloMoneda(MONEDA_PESOS)%></td>
						<td align="right"><%=GF_EDIT_DECIMALS(rs("TOTALDOLARES"),2) & "&nbsp;" & getSimboloMoneda(MONEDA_DOLAR)%></td>
					</tr>
				<%
			rs.movenext
		wend	
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
	end if
end if %>
</table>