<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosHKEY.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<%
Dim fechaCierre, cdCuenta, idVale, idArticulo, idAlmacen, idDivision, idCierre, idSector, estadoCierreAFirmar
Dim mesAux, auxTotalPesos, auxTotalDolares, cdCategoria, idCategoria, idObra, idBudgetArea, idBudgetDetalle, sePuedeFirmar, myAlias
Dim myOnClick, myDescription, myPesosRs, myDolaresRs, msj
fechaCierre = GF_Parametros7("fecha", "", 6)
cdCuenta = GF_Parametros7("cdCuenta", "", 6)
cdCategoria = GF_Parametros7("cdCategoria", "", 6)
estadoCierre = GF_Parametros7("estadoCierre", "", 6)
tipoCuenta = GF_Parametros7("tipoCuenta", "", 6)
idCategoria = GF_Parametros7("idCategoria", 0, 6)
idObra = GF_Parametros7("idObra", 0, 6)
idBudgetArea = GF_Parametros7("idBudgetArea", 0, 6)
idBudgetDetalle = GF_Parametros7("idBudgetDetalle", 0, 6)
idArticulo = GF_Parametros7("idArticulo", 0, 6)
idCierre = GF_Parametros7("idCierre", 0, 6)
idSector = GF_Parametros7("idSector", 0, 6)
idAlmacen = GF_Parametros7("idAlmacen", "", 6)
idDivision = GF_Parametros7("idDivision", 0, 6)
sePuedeFirmar = false
'------------------------------------------------------------------------------------------
sub imprimirTotales()
		%>
		<tr>
			<td>&nbsp;</td>
			<td align="right"><B><%=GF_EDIT_DECIMALS(auxTotalPesos,2) & "&nbsp;" & getSimboloMoneda(MONEDA_PESOS)%>		</B></td>
		</tr>
		<%	
end sub
'------------------------------------------------------------------------------------------
sub imprimirTotalesVales()
		%>
		<tr>
			<td>&nbsp;</td>
			<td align="right"><B><%=GF_EDIT_DECIMALS(auxTotalPesos,2) & "&nbsp;" & getSimboloMoneda(MONEDA_PESOS)%>		</B></td>
			<td>&nbsp;</td>
		</tr>
		<%	
end sub
'------------------------------------------------------------------------------------------
sub imprimirTitulo(pTitulo)
		%>
		<tr class="reg_Header_nav">
			<td align="center"><%=GF_TRADUCIR(pTitulo)%></td>
			<td align="center"><%=GF_TRADUCIR("Valorización")%></td>
			<td width="2%" align="center"><%=GF_TRADUCIR(".")%></td>
		</tr>
		<%	
end sub
'------------------------------------------------------------------------------------------
sub imprimirTituloVales()
		%>
		<tr class="reg_Header_nav">
			<td rowspan="2" align="center"><%=GF_TRADUCIR("Vales")%></td>
			<td colspan="2" align="center"><%=GF_TRADUCIR("Valorización")%></td>
			<td width="2%" rowspan="2" align="center"><%=GF_TRADUCIR(".")%></td>
		</tr>
		<tr class="reg_Header_nav">
			<td align="center" width="20%"><%=GF_TRADUCIR("Pesos")%></td>
			<td align="center" width="20%"><%=GF_TRADUCIR("Dólares")%></td>
		</tr>												
		<%	
end sub
'------------------------------------------------------------------------------------------
%>
<table width='100%' class="reg_Header">

<% 
if tipoCuenta = TIPO_CIERRE_DEBE then
	msj = "&nbsp;CUENTAS DEUDOR"
elseif tipoCuenta = TIPO_CIERRE_HABER then
	msj = "&nbsp;CUENTAS ACREEDOR"
end if
if msj <> "" then
%>
		<tr class="reg_Header_navs">
			<td colspan="9" align="left"><b><%=GF_TRADUCIR(msj)%></b></td>
		</tr>
<%
end if
if fechaCierre = "" then 
	'call imprimirTitulo("Cierres Realizados")
	%>
		<tr class="reg_Header_nav">
			<td width="25%" rowspan="2" align="center"><%=GF_TRADUCIR("Cierres Realizados")%></td>
			<td width="15%" colspan="2" align="center"><%=GF_TRADUCIR("Tipo Cuentas")%></td>
			<td width="10%" rowspan="2" align="center"><%=GF_TRADUCIR("Valorización")%></td>
			<td width="40%" colspan="2" align="center"><%=GF_TRADUCIR("Firmas Responsables")%></td>
			<td width="5%" rowspan="2" align="center"><%=GF_TRADUCIR("STS")%></td>
		</tr>
		<tr class="reg_Header_nav">
			<td align="center"><%=GF_TRADUCIR("Resultado")%></td>
			<td align="center"><%=GF_TRADUCIR("Inventario")%></td>
			<td align="center" width="10%"><%=GF_TRADUCIR("Contable")%></td>
			<td align="center" width="10%"><%=GF_TRADUCIR("Puerto")%></td>
		</tr>	
	<%
	auxRAC = getRolFirma(session("Usuario"), SEC_SYS_ALMACENES)
	puedeFirmarRAC = (auxRAC = FIRMA_ROL_RESP_CONTADURIA)
	puedeFirmarRAP = puedeFirmarAsientos(session("Usuario"),idDivision)
	strSQL =" SELECT TOP 24 CAB.IDCIERRE, CAB.ANIO as ANIO, CAB.MES AS MES, CAB.ESTADO , CF1.SECUENCIA AS SECUENCIA_1, CF1.CDUSUARIO AS CDUSUARIO_1, CF1.FECHAFIRMA AS FECHAFIRMA_1, CF1.HKEY AS HKEY_1, CF2.SECUENCIA AS SECUENCIA_2, CF2.CDUSUARIO AS CDUSUARIO_2, CF2.FECHAFIRMA AS FECHAFIRMA_2, CF2.HKEY AS HKEY_2, sum(IMPORTEPESOS) as TOTALPESOS " & _
			" FROM TBLCIERRESCABECERA2 CAB INNER JOIN TBLCIERRESASIENTOS2 ASI  " & _
			"    ON CAB.IDCIERRE=ASI.IDCIERRE AND CAB.IDDIVISION=" & idDivision & " AND DBCR=" & TIPO_CIERRE_DEBE & _
			"    INNER JOIN TBLCIERRESFIRMAS2 CF1 ON CAB.IDCIERRE=CF1.IDCIERRE AND CF1.SECUENCIA= " & FIRMA_ROL_RESP_CONTADURIA & _
			"    INNER JOIN TBLCIERRESFIRMAS2 CF2 ON CAB.IDCIERRE=CF2.IDCIERRE AND CF2.SECUENCIA= " & FIRMA_ROL_RESP_PUERTO & _
			" GROUP BY CAB.IDCIERRE, CAB.ANIO, CAB.MES, CAB.IDDIVISION, CAB.ESTADO, CF1.SECUENCIA, CF1.CDUSUARIO, CF1.FECHAFIRMA, CF1.HKEY, CF2.SECUENCIA, CF2.CDUSUARIO, CF2.FECHAFIRMA, CF2.HKEY " & _
			" order by CAB.ANIO DESC, CAB.MES DESC, CAB.IDDIVISION, CAB.ESTADO "
	'Response.Write strSQL
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	sePuedeFirmar = true
	while not rs.eof
		if sePuedeFirmar then 
			if rs("ESTADO") = TIPO_CIERRE_PROVISORIO AND (trim(rs("CDUSUARIO_1")) = "NAU" OR trim(rs("CDUSUARIO_2")) = "NAU") then esCierreAFirmar = true
		end if
		mesAux = rs("MES")
		if len(mesAux) = 1 then mesAux = "0" & mesAux
			%>
				<tr class="reg_Header_navdos" onMouseOver="javascript:lightOn(this)" onMouseOut="javascript:lightOff(this)">
					<td valign="top">
						<%=rs("ANIO") & "/" & mesAux%>
					</td>
					<td onclick="addFechaCierre('<%=rs("ANIO") & mesAux%>',<%=rs("IDCIERRE")%>,'<%=rs("ESTADO")%>',<%=TIPO_CIERRE_DEBE%>);" valign="top" WIDTH="3%" ALIGN="CENTER">
						<img src="images/almacenes/see_all-16x16.png" TITLE="<%=GF_Traducir("Ver cuentas de resultados")%>"></td>
					<td onclick="addFechaCierre('<%=rs("ANIO") & mesAux%>',<%=rs("IDCIERRE")%>,'<%=rs("ESTADO")%>',<%=TIPO_CIERRE_HABER%>);" valign="top" WIDTH="3%" ALIGN="CENTER">
						<img src="images/almacenes/see_all-16x16.png" TITLE="<%=GF_Traducir("Ver cuentas de inventario")%>"></td>
					</td>
					<td valign="top" align="right"><%=GF_EDIT_DECIMALS(cDbl(rs("TOTALPESOS"))/100,2) & "&nbsp;" & getSimboloMoneda(MONEDA_PESOS)%></td>
					<td align="center">	
						<%	
						'Response.Write "puedeFirmarRAC(" & puedeFirmarRAC & "), esCierreAFirmar(" & esCierreAFirmar & ")"
						firma1 = armarTextoFirma(rs("HKEY_1"), trim(rs("FECHAFIRMA_1")))
						if puedeFirmarRAC and esCierreAFirmar then								
								idCierreAFirmar = rs("IDCIERRE")
								if (firma1 = "") then		
									if getRolFirma(session("Usuario"), SEC_SYS_ALMACENES) = FIRMA_ROL_RESP_CONTADURIA then
										%>
										<br><div align="center" id="hk1"></div><br>
										<%	
										else	
										%>
										<br><br><br>
										<%	
									end if	
								else
									Response.Write "<img src='images/firmas/" & obtenerFirma(rs("CDUSUARIO_1")) & "'>"
									Response.Write firma1
								end if
						else
							'Response.Write trim(rs("CDUSUARIO_1"))
							if trim(rs("CDUSUARIO_1")) = "NAU" then
								Response.Write "<img src='images/icon_del.gif' title='No Aprobado'>"
							else
								'Response.Write "<img src='images/icon_ok.gif' title='Aprobado'>"
								Response.Write "<img src='images/firmas/" & obtenerFirma(rs("CDUSUARIO_1")) & "'><br>"
								response.Write getUserDescription(rs("CDUSUARIO_1"))
								Response.Write firma1
							end if
						end if
						%>					
					</td>	
					<td align="center">	
						<%	
						firma2 = armarTextoFirma(rs("HKEY_2"), rs("FECHAFIRMA_2"))
						if puedeFirmarRAP and esCierreAFirmar then														
								idCierreAFirmar = rs("IDCIERRE")
								if (firma2 = "") then							
									if (puedeFirmarAsientos(session("Usuario"),idDivision)) then	%>
										<br><div id="hk2"></div><br>
									<%	
									else	
									%>
										<br><br><br>
									<%	
									end if	
								else
									Response.Write "<img src='images/firmas/" & obtenerFirma(rs("CDUSUARIO_2")) & "'>"
									Response.Write firma2
								end if
						else
							if trim(rs("CDUSUARIO_2")) = "NAU" then
								Response.Write "<img src='images/icon_del.gif' title='No Aprobado'>"
							else
								'Response.Write "<img src='images/icon_ok.gif' title='Aprobado'>"								
								Response.Write "<img src='images/firmas/" & obtenerFirma(rs("CDUSUARIO_2")) & "'><br>"
								response.Write getUserDescription(rs("CDUSUARIO_2"))
								Response.Write firma2
							end if
						end if								
						%>					
					</td>	
					<%	
						if rs("ESTADO") = TIPO_CIERRE_DEFINITIVO then 
							myImage = "images/almacenes/lock-16x16.png" 
							myTitle = "DEFINITIVO"
							'sePuedeFirmar = false
						else
							myImage = "images/almacenes/edit-16x16.png" 
							myTitle = "PROVISORIO"
							'sePuedeFirmar = true
						end if	
					%>
					<td valign="top" WIDTH="3%" ALIGN="CENTER">
						<img src="<%=myImage%>" TITLE="<%=myTitle%>">
					</td>
				</tr>
			<%
			if esCierreAFirmar then 
				esCierreAFirmar = false
				sePuedeFirmar = false
			end if	
		rs.movenext
	wend	
	%>
	<input type="hidden" id="idCierreAFirmar" value="<%=idCierreAFirmar%>">
	<%
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
else 
	if cdCuenta <> "" then
		if tipoCuenta = TIPO_CUENTA_HABER then 
			if idCategoria <> 0 then 'FECHA-CUENTA-OBRA-CATEGORIA
				if idArticulo <> 0 then 'FECHA-CUENTA-OBRA-CATEGORIA-ARTICULO
					'Sin Articulo
					call imprimirTitulo("Vales")
					strLink = "openVale"
					editarVale = true
					strSQL = "SELECT VC.IDVALE AS ITEM_ID, VC.CDVALE AS ITEM_CD, VC.NRVALE AS ITEM_DS, SUM(CANTIDAD*VLUPESOS) AS TOTALPESOS " & _  		 
							 "	FROM TBLARTCTACTE CTACTE " & _ 
							 "		INNER JOIN TBLVALESCABECERA VC ON CTACTE.IDVALE=VC.IDVALE " & _
							 "	    WHERE CTACTE.IDARTICULO=" & idArticulo & " AND CTACTE.FECHACIERRE=" & fechaCierre & " AND IDDIVISION=" & idDivision & _
							 "		     AND CTACTE.TIPOVALUACION IN ('" & GASTO & "','" & PROVISION & "','" & REVERSION_PROVISION & "') AND CTACTE.CUENTAINVENTARIO='" & cdCuenta & "'" & _
							 "		GROUP BY VC.IDVALE, VC.NRVALE, VC.CDVALE  " & _
							 "		ORDER BY VC.IDVALE, VC.NRVALE, VC.CDVALE "
				else
					'Sin Articulo
					call imprimirTitulo("Articulos")
					strLink = "addArticulo"
					strSQL = "SELECT ART.IDARTICULO AS ITEM_ID, ART.IDARTICULO AS ITEM_CD, ART.DSARTICULO AS ITEM_DS, SUM(CANTIDAD*VLUPESOS) AS TOTALPESOS " & _  		 
							 "	FROM TBLARTCTACTE CTACTE " & _ 
							 "	    INNER JOIN TBLARTICULOS ART ON ART.IDARTICULO=CTACTE.IDARTICULO " & _
							 "	    INNER JOIN TBLARTCATEGORIAS CAT ON CAT.IDCATEGORIA=ART.IDCATEGORIA " & _
							 "	    WHERE CAT.IDCATEGORIA=" & idCategoria & " AND CTACTE.FECHACIERRE=" & fechaCierre & " AND IDDIVISION=" & idDivision & _
							 "		     AND CTACTE.TIPOVALUACION IN ('" & GASTO & "','" & PROVISION & "','" & REVERSION_PROVISION & "') AND CTACTE.CUENTAINVENTARIO='" & cdCuenta & "'" & _
							 "		GROUP BY ART.IDARTICULO, ART.DSARTICULO " & _
							 "		ORDER BY ART.DSARTICULO "
				end if		
			else
				'Sin categoria
				call imprimirTitulo("Categorias")
				strLink = "addCategoria"
				strSQL = "SELECT CAT.IDCATEGORIA AS ITEM_ID, CAT.CDCATEGORIA AS ITEM_CD, CAT.DSCATEGORIA AS ITEM_DS, SUM(CANTIDAD*VLUPESOS) AS TOTALPESOS " & _  		 
						 "	FROM TBLARTCTACTE CTACTE " & _ 
						 "	    INNER JOIN TBLARTICULOS ART ON ART.IDARTICULO=CTACTE.IDARTICULO " & _
						 "	    INNER JOIN TBLARTCATEGORIAS CAT ON CAT.IDCATEGORIA=ART.IDCATEGORIA " & _
						 "	    WHERE CTACTE.FECHACIERRE=" & fechaCierre & " AND IDDIVISION=" & idDivision & _
						 "		     AND CTACTE.TIPOVALUACION IN ('" & GASTO & "','" & PROVISION & "','" & REVERSION_PROVISION & "') AND CTACTE.CUENTAINVENTARIO='" & cdCuenta & "'" & _
						 "		GROUP BY CAT.IDCATEGORIA, CAT.CDCATEGORIA, CAT.DSCATEGORIA " & _
						 "		ORDER BY CAT.CDCATEGORIA, CAT.DSCATEGORIA "
			end if		
			
			'Response.write strSQL	
			'Response.End 
			'Impresion cuentas tipo haber
			Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
			while not rs.eof
					myPesosRs = 0
					if not isNull(rs("TOTALPESOS"))   then myPesosRs = rs("TOTALPESOS")
						%>
						<tr class="reg_Header_navdos" onMouseOver="javascript:lightOn(this)" onMouseOut="javascript:lightOff(this)">
							<td onclick="<%=strLink%>(<%=rs("ITEM_ID")%>,'<%=rs("ITEM_CD")%>');">
								<%=rs("ITEM_CD") & "-" & rs("ITEM_DS")%>
							</td>
							<td align="right"><%=GF_EDIT_DECIMALS(cDbl(myPesosRs)/100,2) & "&nbsp;" & getSimboloMoneda(MONEDA_PESOS)%></td>
							<td align="center">
								<% if editarVale then %>
									<a class="link" onclick="editVale('');"><img src="images/Almacenes/edit-16x16.png"></a>
								<% else 
									Response.Write "."
								   end if %>							
								
							</td>														
						</tr>
						<%
						auxTotalPesos = auxTotalPesos + cDbl(myPesosRs)/100
				rs.movenext
			wend	
			Call executeQueryDB(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
			call imprimirTotales	
		else
			if idObra <> 0 OR idSector <> 0 then 'FECHA-CUENTA-OBRA/SECTOR
				if idCategoria <> 0 then 'FECHA-CUENTA-OBRA-CATEGORIA
					if idArticulo <> 0 then 'FECHA-CUENTA-OBRA-CATEGORIA-ARTICULO
						call imprimirTitulo("Vales")
						strSQL = " SELECT '7' AS ETAPA, VC.IDVALE AS ITEM_ID, VC.CDVALE AS ITEM_CD, VC.NRVALE AS ITEM_DS, SUM(CANTIDAD*VLUPESOS) AS TOTALPESOS " & _
								 "	FROM TBLARTCTACTE CTACTE " & _ 
								 "	    INNER JOIN TBLVALESCABECERA VC ON CTACTE.IDVALE=VC.IDVALE " & _  							 
								 "		INNER JOIN TBLARTICULOS ART ON ART.IDARTICULO=CTACTE.IDARTICULO AND CTACTE.IDARTICULO=" & idArticulo & _ 
								 "	    INNER JOIN TBLBUDGETOBRAS BO ON BO.IDOBRA=VC.IDOBRA AND BO.IDAREA=VC.IDBUDGETAREA AND BO.IDDETALLE=VC.IDBUDGETDETALLE " & _  
								 "	    INNER JOIN TBLDATOSOBRAS DOB ON DOB.IDOBRA=VC.IDOBRA " & _ 
								 "	    WHERE CTACTE.FECHACIERRE=" & fechaCierre & " AND CTACTE.IDDIVISION=" & idDivision & _
								 "		     AND CTACTE.TIPOVALUACION IN ('" & GASTO & "','" & PROVISION & "','" & REVERSION_PROVISION & "') AND CTACTE.CUENTAGASTOS='" & cdCuenta & "' AND DOB.IDOBRA=" & idObra & _
								 "			 and BO.IDAREA= " & idBudgetArea & " and BO.IDDETALLE = " & idBudgetDetalle & _
								 "		GROUP BY VC.IDVALE, VC.CDVALE, VC.NRVALE " & _
								 " UNION " & _
								 " SELECT '8' AS ETAPA, VC.IDVALE AS ITEM_ID, VC.CDVALE AS ITEM_CD, VC.NRVALE AS ITEM_DS, SUM(CANTIDAD*VLUPESOS) AS TOTALPESOS " & _
								 "	FROM TBLARTCTACTE CTACTE " & _ 
								 "		INNER JOIN TBLARTICULOS ART ON ART.IDARTICULO=CTACTE.IDARTICULO AND CTACTE.IDARTICULO=" & idArticulo & _ 
								 "	    INNER JOIN TBLVALESCABECERA VC ON CTACTE.IDVALE=VC.IDVALE " & _  							 
								 "	    INNER JOIN TBLSECTORES SEC ON VC.IDSECTOR=SEC.IDSECTOR " & _ 
								 "	    WHERE CTACTE.FECHACIERRE=" & fechaCierre & " AND CTACTE.IDDIVISION=" & idDivision & _
								 "		     AND CTACTE.TIPOVALUACION IN ('" & GASTO & "','" & PROVISION & "','" & REVERSION_PROVISION & "') AND CTACTE.CUENTAGASTOS='" & cdCuenta & "' AND SEC.IDSECTOR=" & idSector & _
								 "		GROUP BY VC.IDVALE, VC.CDVALE, VC.NRVALE"															 
					else
						'Sin articulo
						call imprimirTitulo("Articulos")
						strSQL = " SELECT '5' AS ETAPA, ART.IDARTICULO AS ITEM_ID, ART.IDARTICULO AS ITEM_CD, ART.DSARTICULO AS ITEM_DS, SUM(CANTIDAD*VLUPESOS) AS TOTALPESOS " & _
								 "	FROM TBLARTCTACTE CTACTE " & _ 
								 "	    INNER JOIN TBLVALESCABECERA VC ON CTACTE.IDVALE=VC.IDVALE " & _  							 
								 "		INNER JOIN TBLARTICULOS ART ON ART.IDARTICULO=CTACTE.IDARTICULO " & _ 
								 "		INNER JOIN TBLARTCATEGORIAS CAT ON CAT.IDCATEGORIA=ART.IDCATEGORIA AND CAT.IDCATEGORIA=" & idCategoria & _
								 "	    INNER JOIN TBLBUDGETOBRAS BO ON BO.IDOBRA=VC.IDOBRA AND BO.IDAREA=VC.IDBUDGETAREA AND BO.IDDETALLE=VC.IDBUDGETDETALLE " & _  
								 "	    INNER JOIN TBLDATOSOBRAS DOB ON DOB.IDOBRA=VC.IDOBRA " & _ 
								 "	    WHERE CTACTE.FECHACIERRE=" & fechaCierre & " AND CTACTE.IDDIVISION=" & idDivision & _
								 "		     AND CTACTE.TIPOVALUACION IN ('" & GASTO & "','" & PROVISION & "','" & REVERSION_PROVISION & "') AND CTACTE.CUENTAGASTOS='" & cdCuenta & "' AND DOB.IDOBRA=" & idObra & _
								 "			 and BO.IDAREA= " & idBudgetArea & " and BO.IDDETALLE = " & idBudgetDetalle & _
								 "		GROUP BY ART.IDARTICULO, ART.DSARTICULO " & _
								 " UNION " & _
								 " SELECT '6' AS ETAPA, ART.IDARTICULO AS ITEM_ID, ART.IDARTICULO AS ITEM_CD, ART.DSARTICULO AS ITEM_DS, SUM(CANTIDAD*VLUPESOS) AS TOTALPESOS " & _
								 "	FROM TBLARTCTACTE CTACTE " & _ 
								 "		INNER JOIN TBLARTICULOS ART ON ART.IDARTICULO=CTACTE.IDARTICULO " & _ 
								 "		INNER JOIN TBLARTCATEGORIAS CAT ON CAT.IDCATEGORIA=ART.IDCATEGORIA AND CAT.IDCATEGORIA=" & idCategoria & _
								 "	    INNER JOIN TBLVALESCABECERA VC ON CTACTE.IDVALE=VC.IDVALE " & _  							 
								 "	    INNER JOIN TBLSECTORES SEC ON VC.IDSECTOR=SEC.IDSECTOR " & _ 
								 "	    WHERE CTACTE.FECHACIERRE=" & fechaCierre & " AND CTACTE.IDDIVISION=" & idDivision & _
								 "		     AND CTACTE.TIPOVALUACION IN ('" & GASTO & "','" & PROVISION & "','" & REVERSION_PROVISION & "') AND CTACTE.CUENTAGASTOS='" & cdCuenta & "' AND SEC.IDSECTOR=" & idSector & _
								 "		GROUP BY ART.IDARTICULO, ART.DSARTICULO"							
					end if		
				else
					'Sin categoria
					call imprimirTitulo("Categorias")
					strSQL = " SELECT '3' AS ETAPA, CAT.IDCATEGORIA AS ITEM_ID, CAT.CDCATEGORIA AS ITEM_CD, CAT.DSCATEGORIA AS ITEM_DS, SUM(CANTIDAD*VLUPESOS) AS TOTALPESOS " & _
							 "	FROM TBLARTCTACTE CTACTE " & _ 
							 "	    INNER JOIN TBLVALESCABECERA VC ON CTACTE.IDVALE=VC.IDVALE " & _  							 
							 "		INNER JOIN TBLARTICULOS ART ON ART.IDARTICULO=CTACTE.IDARTICULO " & _ 
							 "		INNER JOIN TBLARTCATEGORIAS CAT ON CAT.IDCATEGORIA=ART.IDCATEGORIA " & _
							 "	    INNER JOIN TBLBUDGETOBRAS BO ON BO.IDOBRA=VC.IDOBRA AND BO.IDAREA=VC.IDBUDGETAREA AND BO.IDDETALLE=VC.IDBUDGETDETALLE " & _  
							 "	    INNER JOIN TBLDATOSOBRAS DOB ON DOB.IDOBRA=VC.IDOBRA " & _ 
							 "	    WHERE CTACTE.FECHACIERRE=" & fechaCierre & " AND CTACTE.IDDIVISION=" & idDivision & _
							 "		     AND CTACTE.TIPOVALUACION IN ('" & GASTO & "','" & PROVISION & "','" & REVERSION_PROVISION & "') AND CTACTE.CUENTAGASTOS='" & cdCuenta & "' AND DOB.IDOBRA=" & idObra & _
							 "			 and BO.IDAREA= " & idBudgetArea & " and BO.IDDETALLE = " & idBudgetDetalle & _
							 "		GROUP BY CAT.IDCATEGORIA, CAT.CDCATEGORIA, CAT.DSCATEGORIA " & _
							 " UNION " & _
							 " SELECT '4' AS ETAPA, CAT.IDCATEGORIA AS ITEM_ID, CAT.IDCATEGORIA AS ITEM_CD, CAT.DSCATEGORIA AS ITEM_DS, SUM(CANTIDAD*VLUPESOS) AS TOTALPESOS " & _
							 "	FROM TBLARTCTACTE CTACTE " & _ 
							 "		INNER JOIN TBLARTICULOS ART ON ART.IDARTICULO=CTACTE.IDARTICULO " & _ 
							 "		INNER JOIN TBLARTCATEGORIAS CAT ON CAT.IDCATEGORIA=ART.IDCATEGORIA " & _
							 "	    INNER JOIN TBLVALESCABECERA VC ON CTACTE.IDVALE=VC.IDVALE " & _  							 
							 "	    INNER JOIN TBLSECTORES SEC ON VC.IDSECTOR=SEC.IDSECTOR " & _ 
							 "	    WHERE CTACTE.FECHACIERRE=" & fechaCierre & " AND CTACTE.IDDIVISION=" & idDivision & _
							 "		     AND CTACTE.TIPOVALUACION IN ('" & GASTO & "','" & PROVISION & "','" & REVERSION_PROVISION & "') AND CTACTE.CUENTAGASTOS='" & cdCuenta & "' AND SEC.IDSECTOR=" & idSector & _
							 "		GROUP BY CAT.IDCATEGORIA, CAT.DSCATEGORIA"					
				end if					
			else 'SIN OBRA
				if trim(cdCuenta) = trim(CUENTA_AJUSTE_STOCK) then 'FECHA-CUENTAAJUSTES
					if idCategoria <> 0 then 'FECHA-CUENTAAJUSTE-CATEGORIA
						if idArticulo <> 0 then 'FECHA-CUENTAAJUSTE-CATEGORIA-ARTICULO
								call imprimirTitulo("Vales")
								strSQL = " SELECT '7' AS ETAPA, VC.IDVALE AS ITEM_ID, VC.CDVALE AS ITEM_CD, VC.NRVALE AS ITEM_DS, SUM(CANTIDAD*VLUPESOS) AS TOTALPESOS " & _
										 "	FROM TBLARTCTACTE CTACTE " & _ 
										 "	    INNER JOIN TBLVALESCABECERA VC ON CTACTE.IDVALE=VC.IDVALE " & _  							 
										 "		INNER JOIN TBLARTICULOS ART ON ART.IDARTICULO=CTACTE.IDARTICULO AND CTACTE.IDARTICULO=" & idArticulo & _ 
										 "	    WHERE CTACTE.FECHACIERRE=" & fechaCierre & " AND CTACTE.IDDIVISION=" & idDivision & _
										 "		     AND CTACTE.TIPOVALUACION IN ('" & GASTO & "','" & PROVISION & "','" & REVERSION_PROVISION & "') AND CTACTE.CUENTAGASTOS='" & cdCuenta & "'" & _
										 "		GROUP BY VC.IDVALE, VC.CDVALE, VC.NRVALE "
    							'response.write strSQL	
    							'Response.End 
						else
							call imprimirTitulo("Artículos")
							strSQL = " SELECT '5' AS ETAPA, ART.IDARTICULO AS ITEM_ID, ART.IDARTICULO AS ITEM_CD, ART.DSARTICULO AS ITEM_DS, SUM(CANTIDAD*VLUPESOS) AS TOTALPESOS " & _
									 "	FROM TBLARTCTACTE CTACTE " & _ 
									 "	    INNER JOIN TBLVALESCABECERA VC ON CTACTE.IDVALE=VC.IDVALE " & _  							 
									 "		INNER JOIN TBLARTICULOS ART ON ART.IDARTICULO=CTACTE.IDARTICULO " & _ 
									 "		INNER JOIN TBLARTCATEGORIAS CAT ON CAT.IDCATEGORIA=ART.IDCATEGORIA AND CAT.IDCATEGORIA=" & idCategoria & _
									 "	    WHERE CTACTE.FECHACIERRE=" & fechaCierre & " AND CTACTE.IDDIVISION=" & idDivision & _
									 "		     AND CTACTE.TIPOVALUACION IN ('" & GASTO & "','" & PROVISION & "','" & REVERSION_PROVISION & "') AND CTACTE.CUENTAGASTOS='" & cdCuenta & "'" & _
									 "		GROUP BY ART.IDARTICULO, ART.DSARTICULO "
    						'response.write strSQL	
    						'Response.End 
						end if
					else
						call imprimirTitulo("Categorías")						
						strSQL ="SELECT '3' AS ETAPA, CAT.IDCATEGORIA AS ITEM_ID, CAT.CDCATEGORIA AS ITEM_CD, CAT.DSCATEGORIA AS ITEM_DS, SUM(CANTIDAD*VLUPESOS) AS TOTALPESOS " & _
							 "	FROM TBLARTCTACTE CTACTE " & _ 
							 "	    INNER JOIN TBLVALESCABECERA VC ON CTACTE.IDVALE=VC.IDVALE " & _  							 
							 "		INNER JOIN TBLARTICULOS ART ON ART.IDARTICULO=CTACTE.IDARTICULO " & _ 
							 "		INNER JOIN TBLARTCATEGORIAS CAT ON CAT.IDCATEGORIA=ART.IDCATEGORIA " & _
							 "	    WHERE CTACTE.FECHACIERRE=" & fechaCierre & " AND CTACTE.IDDIVISION=" & idDivision & _
							 "		     AND CTACTE.TIPOVALUACION IN ('" & GASTO & "','" & PROVISION & "','" & REVERSION_PROVISION & "') AND CTACTE.CUENTAGASTOS='" & cdCuenta & "'" & _
							 "		GROUP BY CAT.IDCATEGORIA, CAT.CDCATEGORIA, CAT.DSCATEGORIA " 
    					'response.write strSQL	
    					'Response.End 
					end if				
				else
					'Solo la cuenta
					call imprimirTitulo("Obras/Sectores")
					strSQL = " SELECT '1' AS ETAPA, DOB.IDOBRA AS ITEM_ID, DOB.CDOBRA AS ITEM_CD, BO.DSBUDGET AS ITEM_DS, VC.IDBUDGETAREA AS ITEM_AUX1, VC.IDBUDGETDETALLE AS ITEM_AUX2, SUM(CANTIDAD*VLUPESOS) AS TOTALPESOS " & _
							 "	FROM TBLARTCTACTE CTACTE " & _ 
							 "	    INNER JOIN TBLVALESCABECERA VC ON CTACTE.IDVALE=VC.IDVALE " & _  							 
							 "	    INNER JOIN TBLBUDGETOBRAS BO ON BO.IDOBRA=VC.IDOBRA AND BO.IDAREA=VC.IDBUDGETAREA AND BO.IDDETALLE=VC.IDBUDGETDETALLE " & _  
							 "	    INNER JOIN TBLDATOSOBRAS DOB ON DOB.IDOBRA=VC.IDOBRA " & _ 
							 "	    WHERE CTACTE.FECHACIERRE=" & fechaCierre & " AND CTACTE.IDDIVISION=" & idDivision & _
							 "		     AND CTACTE.TIPOVALUACION IN ('" & GASTO & "','" & PROVISION & "','" & REVERSION_PROVISION & "') AND CTACTE.CUENTAGASTOS='" & cdCuenta & "'" & _
							 "		GROUP BY DOB.IDOBRA, DOB.CDOBRA, BO.DSBUDGET, VC.IDBUDGETAREA, VC.IDBUDGETDETALLE " & _
							 " UNION " & _
							 " SELECT '2' AS ETAPA, SEC.IDSECTOR AS ITEM_ID, CAST(SEC.IDSECTOR AS VARCHAR(20)) AS ITEM_CD, SEC.DSSECTOR AS ITEM_DS, 0 AS ITEM_AUX1, 0 AS ITEM_AUX2, SUM(CANTIDAD*VLUPESOS) AS TOTALPESOS " & _
							 "	FROM TBLARTCTACTE CTACTE " & _ 
							 "	    INNER JOIN TBLVALESCABECERA VC ON CTACTE.IDVALE=VC.IDVALE " & _  							 
							 "	    INNER JOIN TBLSECTORES SEC ON VC.IDSECTOR=SEC.IDSECTOR " & _ 
							 "	    WHERE CTACTE.FECHACIERRE=" & fechaCierre & " AND CTACTE.IDDIVISION=" & idDivision & _
							 "		     AND CTACTE.TIPOVALUACION IN ('" & GASTO & "','" & PROVISION & "','" & REVERSION_PROVISION & "') AND CTACTE.CUENTAGASTOS='" & cdCuenta & "'" & _
							 "		GROUP BY SEC.IDSECTOR, SEC.DSSECTOR" &_
							 " UNION " & _
							 " SELECT '3' AS ETAPA, CAT.IDCATEGORIA AS ITEM_ID, CAT.CDCATEGORIA AS ITEM_CD, CAT.DSCATEGORIA AS ITEM_DS, 0 AS ITEM_AUX1, 0 AS ITEM_AUX2, SUM(CANTIDAD*VLUPESOS) AS TOTALPESOS " & _
							 "	FROM TBLARTCTACTE CTACTE " & _ 							  
							 "	    INNER JOIN TBLVALESCABECERA VC ON CTACTE.IDVALE=VC.IDVALE " & _
							 "	    INNER JOIN TBLARTICULOS ART ON ART.IDARTICULO=CTACTE.IDARTICULO " & _
							 "	    INNER JOIN TBLARTCATEGORIAS CAT ON CAT.IDCATEGORIA=ART.IDCATEGORIA " & _							 
							 "	    WHERE CTACTE.FECHACIERRE=" & fechaCierre & " AND CTACTE.IDDIVISION=" & idDivision & _
							 "		     AND CTACTE.TIPOVALUACION IN ('" & GASTO & "','" & PROVISION & "','" & REVERSION_PROVISION & "') AND CTACTE.CUENTAGASTOS='" & cdCuenta & "' and CDVALE in ('VMR', 'VMT', 'AJT')" & _
							 "		GROUP BY CAT.IDCATEGORIA, CAT.CDCATEGORIA, CAT.DSCATEGORIA"
				end if			 
						    
			end if	
			'Response.write strSQL	
			'Response.End 
			Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
			while not rs.eof
					myPesosRs = 0
					if not isNull(rs("TOTALPESOS"))   then myPesosRs = rs("TOTALPESOS")
					if rs("ETAPA") = 2 then 
						myOnClick = "addSector(" & rs("ITEM_ID") & ");"
					elseif rs("ETAPA") = 1 then 
						myOnClick = "addObra(" & rs("ITEM_ID") & ",'" & rs("ITEM_CD") & "'," & rs("ITEM_AUX1") & "," & rs("ITEM_AUX2") & ");"
					elseif rs("ETAPA") = 3 or rs("ETAPA") = 4 then 
						myOnClick = "addCategoria(" & rs("ITEM_ID") & "," & rs("ITEM_CD") & ")" 
					elseif rs("ETAPA") = 5 or rs("ETAPA") = 6 then 
						myOnClick = "addArticulo(" & rs("ITEM_ID") & "," & rs("ITEM_CD") & ")" 						
					elseif rs("ETAPA") = 7 or rs("ETAPA") = 8 then 
						myOnClick = "openVale(" & rs("ITEM_ID") & ",'" & rs("ITEM_CD") & "')" 						
						editarVale = true
					end if 							
						%>
						<tr class="reg_Header_navdos" onMouseOver="javascript:lightOn(this)" onMouseOut="javascript:lightOff(this)">
							<td onclick="<%=myOnClick%>">
								<%=rs("ITEM_CD") & "-" & rs("ITEM_DS")%>
							</td>
							<td onclick="<%=myOnClick%>" align="right"><%=GF_EDIT_DECIMALS(cDbl(myPesosRs)/100,2) & "&nbsp;" & getSimboloMoneda(MONEDA_PESOS)%></td>
							<td align="center">
								<% if editarVale then %>
									<a class="link"><img src="images/Almacenes/edit-16x16.png"></a>
								<% else 
									Response.Write "."
								   end if %>	
							</td>														
						</tr>
						<%
						auxTotalPesos = auxTotalPesos + cDbl(myPesosRs)/100
				rs.movenext
			wend	
			Call executeQueryDB(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
			call imprimirTotales	
		end if
		
								
	else 
		'Sin cuenta
		call imprimirTitulo("CUENTAS")
		strSQL ="SELECT CDCUENTA, NOMCUE AS DSCUENTA, SUM(IMPORTEPESOS) AS TOTALPESOS " & _
				"	FROM TBLCIERRESASIENTOS2 " & _
				"		LEFT JOIN [Database].[dbo].[CGT020A] ON CIA='" &  getCIADivision(idDivision) & "' AND LEFT(CUENTA,8) COLLATE Modern_Spanish_CI_AS = CONVERT(VARCHAR(8),CDCUENTA) " & _
				"		WHERE IDCIERRE = " & idCierre & " AND DBCR=" & tipoCuenta & _
				" GROUP BY CDCUENTA,NOMCUE ORDER BY CDCUENTA	"		
			
			
			'Response.write strSQL	
			'Response.End 
			Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
			while not rs.eof
					myPesosRs = 0
					if not isNull(rs("TOTALPESOS"))   then myPesosRs = rs("TOTALPESOS")
						%>
						<tr onclick="addCuenta('<%=rs("CDCUENTA")%>');" class="reg_Header_navdos" onMouseOver="javascript:lightOn(this)" onMouseOut="javascript:lightOff(this)">
							<td>
								<%=formatCuentaPantalla(rs("CDCUENTA")) & " - " & rs("DSCUENTA")%>
							</td>
							<td align="right"><%=GF_EDIT_DECIMALS(cDbl(myPesosRs)/100,2) & "&nbsp;" & getSimboloMoneda(MONEDA_PESOS)%></td>
							<td align="center">
								.
							</td>														
						</tr>
						<%
						auxTotalPesos = auxTotalPesos + cDbl(myPesosRs)/100
				rs.movenext
			wend	
			Call executeQueryDB(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
			call imprimirTotales 	
	end if
					
end if 
%>
</table>
