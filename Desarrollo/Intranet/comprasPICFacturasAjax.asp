<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<%
'----------------------------------------------------------------------------------
Function armarDetalle(rsFAC, codMone, minuta, linea, ByRef totalPesos, ByRef totalDolares) 
	Dim auxMinuta, ret, classStyle
	ret = "<tr>"
	ret = ret & "<td colspan='2'></td>"
	ret = ret & "	<td colspan='6' align='left'>"
	ret = ret & "		<div id='detalleArt_FAC_" & linea & "' style='display:none;'>"
	ret = ret & "			<table width='100%'>"
	ret = ret & "				<tr class='reg_Header_nav'>"
	ret = ret & "					<td align='center'>" & GF_TRADUCIR("ID Art.") & "</td>"
	ret = ret & "					<td align='center'>" & GF_TRADUCIR("Descripci�n") & "</td>"
	ret = ret & "					<td align='center'>" & GF_TRADUCIR("P.Presup.") & "</td>"
	if (codMone = MONEDA_PESO) then
		ret = ret & "					<td align='center'>" & GF_TRADUCIR("$/U") & "</td>"
	else
		ret = ret & "					<td align='center'>" & GF_TRADUCIR("u$s/U") & "</td>"
	end if
	ret = ret & "					<td align='center'>" & GF_TRADUCIR("Cant.") & "</td>"
	if (codMone = MONEDA_PESO) then	
		ret = ret & "					<td align='center'>" & GF_TRADUCIR("Imp. Pesos") & "</td>"
	else
		ret = ret & "					<td align='center'>" & GF_TRADUCIR("Imp. Dolares") & "</td>"
	end if
	ret = ret & "				</tr>"
	auxMinuta = minuta
	While ((not rsFAC.eof) and (minuta = auxMinuta))
		classStyle = "TDEXTERNOS"
		if (rsFAC("TIPOCAT") <> TIPO_CAT_IMPUESTOS) then
			totalPesos = totalPesos + cDbl(rsFAC("TOTALPESOS"))
			totalDolares = totalDolares + cDbl(rsFAC("TOTALDOLAR"))
			classStyle = "reg_Header_navdos"
		end if
		ret = ret & "<tr class='" & classStyle & "'>"
		ret = ret & "	<td align='center'><b>" & rsFAC("IDARTICULO") & "</b></td>"		
		ret = ret & "	<td align='left'>" & rsFAC("DSARTICULO") & "</td>"
		ret = ret & "	<td align='center'>" & rsFAC("IDAREA") & "-" & rsFAC("IDDETALLE") & "</td>"
		if (codMone = MONEDA_PESO) then
			ret = ret & "	<td align='right'>" & getSimboloMoneda(MONEDA_PESOS) & "&nbsp;" & GF_EDIT_DECIMALS(cDbl(rsFAC("PPUNIDAD"))*100,2) & "</td>"
		else
			ret = ret & "	<td align='right'>" & getSimboloMoneda(MONEDA_DOLAR) & "&nbsp;" & GF_EDIT_DECIMALS(cDbl(rsFAC("PDUNIDAD"))*100,2) & "</td>"
		end if
		ret = ret & "	<td align='right'>" & rsFAC("CANTIDAD") & " " & rsFAC("UNIDAD") & " </td>"
		if (codMone = MONEDA_PESO) then
			ret = ret & "	<td align='right'>" & getSimboloMoneda(MONEDA_PESOS) & "&nbsp;" & GF_EDIT_DECIMALS(cDbl(rsFAC("TOTALPESOS"))*100,2) & "</td>"
		else
			ret = ret & "	<td align='right'>" & getSimboloMoneda(MONEDA_DOLAR) & "&nbsp;" & GF_EDIT_DECIMALS(cDbl(rsFAC("TOTALDOLAR"))*100,2) & "</td>"
		end if
		ret = ret & "</tr>"		
		rsFAC.MoveNext 
		if (not rsFAC.eof) then 
			auxMinuta = CDbl(rsFAC("MINUTA"))
		else
			auxMinuta = 0
		end if
	Wend 
	ret = ret & "</table></div></td></tr>"
	armarDetalle = ret
end Function

'----------------------------------------------------------------------------------
Function mostrarFacturas(idPIC, codMone, picTotalPesos, picTotalDolares)
	dim rsFAC, reg, minuta, detalle, cbte
	dim rsArt, conn, cdInterno, totalpesos, totaldolares, acumPesos, acumDolares
	Dim fechaFactura, dsProveedor, idProveedorOld, tipocbte, nroOP
	Set rsFAC = getRsFAC(idPIC)
%>	
	<table align="center" width="90%" class="reg_Header">
		<tr>
			<td align="center" colspan="8"><b><% =GF_TRADUCIR("Facturas asociadas al PIC") %></b></td>
		</tr>
		<tr class="reg_Header_nav">
			<td align="center"><% =GF_TRADUCIR(".") %></td>
			<td align="center"><% =GF_TRADUCIR("O.Pago/Cobro") %></td>
			<td align="center"><% =GF_TRADUCIR("Minuta") %></td>
			<td align="center"><% =GF_TRADUCIR("Fecha Pago") %></td>
			<td align="center"><% =GF_TRADUCIR("Comprobante") %></td>
			<td align="center"><% =GF_TRADUCIR("Proveedor") %></td>			
<%			if (codMone = MONEDA_PESO) then	%>			
			<td align="center"><% =GF_TRADUCIR("Total Pesos") %></td>
<%			else		%>			
			<td align="center"><% =GF_TRADUCIR("Total Dolares") %></td>
<%			end if		%>			
			<!-- <td align="center"><% =GF_TRADUCIR(".") %></td> -->
		</tr>

<%
	reg=0
	idProveedorOld=0
	acumPesos= 0
	acumDolares= 0
	if (not rsFAC.eof) then
		'Proceso las facturas.
		while (not rsFAC.eof)
			reg=reg+1
			tipocbte = rsFAC("TCBTE")
			minuta = CDbl(rsFAC("MINUTA"))	
			nroOP = rsFAC("OP")
			cbte = GF_EDIT_CBTE(rsFAC("CBTE"))
			idProveedor = CDbl(rsFAC("IDPROVEEDOR"))
			'Mientras el proveedor sea el mismo no busco la descripcion.
			if (idProveedor <> idProveedorOld) then
				idProveedorOld = idProveedor
				dsProveedor = getDescripcionProveedor(idProveedor)
			end if
			fechaFactura = GF_FN2DTE(rsFAC("FECHA"))			
			totalpesos = 0
			totaldolares= 0
			detalle = armarDetalle(rsFAC, codMone, minuta, reg, totalpesos, totaldolares)			
			'Las notas de credito restan. (Esto se comento dado que los importes de la ACD7REP para las NCR ya vienen en negativo)
			'if (tipocbte = PREFIX_NCR) then	
			'	totalpesos = -1 * totalpesos
			'	totaldolares = -1 * totaldolares
			'end if
			acumPesos= acumPesos + totalpesos
			acumDolares= acumDolares + totaldolares
			
		%>							
		<tr class="reg_Header_navdos" onMouseOver="javascript:lightOn(this, '')" onMouseOut="javascript:lightOff(this, '')">
			<td align="center"><img style="cursor:pointer;" onclick="verDetalle(this,'<%="detalleArt_FAC_" & reg%>')" title="<%=GF_TRADUCIR("Detalle Articulos")%>" src="images/Mas.gif"></td>
			<td align="center"	onclick="abrirFAC()"><% =nroOP %></td>
			<td align="center"	onclick="abrirFAC()"><% =minuta %></td>
			<td align="center"	onclick="abrirFAC()"><% =fechaFactura %></td>
			<td align="center"	onclick="abrirFAC()"><% =tipocbte & " " & cbte %></td>
			<td align="left"	onclick="abrirFAC()"><% =idProveedor & " - " & dsProveedor %></td>			
<%			if (codMone = MONEDA_PESO) then	%>
			<td align="right"	onclick="abrirFAC()"><% =getSimboloMoneda(MONEDA_PESOS) & "&nbsp;" & GF_EDIT_DECIMALS(cDbl(totalpesos)*100,2) %></td>
<%			else		%>
			<td align="right"	onclick="abrirFAC()"><% =getSimboloMoneda(MONEDA_DOLAR) & "&nbsp;" & GF_EDIT_DECIMALS(cDbl(totaldolares)*100,2) %></td>
<%			end if		%>			
			<!--<td align="center" width="3%"><img style="cursor:pointer;" onclick="abrirFAC()" title="<%=GF_TRADUCIR("Ver Factura")%>" src="images/compras/Invoice-16x16.png"></td>-->
		</tr>					
	<%	Response.Write detalle
		wend	
		acumPesos= acumPesos * 100
		acumDolares= acumDolares * 100		
		%>
		<tr><td colspan="7"><hr></td></tr>
		<tr>
			<td class="reg_header_navdos" align="right" colspan="6"><font size="+1"><b>TOTAL PAGADO</b></font></td>
<%			if (codMone = MONEDA_PESO) then	%>
			<td align="right"><font size="+1"><b><% =getSimboloMoneda(MONEDA_PESOS) & "&nbsp;" & GF_EDIT_DECIMALS(acumPesos,2) %></b></font></td>
<%			else		%>					
			<td align="right"><font size="+1"><b><% =getSimboloMoneda(MONEDA_DOLAR) & "&nbsp;" & GF_EDIT_DECIMALS(acumDolares,2) %></b></font></td>
<%			end if		%>			
		</tr>	
		<tr>
			<td class="reg_header_navdos" align="right" colspan="6"><font size="+1"><b>SALDO A PAGAR</b></font></td>
<%			if (codMone = MONEDA_PESO) then	%>			
			<td align="right"><font size="+1"><b><% =getSimboloMoneda(MONEDA_PESOS) & "&nbsp;" & GF_EDIT_DECIMALS((picTotalPesos - acumPesos),2) %></b></font></td>
<%			else		%>			
			<td align="right"><font size="+1"><b><% =getSimboloMoneda(MONEDA_DOLAR) & "&nbsp;" & GF_EDIT_DECIMALS((picTotalDolares - acumDolares),2) %></b></font></td>
<%			end if		%>						
		</tr>						
<%		end if 
		if (reg = 0) then %>
		<tr><td class="TDNOHAY" colspan="8"><% =GF_TRADUCIR("No se encontraron datos para mostrar") %></td></tr>
	<% end if %>	
	</table>
<%
End Function
'----------------------------------------------------------------------------------
Function getRsFAC(idPIC)
	dim strSQL, conn, rs
	
	strSQL = "Select	  C.nroInt as MINUTA," &_
						" OP.nroordpag as OP, " &_
						" CASE WHEN D.tipcbt = "& CBTE_PROVEEDORES_FAC &" then '" & PREFIX_FAC & "' when D.tipcbt = "& CBTE_PROVEEDORES_NDB &" then '" & PREFIX_NDB & "' else '" & PREFIX_NCR & "' end	AS TCBTE, " &_						
						" Format(D.succbt, '0000') + Format(D.nrocbt, '00000000') as CBTE," &_
						" convert(varchar(10), OP.fecvto, 112)  as FECHA," &_
						" D.nrovende as IDPROVEEDOR," &_
						" C.IDARTICULO as IDARTICULO," &_
						" A.DSARTICULO as DSARTICULO," &_
						" U.ABREVIATURA as UNIDAD," &_
						" C.CANTIDAD as CANTIDAD," &_
						" C.ImporteUniPesos as PPUNIDAD," &_
						" C.ImporteUniDolares as PDUNIDAD, " &_
						" C.ImportePesos as TOTALPESOS," &_
						" C.ImporteDolares as TOTALDOLAR, " &_
						" C.IDAREA as IDAREA," &_
						" C.IDDETALLE as IDDETALLE, " &_
						" Z.TIPOCATEGORIA as TIPOCAT" 
	strSQL = strSQL & " From (Select * from VWMEP001C Where IDPIC = " & idPIC & ") C "
	strSQL = strSQL & "		Inner Join VWCOMPROBANTES D on C.nroInt = D.nroInt "
	strSQL = strSQL & "		left Join [Database].dbo.MEP006A OP on OP.nroInt = C.nroInt "
	strSQL = strSQL & "		Inner join TBLARTICULOS A on C.idarticulo = A.IDARTICULO"
	strSQL = strSQL & "		Inner join TBLARTCATEGORIAS Z on Z.IDCATEGORIA = A.IDCATEGORIA"
	strSQL = strSQL & "		Inner join TBLUNIDADES U on A.IDUNIDAD = U.IDUNIDAD"
	strSQL = strSQL & " Order by D.fecvto, C.nroInt "		
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set getRsFAC = rs
End Function
'----------------------------------------------------------------------------------
'*************************************************
'************** COMIENZO DE LA PAGINA ************
'*************************************************
dim idPIC, cdMoneda

idPIC = GF_PARAMETROS7("id", 0, 6)
cdMoneda = GF_PARAMETROS7("CM", "", 6)
picTotalPesos = GF_PARAMETROS7("pesos", 0, 6)
picTotalDolares = GF_PARAMETROS7("dolares", 0, 6)

if (idPIC > 0) then 
	Call mostrarFacturas(idPIC, cdMoneda, picTotalPesos, picTotalDolares)	
end if

%>