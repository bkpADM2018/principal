<%
'----------------------------------------------------------------------------------
'***************************************************
'************  COMIENZO DE LA PAGINA  **************
'***************************************************
'Esta pagina se llama desde comprasCTCPopUp  *******
'Lista los contratos asignados a una obra    *******
'***************************************************
Dim rsCTC, conn, strSQL,   myCTC_idObra

myCTC_idObra = GF_PARAMETROS7("idObra",0,6)

strSQL = "Select CTC.*, PCT.CDPEDIDO from TBLOBRACONTRATOS CTC "
strSQL = strSQL & " inner join TBLPCTCABECERA PCT on CTC.IDPEDIDO = PCT.IDPEDIDO "
strSQL = strSQL & " where CTC.IDOBRA = " & myCTC_idObra & " order by IDCONTRATO"
Call executeQueryDb(DBSITE_SQL_INTRA, rsCTC, "OPEN", strSQL)
%>
<script type="text/javascript">
	function lightOn(tr) {
		tr.className = "reg_Header_navdosHL";
	}
		
	function lightOff(tr) {
		tr.className = "reg_Header_navdos";
	}

	function abrirCTC(id){
		window.open("comprasCTC.asp?idContrato=" + id, "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes",false);
	}

	function abrirPCT(id){
		window.open("comprasPedidoCotizacion.asp?idPedido=" + id, "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes,height=500,width=700",false);
	}

	function abrirObra(id) {
		window.open("comprasTableroObra.asp?idObra=" + id, "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes,height=600,width=900",false);
	}
</script>
<table width="100%" align="center" class="reg_Header">
	<tr><td colspan="6" style="cursor:pointer;" onclick="abrirObra(<%=myCTC_idObra%>)"><h4><% =GF_TRADUCIR("Obra") %>: <% =getDescripcionObra(myCTC_idObra) %></h4></td></tr>	
	<tr class="reg_Header_nav">
		<td align="center" width="30%" colspan="2"><% =GF_TRADUCIR("Nro") %></td>
		<td align="center" width="20%" colspan="2"><% =GF_TRADUCIR("Pedido") %></td>
		<td align="center" width="40%"><% =GF_TRADUCIR("Importe") %></td>
	</tr>
<%	if (not rsCTC.eof) then		%>
<%		while (not rsCTC.eof)	%>
			<tr class="reg_Header_navdos" onMouseOver="javascript:lightOn(this)" onMouseOut="javascript:lightOff(this)" style="cursor:pointer;">
				<td onclick="abrirCTC(<%=rsCTC("IDCONTRATO")%>)"><% =rsCTC("CDCONTRATO") %></td>
				<td width="3%" align="center" onclick="abrirCTC(<%=rsCTC("IDCONTRATO")%>)"><img src="images/compras/CTC-16X16.png"></td>
				<td onclick="abrirPCT(<%=rsCTC("IDPEDIDO")%>)"><% =rsCTC("CDPEDIDO") %></td>
				<td width="3%" onclick="abrirPCT(<%=rsCTC("IDPEDIDO")%>)"><img src="images/compras/PCT-16X16.png"></td>
				<td align="right" onclick="abrirCTC(<%=rsCTC("IDCONTRATO")%>)">
<%					response.write getSimboloMoneda(MONEDA_DOLAR) & " " & GF_EDIT_DECIMALS(rsCTC("IMPORTEDOLARES"),2) %>
				</td>
			</tr>
<%			rsCTC.MoveNext()
		wend	
	else	%>
		<tr><td class="TDNOHAY" colspan="6"><% =GF_TRADUCIR("La obra no registra Contratos") %></td></tr>
<%	end if	%>
	</table>