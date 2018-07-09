<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosTraducir.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->
<%
'----------------------------------------------------------------------------------
Function mostrarAjustes(idContrato,cdMoneda)
	dim rsAJU
	Set rsAJU = getRsAJU(idContrato)
	
%>	
	<table align="center" width="100%" class="reg_Header">
		<tr>
			<td align="center" colspan="8"><b><% =GF_TRADUCIR("Ajustes realizados al Contrato") %></b></td>
		</tr>
		<tr class="reg_Header_nav">
			<td align="center"><% =GF_TRADUCIR("Id") %></td>
			<td align="center"><% =GF_TRADUCIR("Tipo") %></td>
			<td align="center"><% =GF_TRADUCIR("Fecha") %></td>
			<td align="center"><% =GF_TRADUCIR("Importe") %></td>
			<td align="center"><% =GF_TRADUCIR("Observaciones") %></td>
			<td width="5%" align="center"><% =GF_TRADUCIR("Apl.") %></td>
			<td width="2%" align="center"><% =GF_TRADUCIR("Elim.") %></td>
		</tr>

<%
	reg=0
	if (not rsAJU.eof) then
		while (not rsAJU.eof)
			reg = reg + 1
			if (rsAJU("TIPOAJUSTE") = CTC_AJUSTE_GENERAL) then 
		        strCaption = "Ajuste del Presupuesto del Contrato"
		    else
		        strCaption = "Ajuste del Valor Unitario"
		    end if
			%>							
			<tr class="reg_Header_navdos" onMouseOver="javascript:lightOn(this, '')" onMouseOut="javascript:lightOff(this, '')" title="<% =strCaption %>">
				<td align="center"><%=rsAJU("IDAJUSTE")%></td>				
				<td align="center"><%=rsAJU("TIPOAJUSTE")%></td>
				<td align="center"><%=GF_FN2DTE(Left(rsAJU("MOMENTO"), 8)) %></td>
<%				if (cdMoneda = MONEDA_PESO) then		%>
				<td align="RIGHT"><% =getSimboloMoneda(MONEDA_PESO) & "&nbsp;" & GF_EDIT_DECIMALS(cDbl(rsAJU("IMPORTEPESOS")),2) %></td>
<%				else			%>				
				<td align="RIGHT"><% =getSimboloMoneda(MONEDA_DOLAR) & "&nbsp;" & GF_EDIT_DECIMALS(cDbl(rsAJU("IMPORTEDOLARES")),2) %></td>
<%				end if			%>				
				<td TITLE="<%=rsAJU("OBSERVACIONES")%>">
					<% 
					if len(rsAJU("OBSERVACIONES")) > 30 then
						Response.Write LEFT(rsAJU("OBSERVACIONES"),27) & "..."
					else
						Response.Write rsAJU("OBSERVACIONES")
					end if
					%>
					</td>
				
					<% if rsAJU("APLICADO") = TIPO_AFIRMACION then %>
						<td align="center">
							<img title="Ajuste aprobado" src="images/compras/accept-16x16.png">
						</td>	
						<td>&nbsp;</td>
					<% else %>
						<td align="center" title="A la espera de aprobación" align="center">
							<img src="images/compras/reception_waiting-16x16.png">
						</td>
						<td align="center" title="Eliminar Ajuste"><img onclick="deleteAjuste(this, <%=rsAJU("IDAJUSTE")%>)" src="images/compras/remove-16x16.png"></td>
					<% end if %>
				</td>
			</tr>
			<%
			rsAJU.movenext
		wend
	end if %>
	<% if (reg = 0) then %>
		<tr><td class="TDNOHAY" colspan="8"><% =GF_TRADUCIR("No se encontraron datos para mostrar") %></td></tr>
	<% end if %>
	</table>
<%
End Function
'----------------------------------------------------------------------------------
Function getRsAJU(idContrato)
	dim strSQL, conn, rs
	strSQL = "Select * from TBLOBRACTCAJUSTES WHERE IDCONTRATO=" & idContrato
	'Response.Write strSQL
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)								
	
	Set getRsAJU = rs
End Function
'----------------------------------------------------------------------------------
'*************************************************
'************** COMIENZO DE LA PAGINA ************
'*************************************************
dim idContrato,cdMoneda

idContrato = GF_PARAMETROS7("id", 0, 6)
cdMoneda   = GF_PARAMETROS7("cdMoneda", "", 6)

if (idContrato > 0) then Call mostrarAjustes(idContrato,cdMoneda)

%>

