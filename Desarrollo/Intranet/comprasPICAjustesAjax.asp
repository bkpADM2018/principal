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
Function mostrarAjustes(idPIC)
	dim rsAJU, reg, minuta, auxminuta, linkFac, descArticulo, cbte
	dim pAbrev, rsArt, conn, cdInterno, totalpesos, totaldolares, myMoneda
	Set rsAJU = getRsAJU(idPIC)
%>	
	<table align="center" width="90%" class="reg_Header">
		<tr>
			<td align="center" colspan="8"><b><% =GF_TRADUCIR("Ajustes realizados al PIC") %></b></td>
		</tr>
		<tr class="reg_Header_nav">
			<td align="center"><% =GF_TRADUCIR("Id Aju") %></td>
			<td align="center"><% =GF_TRADUCIR("Articulo") %></td>
			<td align="center"><% =GF_TRADUCIR("Ptda. Presup.") %></td>
<%          myMoneda = MONEDA_PESO
            if (not rsAJU.eof) then myMoneda = rsAJU("CDMONEDA") 
            if (myMoneda = MONEDA_PESO) then                                      %>			
			<td align="center"><% =GF_TRADUCIR("Ajuste $") %></td>
<%          else                                                                           %>			
			<td align="center"><% =GF_TRADUCIR("Ajuste U$S") %></td>
<%          end if                                                                         %>			
			<td align="center"><% =GF_TRADUCIR("Cantidad") %></td>
			<td align="center"><% =GF_TRADUCIR("Observaciones") %></td>
			<td width="5%" align="center"><% =GF_TRADUCIR("Apl.") %></td>
			<td width="2%" align="center"><% =GF_TRADUCIR("Elim.") %></td>
		</tr>

<%
	reg=0
	if (not rsAJU.eof) then
		while (not rsAJU.eof)
			reg = reg + 1
			%>							
			<tr class="reg_Header_navdos" onMouseOver="javascript:lightOn(this, '')" onMouseOut="javascript:lightOff(this, '')">
				<td align="center"><%=rsAJU("IDAJUSTE")%></td>
				<td align="LEFT"><%=rsAJU("IDARTICULO")%> - <%=rsAJU("DSARTICULO")%></td>
				<td align="CENTER"><%=rsAJU("IDAREA")%> - <%=rsAJU("IDDETALLE")%></td>
<%          if (rsAJU("CDMONEDA") = MONEDA_PESO) then                                      %>							
				<td align="RIGHT"><% =getSimboloMoneda(MONEDA_PESO) & "&nbsp;" & GF_EDIT_DECIMALS(cDbl(rsAJU("IMPORTEPESOS")),2) %></td>
<%          else                                                                           %>							
				<td align="RIGHT"><% =getSimboloMoneda(MONEDA_DOLAR) & "&nbsp;" & GF_EDIT_DECIMALS(cDbl(rsAJU("IMPORTEDOLARES")),2) %></td>
<%          end if                                                                         %>							
				<td align="RIGHT"><% =rsAJU("CANTIDAD") & " " & rsAJU("ABREVIATURA") %></td>				
				<td align="center" TITLE="<%=rsAJU("OBSERVACIONES")%>">
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
						<td align="center" title="Eliminar Ajuste"><img onclick="deleteAjuste(this, <%=rsAJU("IDAJUSTE")%>,<%=idPIC%>)" src="images/compras/remove-16x16.png"></td>
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
Function getRsAJU(idPIC)
	dim strSQL, conn, rs
	strSQL = "Select IDAJUSTE, CA.*, DSARTICULO, ABREVIATURA, PIC.CDMONEDA from TBLCTZAJUSTES CA INNER JOIN TBLARTICULOS ART ON CA.IDARTICULO=ART.IDARTICULO INNER JOIN TBLUNIDADES UNI on ART.IDUNIDAD=UNI.IDUNIDAD inner join TBLCTZCABECERA PIC on PIC.IDCOTIZACION=CA.IDCOTIZACION WHERE CA.IDCOTIZACION=" & idPIC
	'Response.Write strSQL
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set getRsAJU = rs
End Function
'----------------------------------------------------------------------------------
'*************************************************
'************** COMIENZO DE LA PAGINA ************
'*************************************************
dim idPIC

idPIC = GF_PARAMETROS7("id", 0, 6)

if (idPIC > 0) then Call mostrarAjustes(idPIC)

%>

