<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosMantenimiento.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<%
Call initAccessInfo(RES_OT_SM)
Dim rsList, idOrder
idOrder = GF_PARAMETROS7("idOrder", 0,6)
call executeProcedureDb(DBSITE_SQL_INTRA, rsList, "TBLSMOTEXECUTIONS_GET_BY_IDOT", idOrder)
if rsList.eof then
	Response.Write "No se encontraron ocurrencias"
else
	%>
		<table class="datagrid2" width="95%" align="center">
			<thead>
				<tr>
					<th class="thiconac" align="center"><%=GF_Traducir("Nro OT")%></th>
					<th class="thiconac" align="center"><%=GF_Traducir("Fecha Progr.")%></th>
					<th class="thiconac" align="center"><%=GF_Traducir("Fecha Inicio")%></th>
					<th class="thiconac" align="center"><%=GF_Traducir("Fecha Fin")%></th>
					<th class="thiconac" align="center"><%=GF_Traducir("Solicitante")%></th>
					<th class="thiconac" align="center"><%=GF_Traducir("Responsable")%></th>
					<th class="thiconac" align="center"><%=GF_Traducir("Estado")%></th>
					<th class="thiconac" align="center"><%=GF_Traducir("Accion")%></th>
				</tr>
			</thead>
			<tbody> 	
			<%			
			while not rsList.eof 
				if cint(rsList("IDOTGENERATED")) = 0 then
				%>
					<tr><td colspan="8" align="center"><%=GF_Traducir("Pr�xima ocurrencia estimada para el d�a: ") & GF_FN2DTE(rsList("NEXTEXECUTION"))%></td></tr>
				<%
				else
				%>
				<tr>
					<td class="thicon" align="center" nowrap><%=rsList("NROORDER")%></td> 
					<td class="thicon" align="center"><%=GF_FN2DTE(rsList("SCHEDULEDDATE"))%></td> 
					<td class="thicon" align="center"><%=GF_FN2DTE(rsList("STARTDATE"))%></td> 
					<td class="thicon" align="center"><%=GF_FN2DTE(rsList("FINISHEDDATE"))%></td> 
					<td class="thicon" align="center" title="<%=getUserDescription(rsList("CDAPPLICANT"))%>"><%=rsList("CDAPPLICANT")%></td> 
					<td class="thicon" align="center"><%=rsList("NOMEMP")%></td> 
					<td class="thicon" align="center" TITLE="<%=rsList("OBSERVATIONS")%>"><%=getDsState(rsList("CDSTATE"))%></td> 
					<td class="thicon" align="center">
						<img src="images/ot-16.png" style="cursor: pointer" title="Ir a la Orden" onClick="irA('mantenimientoAdministrarOTs.asp?txtNroOrder=<% =rsList("NROORDER") %>')">						
						<img src="images/print-16.png" style="cursor: pointer" title="Imprimir" onClick="imprimirOT('<% =rsList("IDORDER") %>')">
					</td>
				</tr>
				<%	
				end if
				rsList.movenext
			wend	
			%>
			</tbody>
		</table>	
		<%
end if	
'Response.Write idOrder

%>