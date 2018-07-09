<%
%>
<div style="height:450px;">
<body>
<br>
<table align="center" width="90%" border=0 cellpadding=1 cellspacing=1>
	<tr>
		<td colspan="4" class="TDNOHAY"><b><% =GF_TRADUCIR("CONTRATOS RELACIONADOS") %> </b></td>			
	</tr>
<%
dim lastProveedor, pUserDS, canConfirm, hayContrato
hayContrato = false

strSQL="Select OC.IDCONTRATO, OC.CDCONTRATO, OC.CDRESPONSABLE, OC.CDSUPERVISOR, OC.IDPEDIDO, OC.IDOBRA, OC.IDPROVEEDOR, OC.ESTADO, "
strSQL=strSQL & " VE.NROEMP AS IDEMPRESA, VE.NOMEMP AS DSEMPRESA, OC.CDUSERCONF "
strSQL=strSQL & " from TBLOBRACONTRATOS OC INNER JOIN [Database].[dbo].MET001A VE on VE.NROEMP =  OC.IDPROVEEDOR "
strSQL=strSQL & " WHERE OC.IDPEDIDO=" & pct_idPedido & " ORDER BY OC.IDPROVEEDOR"
Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
if (not rs.eof) then
	while (not rs.eof)
		if (lastProveedor <> rs("IDPROVEEDOR")) then
			if (lastProveedor <> "") then Response.Write "<tr><td colspan='4'><hr></td></tr>"
			%>
			<tr>
				<td colspan="4"><b><% =rs("IDEMPRESA") & "-" & rs("DSEMPRESA") %> </b></td>			
			</tr>
			<%
			lastProveedor = rs("IDPROVEEDOR")
		else
			Response.Write "<tr><td colspan='3'>&nbsp;</td><td colspan=1><hr></td></tr>"	
		end if
		hayContrato = false
		if (CInt(rs("ESTADO")) <> ESTADO_CTC_CANCELADO) then hayContrato = true
		'Se asume que si la persona puede confirmar un contrato, tambien puede anularlo.
		canConfirm = canConfirmCTC(session("usuario"), rs("IDCONTRATO"))
		%>
			<tr>
				<td rowspan="3" width="5%">&nbsp;</td>
				<td rowspan="3" width="10%" valign="middle">
					<span style="cursor:pointer;" onClick="abrirContrato(<% =rs("IDCONTRATO") %>, <% =rs("ESTADO") %>, <% =Lcase(canConfirm) %>)">
						<img src="images/compras/CTC-48x48.png">
					</span>
				</td>
				<td>
					<% if (hayContrato) then %>
					<span style="cursor:pointer;color:blue;" onClick="abrirContrato(<% =rs("IDCONTRATO") %>, <% =rs("ESTADO") %>, <% =Lcase(canConfirm) %>)">
						Contrato: <% =rs("CDCONTRATO") %>
					</span>
					<% else %>
					<span class="TDERROR">Contrato: ANULADO</span>			
					<% end if %>
				</td>				
				<%	if (hayContrato) then		%>
					<td align="center" style="cursor:pointer"> 
					<%  if (canConfirm) then
							imagen="CTC_Confirm-16X16.png"
							textoAccion = "Confirmar" 
						else
							imagen="edit-16x16.png"
							textoAccion = "Editar" 
						end if
					%>
						
						<img title="<%=GF_TRADUCIR(textoAccion)%>" src="images/compras/<% =imagen %>" onclick="abrirContrato(<% =rs("IDCONTRATO") %>, <% =rs("ESTADO") %>, <% =Lcase(canConfirm) %>)"></td>
				<%	else	%>
					<td align='center'>&nbsp;</td>
				<%	end if	%>				
			</tr>
			<tr>				
				<td>
				<% 
				pUserDS = getUserDescription(rs("CDRESPONSABLE"))				
				response.write GF_TRADUCIR("Responsable: ") & rs("CDRESPONSABLE") & " - " & pUserDS								
				if (rs("ESTADO") > ESTADO_CTC_PENDIENTE) then
					pUserDS = getUserDescription(rs("CDUSERCONF"))  
					if (rs("ESTADO") = ESTADO_CTC_CANCELADO) then 
						textoAccion = "Canceló: "
					else
						textoAccion = "Confirmó: "
					end if
					response.write "<br><br>" & GF_TRADUCIR(textoAccion) & rs("CDUSERCONF") & " - " & pUserDS
				end if
				%>									
				</td>			
				<% if (canConfirm and hayContrato) then %>
					<td align="center" width="32px" style="cursor:pointer"> <img title="<%=GF_TRADUCIR("Anular Contrato")%>" id="ID_<%=rs("IDCONTRATO")%>" src="images/compras/CTZ_cancel-16x16.png" onclick="anularContrato(<% =rs("IDCONTRATO") %>, this)"></td>
				<%	else	%>
					<td align='center' width="32px">&nbsp;</td>
				<%	end if  %>								
			</tr>
		<%
		rs.movenext
	wend
	Response.Write "<tr><td colspan='4'><hr></td></tr>"
else
	%>
	<tr>
		<td colspan="3" align="center"><b><% =GF_TRADUCIR("No se encontraron Contratos asociados al pedido.") %> </b></td>			
	</tr>
	<tr><td colspan=3><hr></td></tr>
	<%
end if
if pct_idestado >= ESTADO_PCT_ADJUDICADO then
	
	%>
	<tr>
		<td colspan="3" onClick="cargarContrato(<% =pct_idPedido %>);">
		<%	if (not hayContrato) then %>
				<span style="cursor:pointer;color:blue">
					<img align="absMiddle" src="images/compras/CTC_new-16x16.png">&nbsp;<% =GF_TRADUCIR("Cargar Contrato")%>
				</span>
		<%	end if %>
		</td>
	</tr>		
	<% 
end if
%>	
</table>
<br>
<table align="center" width="90%" border=0 cellpadding=1 cellspacing=1>
	<tr>
		<td colspan="4" class="TDNOHAY"><b><% =GF_TRADUCIR("PEDIDOS INTERNOS DE COMPRAS RELACIONADOS") %> </b></td>			
	</tr>
<%
strSQL="Select CC.*, VE.NROEMP AS IDEMPRESA, VE.NOMEMP AS DSEMPRESA, RP.IDPIC from " 
strSQL= strSQL & " (Select C.IDCOTIZACION, C.CDUSUARIO, C.IDPROVEEDOR, C.ESTADO, C.IDPEDIDO, C.MOMENTO, SUM(FACTURADO) FACTURADO from TBLCTZCABECERA C inner join tblctzdetalle D on C.IDCOTIZACION=D.IDCOTIZACION WHERE C.IDPEDIDO=" & pct_idPedido & " AND C.IDCONTRATO = 0 group by C.IDCOTIZACION, C.CDUSUARIO, C.IDPROVEEDOR, C.ESTADO, C.IDPEDIDO, C.MOMENTO) CC"
strSQL= strSQL & " INNER JOIN  [Database].[dbo].MET001A VE on VE.NROEMP = CC.IDPROVEEDOR " 
strSQL= strSQL & " LEFT JOIN (Select Distinct IDPIC from TBLREMPIC) RP on RP.IDPIC=CC.IDCOTIZACION"
srtSQL= strSQL & " ORDER BY VE.NROEMP"
Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
imagen="PIC-48x48.png"
if not rs.eof then
	lastProveedor = 0
	while not rs.eof
		if lastProveedor <> CLng(rs("IDEMPRESA")) then
			if lastProveedor <> "" then Response.Write "<tr><td colspan='4'><hr></td></tr>"
			%>
			<tr>
				<td colspan="4"><b><% =rs("IDEMPRESA") & "-" & rs("DSEMPRESA") %> </b></td>			
			</tr>
			<%
			lastProveedor = CLng(rs("IDEMPRESA"))
		else
			Response.Write "<tr><td colspan=2>&nbsp;</td><td colspan='4'><hr></td></tr>"	
		end if	
		%>
			<tr>
				<td rowspan="3" width="5%">&nbsp;</td>
				<td rowspan="3" width="10%" valign=middle><img src="images/compras/<% =imagen %>"></td>
				<td><a href="comprasPICPrint.asp?idCotizacionElegida=<% =rs("IDCOTIZACION") %>" target="_blank">Nro. Ref.: <% =GF_TRADUCIR(rs("IDCOTIZACION")) %></a></td>
				<%							
					'Sin no esta anulada y no esta pagada, permite modificar.					
					if ((CDbl(rs("FACTURADO")) = 0) and (CStr(rs("ESTADO")) <> CTZ_ANULADA)) then
				%>
					<td align="center" style="cursor:pointer"> <img title="<%=GF_TRADUCIR("Editar Pedido Interno")%>" id="ID_<%=rs("IDCOTIZACION")%>" src="images\compras\edit-16x16.png" onclick='editarCTZ(<%=rs("IDCOTIZACION")%>, this)'></td>
				<%	else	%>
					<td align='center'>&nbsp;</td>
				<%	end if	%>
			</tr>	
			<tr>
				<% 
				pUserDS = getUserDescription(rs("CDUSUARIO")) 
				response.write "<td>" & GF_TRADUCIR("Cargó: ") & rs("CDUSUARIO") & " - " & pUserDS & "</td>"
				'Si no está anulada, ni facturada y no se recibió mercadería, permite anular.
					if ((CStr(rs("ESTADO")) <> CTZ_ANULADA) and (isNull(rs("IDPIC"))) and (CDbl(rs("FACTURADO")) = 0)) then %>
					<td align="center" width="32px" style="cursor:pointer"> <img title="<%=GF_TRADUCIR("Anular Pedido Interno")%>" id="ID_<%=rs("IDCOTIZACION")%>" src="images\compras\CTZ_cancel-16x16.png" onclick='anularCTZ(<%=rs("IDCOTIZACION")%>, <%=rs("IDPEDIDO")%>, this)'></td>
				<%	else	%>
					<td align='center' width="32px">&nbsp;</td>
				<%	end if  %>				
			</tr>
			<tr>
				<% 
				if rs("ESTADO") = CTZ_ANULADA then 
					txt = "<font color='red'><b>" & GF_TRADUCIR("ANULADO") & "</b></font>"
				else
					txt = GF_TRADUCIR("Momento de carga: ") & GF_FN2DTE(rs("MOMENTO")) 
				end if	
				%>
				<td><% =txt%></td>
				<% 'Si no está anulada y se recibió mercadería, permite ver los remitos que tiene asocido.
					if (CStr(rs("ESTADO")) <> CTZ_ANULADA) then %>
					<td align="center" style="cursor:pointer"> <img title="<%=GF_TRADUCIR("Ver Cumplimiento")%>" id="ID_<%=rs("IDCOTIZACION")%>" src="images\compras\PIC-16x16.png" onclick='abrirREMPIC(<%=rs("IDCOTIZACION")%>)'></td>
				<%	else	%>
					<td align='center'>&nbsp;</td>
				<%	end if  %>
			</tr>
		<%
		rs.movenext
	wend 
	Response.Write "<tr><td colspan='4'><hr></td></tr>"	
else
	%>
	<tr>
		<td colspan="3" align="center"><b><% =GF_TRADUCIR("No se encontran Pedidos Internos de Compras asociados al pedido.") %> </b></td>			
	</tr>
	<tr><td colspan=3><hr></td></tr>
	<%
end if
if pct_idestado >= ESTADO_PCT_ADJUDICADO then
	%>
	<tr>
		<td colspan="4">
			<a href="comprasPIC.asp?idPedido=<% =pct_idPedido %>" target="_blank"><img align="absMiddle" src="images/compras/PIC_new-16x16.png">&nbsp;<% =GF_TRADUCIR("Cargar Pedido Interno")%></a>
		</td>
	</tr>		
	<% 
end if
%>	
</table>
<br>
<!--------------------------------------------------------------------------------------------------------------------->
<!--------------------------------------------------POLIZAS  DE CAUCION ----------------------------------------------->
<!--------------------------------------------------------------------------------------------------------------------->
<table align="center" width="90%" border=0 cellpadding=1 cellspacing=1>
	<tr>
		<td colspan="4" class="TDNOHAY"><b><% =GF_TRADUCIR("POLIZAS DE CAUCION RELACIONADAS") %> </b></td>			
	</tr>
<%
Response.Write "<tr></td><td colspan='4'><hr></td></tr>"	
dim lastTomador

strSQL =		  "		SELECT POL.*, EMP.NROEMP AS IDEMPRESA, EMP.NOMEMP AS DSEMPRESA 	"
strSQL = strSQL & "	 	FROM ( SELECT *								"
strSQL = strSQL & "			   FROM TBLPOLIZASCAUCION		"
strSQL = strSQL & "			   WHERE IDPEDIDO = " & pct_idPedido
strSQL = strSQL & "			 ) AS POL								"
strSQL = strSQL & "		INNER JOIN [Database].[dbo].MET001A EMP			" 
strSQL = strSQL & "			ON EMP.NROEMP = POL.TOMADOR			"
strSQL = strSQL & "		ORDER BY EMP.NROEMP ASC "
Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
if (not rs.eof) then
	while (not rs.eof)
		if (lastTomador <> Cdbl(rs("IDEMPRESA"))) then
			if (lastTomador <> "") then Response.Write "<tr><td colspan='4'><hr></td></tr>"
			%>
			<tr>
				<td colspan="4"><b><% =rs("IDEMPRESA") & "-" & rs("DSEMPRESA") %> </b></td>			
			</tr>
			<%
			lastTomador = Cdbl(rs("IDEMPRESA"))
		else
			Response.Write "<tr><td colspan=2>&nbsp;</td><td colspan='4'><hr></td></tr>"	
		end if		
		%>
			<tr>
				<td rowspan="3" width="5%">&nbsp;</td>
				<td rowspan="3" width="10%" valign="middle">
					<img src="images/compras/PDC-48x48.png">
				</td>
				<td>
					<span style="cursor:pointer;color:blue;">
						Nro Poliza: <% =rs("NROPOLIZA") %>
					</span>					
				</td>								
				<td align='center'>&nbsp;</td>
			</tr>
			<tr>				
				<td>
				<% pUserDS = getUserDescription(rs("CDUSUARIO"))				
				   response.write GF_TRADUCIR("Responsable: ") & rs("CDUSUARIO") & " - " & pUserDS	%>									
				</td>
				<td align='center' width="32px">&nbsp;</td>				
			</tr>
			<tr>				
				<td>
				<% pUserDS = getUserDescription(rs("CDUSUARIO"))				
				   response.write GF_TRADUCIR("Momento de carga: ") & GF_FN2DTE(rs("MMTO"))	%>
				</td>
				<td align='center' width="32px">&nbsp;</td>
			</tr>			
		<%
		rs.movenext
	wend
	
else
	%>
	<tr>
		<td colspan="3" align="center"><b><% =GF_TRADUCIR("No se encontraron Polizas asociadas al pedido.") %> </b></td>
	</tr>
	<tr><td colspan=3><hr></td></tr>
	<%
end if
%>	
</table>
<!--------------------------------------------------------------------------------------------------------------------->
<!--------------------------------------------------POLIZAS  DE CAUCION ----------------------------------------------->
<!--------------------------------------------------------------------------------------------------------------------->

</body>	
</div>
