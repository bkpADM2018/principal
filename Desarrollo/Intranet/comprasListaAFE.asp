<%
'----------------------------------------------------------------------------------
Function puedeModificarAFE(pIdObra, pIdDivision)
	Dim ret
	ret = false
	if (isAdmin(pIdDivision) or (isUser(pIdDivision))) then		
		if (getMmtoFinalizacionObra(idObra) > CLng(Left(session("MmtoDato"),8))) then
			ret = true
		end if		
	end if	
	puedeModificarAFE = ret
End Function
'----------------------------------------------------------------------------------
Function dibujarAFERaiz(rs)  %>	
	<td width="3%" style="cursor:pointer;" onclick="window.open('comprasAFEPrint.asp?idAFE=<%=rs("IDAFE")%>')"><img src="images/compras/AFE-16X16.png" alt="Afe16x16" style="cursor:pointer;"></td>
	<td colspan="2" width="47%" style="cursor:pointer;" onclick="window.open('comprasAFEPrint.asp?idAFE=<%=rs("IDAFE")%>')"><% =rs("CDAFE") %></td>
	<td style="cursor:pointer;text-align:center" onclick="window.open('comprasAFEPrint.asp?idAFE=<%=rs("IDAFE")%>')">
		<%	if (CInt(rs("IDAREA")) <> 0) then
				Response.write rs("IDAREA") & "-" & rs("IDDETALLE")
			end if	%>
	</td>
	<td align="right" onclick="window.open('comprasAFEPrint.asp?idAFE=<%=rs("IDAFE")%>')">
		<%	importe = rs("IMPORTEDOLARES")
			if (cdMoneda = MONEDA_PESO) then importe = rs("IMPORTEPESOS")
			response.write getSimboloMoneda(cdMoneda) & " " & GF_EDIT_DECIMALS(importe,2) 
		%>
		&nbsp;
	</td>	
	<td align="center"><% =getEditAFEIcon(rs("IDAFE")) %></td>
	<td align="center">
		<%
			'cdUsuario = ""
			dsUsuario = ""
			if ((IsNumeric(rs("CONFIRMADO")) ) or (rs("CONFIRMADO") = AFE_NO_CONFIRMADO)) then
				'cdUsuario = getUsuarioAFirmar(rs("IDAFE"))
				dsUsuario = getDSUsuarioAFirmar(rs("IDAFE"))						
				if (dsUsuario <> "") then
		%>
					<span style="cursor:pointer" onclick="alert('Se esta esperando la firma del usuario <% =dsUsuario %>')"><img style="cursor:pointer" src="images/compras/action_warning-16x16.png" title="<% =GF_TRADUCIR("AFE en firma") %>"></span>
		<%		end if %>
		<%	else	%>
				<span style="cursor:pointer" onclick="abrirAFEPrint(<% =rs("IDAFE") %>)"><img style="cursor:pointer;" src="images/print-16.png" title="<% =GF_TRADUCIR("Imprimir AFE") %>"></span>
		<%	end if	%>
	</td>
	<td align="center"><% =getRejectAFEIcon(rs("IDAFE")) %></td>
	<%
End Function
'----------------------------------------------------------------------------------
Function dibujarAFECompl(rs)
	%>
	<td width="3%" style="cursor:pointer;" onclick="window.open('comprasAFEPrint.asp?idAFE=<%=rs("IDAFE")%>')"></td>
	<td width="3%" style="cursor:pointer;" onclick="window.open('comprasAFEPrint.asp?idAFE=<%=rs("IDAFE")%>')"><img src="images/compras/AFE-16X16.png"></td>
	<td style="cursor:pointer;" onclick="window.open('comprasAFEPrint.asp?idAFE=<%=rs("IDAFE")%>')"><% =rs("CDAFE") %></td>
	<td style="cursor:pointer;text-align:center" onclick="window.open('comprasAFEPrint.asp?idAFE=<%=rs("IDAFE")%>')">
		<%	if (CInt(rs("IDAREA")) <> 0) then
				Response.write rs("IDAREA") & "-" & rs("IDDETALLE")
			end if			
		%>
	</td>
	<td align="right" onclick="window.open('comprasAFEPrint.asp?idAFE=<%=rs("IDAFE")%>')">
		<%	importe = rs("IMPORTEDOLARES")
			if (cdMoneda = MONEDA_PESO) then importe = rs("IMPORTEPESOS")
			response.write getSimboloMoneda(cdMoneda) & " " & GF_EDIT_DECIMALS(importe,2) 
		%>
		&nbsp;
	</td>
	<td align="center"><% =getEditAFEIcon(rs("IDAFE")) %></td>
	<td align="center">
		<%
			'cdUsuario = ""
			dsUsuario = ""
			if ((IsNumeric(rs("CONFIRMADO"))) or (rs("CONFIRMADO") = AFE_NO_CONFIRMADO)) then
				'cdUsuario = getUsuarioAFirmar(rs("IDAFE"))
				dsUsuario = getDSUsuarioAFirmar(rs("IDAFE"))
				if (dsUsuario <> "") then
		%>
					<span style="cursor:pointer" onclick="alert('Se esta esperando la firma del usuario <% =dsUsuario %>')"><img style="cursor:pointer" src="images/compras/action_warning-16x16.png" title="<% =GF_TRADUCIR("AFE en firma") %>"></span>
		<%		end if %>
		<%	else	%>
				<span style="cursor:pointer" onclick="abrirAFEPrint(<% =rs("IDAFE") %>)"><img style="cursor:pointer" src="images/print-16.png" title="<% =GF_TRADUCIR("Imprimir AFE") %>"></span>
		<%	end if	%>
	</td>
	<td align="center"><% =getRejectAFEIcon(rs("IDAFE")) %></td>
	<%
End Function
'----------------------------------------------------------------------------------
Function drawTable(rsAFE, title, idPedido, mustBeEqual)
	Dim hayAFE, totalAFEs, importeAFE
	totalAFEs = 0
	hayAFE=false
%>
	<table class="datagrid" width="95%" align="center">
        <thead>
            <tr align="center" height="24px">
                <th colspan="8" class="reg_Header_Info"><% =title %></th>
            </tr>
            <tr>
          	    <th colspan="3" align="center">AFE</th>
                <th width="20%" align="center"> PARTIDA </th>
                <th width="20%" align="center"> IMPORTE </th>
                <th width="5%" align="center">-</th>
                <th width="5%" align="center">-</th>
                <th width="5%" align="center">-</th>
            </tr>
        </thead>
        <tbody>
	<%	if (not rsAFE.eof) then			
				while (not rsAFE.eof) 			
					importeAFE = cdbl(rsAFE("IMPORTEDOLARES"))
					if (cdMoneda = MONEDA_PESO) then importeAFE = cdbl(rsAFE("IMPORTEPESOS"))
					if rsAFE("Confirmado") = AFE_ANULADO then importeAFE = importeAFE * -1
					totalAFEs = totalAFEs + cdbl(importeAFE)
					'Se compara el ID de pedido pasado con el ID de pedido registrado en el AFE
					'se agrega una condición que permite controlar si la comparación es por igualdad o no.
					'Esto es así debido a que para las obras se muestran todos los AFEs de cualquier pedido, pero para los pedidos solo los propios.					
					if (((CLng(rsAFE("IDPEDIDO")) = idPedido) and mustBeEqual) or _
						(((CLng(rsAFE("IDPEDIDO")) <> idPedido) or (idPedido = 0)) and not mustBeEqual)) then
						hayAFE= true
					%>
					<tr <%If (rsAFE("Confirmado") = AFE_ANULADO) then %> class="reg_header_rejected" <% end if %> >
						<% Call dibujarAFERaiz(rsAFE) %>
					</tr>
					<%	cont = cont + 1
						Set rsCompl = listaAFESComplementarios(rsAFE("IDAFE"))
							
						while (not rsCompl.eof)	
							importeAFE = cdbl(rsCompl("IMPORTEDOLARES"))
							if (cdMoneda = MONEDA_PESO) then importeAFE = cdbl(rsCompl("IMPORTEPESOS"))
							if rsCompl("Confirmado") = AFE_ANULADO then importeAFE = importeAFE * -1
							totalAFEs = totalAFEs + cdbl(importeAFE)						
							'totalAFEs = totalAFEs + cdbl(rsCompl("IMPORTEDOLARES"))
						%>	
							<tr <%If (rsCompl("Confirmado") = AFE_ANULADO) then %> class="reg_header_rejected" <% end if %>>
								<% Call dibujarAFECompl(rsCompl) %>
							</tr>
						<%	rsCompl.MoveNext()
							cont = cont + 1
						wend
					end if
					rsAFE.MoveNext()
				wend
			end if%>
            </tbody>
		<%	if (not hayAFE) then%>
				<tr><td colspan="8"><% =GF_TRADUCIR("No se encontraron AFEs registrados.") %></td></tr>
			<% else %>
            <tfoot>
				<tr>
					<td colspan="4" align="right"><% =GF_TRADUCIR("Total") %></td>
					<td width="20%"  align="right"><% = getSimboloMoneda(cdMoneda) & " " & GF_EDIT_DECIMALS(totalAFEs,2)  %>&nbsp;&nbsp;</td>
					<td colspan="3" align="right">&nbsp;</td>
				</tr>
            </tfoot>
			<% 
			end if%>
	</table>
<%
End Function
'***************************************************
'******   COMIENZO DE LA PAGINA
'***************************************************
'	Esta pagina se llama desde comprasPopUpAFE y comprasTableroObra, datos como la moneda se definen alli!
'***************************************************
Dim colorP, myColor1, myColor2, cont, rsAFE, confirmar, idAFE, rsCompl, myIdDivision, obraPedido

if (idObra = "") then idObra = GF_PARAMETROS7("idObra",0,6)
idPedido = GF_PARAMETROS7("idPedido",0,6)
idAFE = GF_PARAMETROS7("idAFE",0,6)
confirmar = GF_PARAMETROS7("confirmar","",6)
idContrato = GF_PARAMETROS7("idContrato",0,6)

if (idContrato > 0) then idPedido = getPedidoCTC(idContrato)'Si indico un contrato, se determina su pedido

if (idPedido > 0) then	'Si indico un pedido, veo si puedo acceder al pedido
	Call comprasControlAccesoCM(RES_CC)
	Call initHeader(idPedido)	
	puedeModificar = checkControlPCT()	
elseif (idObra > 0) then 'Si mando una obra, chequeo permisos de obra
	Call comprasControlAccesoCM(RES_OBR)
	call getDivisionObraFull(idObra, myIdDivision, "")
	puedeModificar = puedeModificarAFE(idObra, myIdDivision)
else	'Si no hay ni obra ni pedido, todo depende de los permisos de AFE
	Call comprasControlAccesoCM(RES_AFE)	
	myIdDivision = 0
end if

Call GP_ConfigurarMomentos


	
%>
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />

<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>	
<script type="text/javascript">
	function editAFE(idAfe) {
		window.open('comprasAFE.asp?idAFE=' + idAfe);
	}
	function anularAFE(idAFE){
		if (confirm("Esta seguro que desea eliminar este AFE?")) {
			window.scrollTo(0,0);
			window.resizeTo(600, 600);
			var puw = new winPopUp('popUpAnularAFE','comprasAFEAnulacion.asp?idAFE=' + idAFE, 500, 350,'Anulación del AFE', 'reload()');
		}
	}
	function reload() {
		window.resizeTo(500, 400);
		document.location.reload();
	}
	function abrirAFEPrint(id){
		window.open("comprasAFEPrint.asp?idAFE=" + id, "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes,height=500,width=700",false);
	}
</script>
<%	
	if (idObra <> 0) then 		
		Set rsAFE = readAFEObra(idObra, AFE_RAIZ)
		Call drawTable(rsAFE, GF_TRADUCIR("Obra") & ": " & getDescripcionObra(idObra), idPedido, false)
	end if
	'Se listan los AFE de un pedido
	if (idPedido <> 0) then 		
%><hr><%				
			Set rsAFE = readAFEPedido(idPedido, AFE_RAIZ)
			Call drawTable(rsAFE, GF_TRADUCIR("Pedido") & " " & pct_cdPedido, idPedido, true)
	end if
%>
<table class="datagrid" width="95%" align="center">
	<% 
	if (puedeModificar) then %>
    <tfoot>
	    <tr>
  	        <td valign="top" align="left">
                <span style="cursor:pointer;" onclick="window.open('comprasAFE.asp?idPedido=<% =idPedido %>&idObra=<%=idObra%>');" >
                    <img src="images/compras/add-16x16.png" alt="Afe16x16"/ >
                    <% =GF_TRADUCIR("Agregar AFE") %>
                </span>
            </td>
        </tr>
    </tfoot>
	<%	end if %>										
</table>