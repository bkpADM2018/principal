<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosHKEY.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->
<%
'***********************************************************************************
'*******	                     COMIENZO DE LA PAGINA                      ********
'***********************************************************************************
dim newTotalPesos, newTotalDolares, saldoPesos, saldoDolares, ajustesPesos, ajustesDolares
dim firmante1Cd, firmante2Cd, firmante3Cd, firmante4Cd, firmante5Cd 
dim firmante1Ds, firmante2Ds, firmante3Ds, firmante4Ds, firmante5Ds
dim firmante1Tx, firmante2Tx, firmante3Tx, firmante4Tx, firmante5Tx
dim firmante1Sec, firmante2Sec, firmante3Sec, firmante4Sec, firmante5Sec
dim firmante1Rol, firmante2Rol, firmante3Rol, firmante4Rol, firmante5Rol
Dim CAB_ObraCD, CAB_ObraDS, CAB_ObraDivID, CAB_ObraDivDS, CAB_ObraImporte, CAB_FechaBudget, CAB_ObraMoneda, CAB_ObraFechaInicio, CAB_ObraFechaFin, CAB_ObraFechaAjustada,CAB_CdResponsable, CAB_DsResponsable,flagDireccion,rolUsuario
Dim ctz_AjusteTotal, ctz_det_importeFacturado, ctz_det_Importe, saldoImporte, ajustesImporte

flagDireccion = false
idAjuste = GF_Parametros7("idAjuste",0,6)
accion = GF_Parametros7("accion","",6)
errFirma = GF_PARAMETROS7("errFirma","",6)
if (errFirma <> "") then Call setError(errFirma)
if idAjuste <> 0 then 
    call cargarFirmas(idAjuste)
    rolUsuario = getRolFirma(session("Usuario"), SEC_SYS_COMPRAS)
    Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLCTZCABECERA_GET_AJUSTES_BY_IDAJUSTE", idAjuste)
    if (not rs.eof) then Call readCTZ(rs("IDCOTIZACION"))
end if
'------------------------------------------------------------------------------------------------------
Function cargarFirmas(pIdAjsute)
    Dim rsFirmas, connFirmas
	
	Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rsFirmas, "TBLCTZAJUSTESFIRMAS_GET_BY_IDAJUSTE", pIdAjsute)
	while not rsFirmas.eof
		if (firmante1Cd ="") then
			firmante1Cd = rsFirmas("CDUSUARIO")
			firmante1Ds = getUserDescription(rsFirmas("CDUSRROL"))
			firmante1Tx = armarTextoPlanoFirma(rsFirmas("HKEY"), rsFirmas("FECHAFIRMA"))
			firmante1Rol = CInt(rsFirmas("IDROL"))
			firmante1Sec = rsFirmas("SECUENCIA")
		elseif (firmante2Cd ="") then
			firmante2Cd = rsFirmas("CDUSUARIO")
			firmante2Ds = getUserDescription(rsFirmas("CDUSRROL"))
			firmante2Tx = armarTextoPlanoFirma(rsFirmas("HKEY"), rsFirmas("FECHAFIRMA"))
			firmante2Rol = CInt(rsFirmas("IDROL"))
			firmante2Sec = rsFirmas("SECUENCIA")
		elseif (firmante3Cd ="") then			
			firmante3Cd = rsFirmas("CDUSUARIO")
			firmante3Ds = getUserDescription(rsFirmas("CDUSRROL"))
			firmante3Tx = armarTextoPlanoFirma(rsFirmas("HKEY"), rsFirmas("FECHAFIRMA"))					
			firmante3Rol = CInt(rsFirmas("IDROL"))
			firmante3Sec = rsFirmas("SECUENCIA")
		elseif (firmante4Cd ="") then
			firmante4Cd = rsFirmas("CDUSUARIO")
			firmante4Ds = getUserDescription(rsFirmas("CDUSRROL"))
			firmante4Tx = armarTextoPlanoFirma(rsFirmas("HKEY"), rsFirmas("FECHAFIRMA"))
			firmante4Rol = CInt(rsFirmas("IDROL"))
			firmante4Sec = rsFirmas("SECUENCIA")
		elseif (firmante5Cd ="") then
			firmante5Cd = rsFirmas("CDUSUARIO")
			firmante5Ds = getUserDescription(rsFirmas("CDUSRROL"))
			firmante5Tx = armarTextoPlanoFirma(rsFirmas("HKEY"), rsFirmas("FECHAFIRMA"))
			firmante5Rol = CInt(rsFirmas("IDROL"))
			firmante5Sec = rsFirmas("SECUENCIA")
		end if				
		rsFirmas.MoveNext()
	wend	
			
End Function

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title><% =GF_TRADUCIR("Sistema de Compras - Ajuste " & ctz_docCode ) %></title>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<style type="text/css">
.titleStyle {
	font-weight: bold;
	font-size: 20px;
}
.labelStyle {
	font-weight: bold;	
}
.numberStyle {
	font-weight: bold;
	font-size: 14px;
}
.msgOK {
	font-weight: bold;
	font-size: 14px;
	color: #44CC66;
}
</style>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/hkey.js"></script>
<script type="text/javascript">
	var link = "comprasFirmarAjustePIC.asp?idAjuste=<%=idAjuste %>&secuencia=";
	var hkey0 = new Hkey('hk0', link + "<%=firmante1Sec%>", '<% =HKEY() %>', 'check_callback()');
	var hkey1 = new Hkey('hk1', link + "<%=firmante2Sec%>", '<% =HKEY() %>', 'check_callback()');
	var hkey2 = new Hkey('hk2', link + "<%=firmante3Sec%>", '<% =HKEY() %>', 'check_callback()');
	var hkey4 = new Hkey('hk4', link + "<%=firmante4Sec%>", '<% =HKEY() %>', 'check_callback()');
	var hkey5 = new Hkey('hk5', link + "<%=firmante5Sec%>", '<% =HKEY() %>', 'check_callback()');

	function check_callback(resp) {	
		if (resp != "<% =RESPUESTA_OK %>") document.getElementById("errFirma").value = resp;		
		document.getElementById("frmSel").submit();
	}
	function bodyOnLoad(){
		hkey0.start();
		hkey1.start();
		hkey2.start();		
		hkey4.start();
		hkey5.start();
	}
	
	function abrirCTC(id){
		window.open("comprasCTC.asp?idContrato=" + id, "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes",false);		
	}
	
	function abrirPCT(id){
		window.open("comprasFichaPedidoCotizacion.asp?idPedido=" + id + "&tab=1", "_blank", "location=no,scrollbars=yes,menubar=no,statusbar=no,height=500,width=500",false);
	}
	
</script>
</head>
<body onLoad="bodyOnLoad()">
<form method="post" id="frmSel" action="comprasAjustePICFirmas.asp?idAjuste=<%=idAjuste%>">
<div id="toolbar"></div><br>
<table class="reg_header" align="center" width="40%" border="0" >				
	<tr>
		<td colspan="6"><% call showErrors() %></td>
	</tr>

		<%if idAjuste <> 0 then 				    
			if not rs.eof then				
				%>
				<tr>
					<td align="right" class="numberStyle" colspan="6"><% =GF_TRADUCIR("Id " & ctz_docCode & ":") %>&nbsp;<% =rs("IDCOTIZACION") %></td>				
				</tr>				
				<%				
				Call initHeaderDB(rs("IDPEDIDO"))
				Call loadDatosObra(rs("IDOBRA"), CAB_ObraCD, CAB_ObraDS, CAB_ObraDivID, CAB_ObraDivDS, CAB_ObraImporte, CAB_FechaBudget, CAB_ObraMoneda, CAB_ObraFechaInicio, CAB_ObraFechaFin, CAB_ObraFechaAjustada,CAB_CdResponsable, CAB_DsResponsable)
				%>
				<tr>
					<td class="reg_header_nav" colspan="9"><% =GF_TRADUCIR("Datos del " & ctz_docCode) %></td>				
				</tr>
				<tr>
					<td class="reg_header_navdos"><% =GF_TRADUCIR("Ptda. Presup.") %></td>	
					<td colspan="5">
						<%=CAB_ObraCD & " - " & CAB_ObraDS %>
					</td>		
				</tr>
				<tr>
					<td class="reg_header_navdos"><% =GF_TRADUCIR("Pedido") %></td>	
					<td colspan="2">
						<% if(cdbl(rs("IDPEDIDO")) >  0)then %>
							<a><img id="imgPCT" src="images/compras/PCT-16X16.png" style="cursor:pointer" onclick="abrirPCT(<%=rs("IDPEDIDO") %>)" title="Abrir Pedido" ></a>&nbsp&nbsp;
						<% end if %>
						<%=pct_cdPedido%>
					</td>
					<td class="reg_header_navdos"><% =GF_TRADUCIR("Division") %></td>	
					<td colspan="2"><%=getDivisionDS(rs("IDDIVISION")) %></td>
				</tr>
				<tr>
					<td class="reg_header_navdos"><% =GF_TRADUCIR("Proveedor") %></td>
					<td colspan="2"><% =rs("IDPROVEEDOR") & " - " & getDescripcionProveedor(rs("IDPROVEEDOR")) %></td>
					<td ></td>
					<td width="3%"></td>		
					<td></td>		
				</tr>
				<% if (cdbl(rs("IDCONTRATO")) > 0) then %>
				<tr>
					<td class="reg_header_navdos"><% =GF_TRADUCIR("Contrato") %></td>
					<td colspan="2">
						<a><img id="imgCTC" src="images/compras/ctc-16x16.png" style="cursor:pointer" onclick="abrirCTC(<%=rs("IDCONTRATO")%>)" title="Abrir CTC" ></a>&nbsp&nbsp;
						<%= getCodigoCTC(rs("IDCONTRATO"))%>
					</td>
				</tr>
				<% end if %>
				<tr>
					<td colspan="6"><b>Observaciones</b></td>
				</tr>
				<tr>
					<td colspan="6"><% = rs("OBSERVACIONES") %></td>
				</tr>
				<tr>
					<td colspan="6"><hr></td>
				</tr>				
				<tr>
					<td class="reg_header_nav recuadroRound" colspan="6"><% =GF_TRADUCIR("Datos del Ajuste de Pedido") %></td>				
				</tr>
				<tr>
					<td colspan="2"></td>				
					<td align="center" width="30%"><b><u><% =GF_TRADUCIR("Importe") %></u></b></td>
					<td colspan="2" align="center" width="20%"><b><u><% =GF_TRADUCIR("Cantidad") %></u></b></td>
				</tr>
				<%				
				while not rs.eof
						ctz_cdMoneda = rs("CDMONEDA")						
						ctz_AjusteTotal = rs("IMPORTEPESOS_AJU")
						if (ctz_cdMoneda = MONEDA_DOLAR) then ctz_AjusteTotal = rs("IMPORTEDOLARES_AJU")
						ctz_AjusteCantidad = rs("CANTIDAD_AJU")
						ctz_det_importeFacturado = rs("IMPORTEPESOSFACTURADO_DET")
						if (ctz_cdMoneda = MONEDA_DOLAR) then ctz_det_importeFacturado = rs("IMPORTEDOLARESFACTURADO_DET")
						ctz_det_Facturado = rs("FACTURADO_DET")
						ctz_det_Importe = rs("IMPORTEPESOS_DET")
						if (ctz_cdMoneda = MONEDA_DOLAR) then ctz_det_Importe = rs("IMPORTEDOLARES_DET")
						ctz_det_ArticuloCantidad = rs("CANTIDAD")						
						saldoImporte = cdbl(ctz_det_Importe)-cdbl(ctz_det_importeFacturado) + cdbl(ctz_AjusteTotal)
						saldoCantidad = cdbl(ctz_det_ArticuloCantidad)-cdbl(ctz_det_Facturado) + cdbl(ctz_AjusteCantidad)
						ajustesImporte = ajustesImporte + CDbl(ctz_AjusteTotal)
						ctz_det_Estado = rs("APLICADO_DET")
						%>
							<tr>
								<td colspan="4"><b><%=rs("IDARTICULO_DET") & " - " & trim(rs("DSARTICULO_DET")) & "&nbsp;&nbsp;&nbsp;(" & rs("IDAREA_AJU") & "-" & rs("IDDET_AJU") & ")"%></b></td>	
							</tr>
							<tr>
								<td width="5%"></td>
								<td class="reg_header_nsav recuadrsoRound"><% =GF_TRADUCIR("Total del articulo") %></td>				
								<td align="right"><font size="+1"><b><%=getSimboloMoneda(ctz_cdMoneda) & " " & GF_EDIT_DECIMALS(ctz_det_importe,2)%></b></font></td>	
								<td align="right"><font size="+1"><b><%=ctz_det_ArticuloCantidad %></b></font></td>		
								<td><% =rs("ARTICULO_UNIDAD") %></td>
							</tr>
							<tr>
								<td></td>
								<td class="reg_header_nsav recuadsroRound"><% =GF_TRADUCIR("Facturado hasta el momento") %></td>				
								<td align="right"><font size="+1"><b><%=getSimboloMoneda(ctz_cdMoneda) & " " & GF_EDIT_DECIMALS(ctz_det_importeFacturado,2)%></b></font></td>	
								<td align="right"><font size="+1"><b><%=ctz_det_Facturado %></b></font></td>		
								<td><% =rs("ARTICULO_UNIDAD") %></td>
							</tr>
							<tr>
								<td></td>
								<td class="reg_header_nsav"><% =GF_TRADUCIR("Ajuste propuesto") %></td>				
								<td align="right"><font size="+1"><b><%=getSimboloMoneda(ctz_cdMoneda) & " " & GF_EDIT_DECIMALS(ctz_AjusteTotal,2)%></b></font></td>	
								<td align="right"><font size="+1"><b><%= ctz_AjusteCantidad %></b></font></td>		
								<td><% =rs("ARTICULO_UNIDAD") %></td>
							</tr>													
							<tr>								
								<td colspan="2" class="reg_headeAr_nav"></td>				
								<td align="right"><HR></td>	
								<td align="right"><HR></td>
								<td align="right" colspan="2"><HR></td>		
							</tr>
							<tr>
								<td></td>
								<td class="reg_header_nxav recuaxdroRound"><b><% =GF_TRADUCIR("Saldo") %></b></td>				
								<td align="right"><font size="+1"><b><%=getSimboloMoneda(ctz_cdMoneda) & " " & GF_EDIT_DECIMALS(saldoImporte,2)%></b></font></td>	
								<td align="right"><font size="+1"><b><%=saldoCantidad %></b></font></td>		
								<td><% =rs("ARTICULO_UNIDAD") %></td>
							</tr>													
							<tr>
								<td></td>
								<td colspan="5" class="reg_header_nav"><% =GF_TRADUCIR("Motivo") %></td>												
							</tr>							
							<tr>
								<td></td>
								<td colspan="5"><% =rs("MOTIVO") %></td>								
							</tr>						
						<%
					rs.MoveNext()
				wend	
			end if	
		end if %>


	<tr>
		<td COLSPAN="6" align="right"><bR></td>	
	</tr>
	<tr>
		<td COLSPAN="6" align="right"><bR></td>	
	</tr>
	<tr>
		<td colspan="6">	
            <%  flagYaFirmo = false		%>		
			<table align="center" width="80%" border="1" cellspacing=0 cellpadding=0>
		        <tr>
			        <td class="reg_header_nav" colspan="6"><% =GF_TRADUCIR("Firmas") %></td>
		        </tr>
		        <tr>
		            <td width="16%"></td>
		            <td width="16%"></td>
		            <td width="16%"></td>
		            <td width="16%"></td>
		            <td width="16%"></td>
		            <td ></td>
		        </tr>		
		        <tr>
			        <td align="center" colspan="2">
				        <%	if (firmante1Tx  <> "") then 
				                if (firmante1Cd = session("Usuario")) then flagYaFirmo = true
				        %>
					        <img src="images/firmas/<% =obtenerFirma(firmante1Cd) %>"><br>
					        <% =firmante1Tx %>
				        <%	else	
				                if ((session("Usuario") = firmante1Cd) or (rolUsuario = firmante1Rol)) then						
                                    flagYaFirmo = true				        
		                %>
							        <br><div id="hk0"></div><br>
					        <%	else	%>
							        <br><br><br>
					        <%	end if	
					        end if	%>
			        </td>
			        <td align="center" colspan="2">
				        <%	if (firmante2Tx <> "") then 
				                if (firmante2Cd = session("Usuario")) then flagYaFirmo = true
				        %>
					        <img src="images/firmas/<% =obtenerFirma(firmante2Cd) %>"><br>
					        <% =firmante2Tx %>
				        <%	else	
				                'response.Write session("Usuario") & "|" & firmante2Cd & "|" & rolUsuario & "|" & firmante2Rol & "|" & isNumeric(firmante2Cd) & "|" & flagBoss
						        if (((session("Usuario") = firmante2Cd) or (rolUsuario = firmante2Rol)) and (not flagYaFirmo)) then						
			                        flagYaFirmo = true
			            %>
							        <br><div id="hk1"></div><br>
					        <%	else	%>
							        <br><br><br>
					        <%	end if	
					        end if	%>
			        </td>
			        <td align="center" colspan="2">
				        <%	if (firmante3Tx <> "") then 
				                if (firmante3Cd = session("Usuario")) then flagYaFirmo = true
				        %>				
					        <img src="images/firmas/<% =obtenerFirma(firmante3Cd) %>"><br>
					        <% =firmante3Tx %>
				        <%	else	
				                'response.Write "USR Sess:" & session("Usuario") & "|CDUSUARIO:" & firmante3Cd & "|ROL:" & rolUsuario & "|FIRMA ROL:" & firmante3Rol & "|Numerico?:" & isNumeric(firmante2Cd) & "|Jefe:" & flagBoss
						        if (((session("Usuario") = firmante3Cd) or (rolUsuario = firmante3Rol))  and (not flagYaFirmo)) then						
				                    flagYaFirmo = true		
				        %>
							        <br><div id="hk2"></div><br>
					        <%	else	%>
							        <br><br><br>
					        <%	end if	
					        end if	%>
			        </td>
		        </tr>
		        <tr>
			        <td ALIGN="CENTER" colspan="2"><%=firmante1Ds%></td>
			        <td ALIGN="CENTER" colspan="2"><%=firmante2Ds%></td>
			        <td ALIGN="CENTER" colspan="2"><%=firmante3Ds%></td>										
		        </tr>
        <%      if ((firmante4Cd <> "") or (firmante5Cd <> "")) then %>		
		        <tr>
			        <td align="center" colspan="3">
				        <%	if (firmante4Tx <> "") then 
				                if (firmante4Cd = session("Usuario")) then flagYaFirmo = true
				        %>
					        <img src="images/firmas/<% =obtenerFirma(firmante4Cd) %>"><br>
					        <% =firmante4Tx %>
				        <%	else
						        if (((session("Usuario") = firmante4Cd) or (rolUsuario = firmante4Rol)) and (not flagYaFirmo)) then						
				                    flagYaFirmo = true		
				        %>
							        <br><div id="hk4"></div><br>
					        <%	else	%>
							        <br><br><br>
					        <%	end if	
					        end if	%>
			        </td>
			        <td align="center" colspan="3">
				        <%	if (firmante5Tx  <> "") then %>
					        <img src="images/firmas/<% =obtenerFirma(firmante5Cd) %>"><br>
					        <% =firmante5Tx %>
				        <%	else	
						        if (((session("Usuario") = firmante5Cd) or (rolUsuario = firmante5Rol)) and (not flagYaFirmo)) then						%>
							        <br><div id="hk5"></div><br>
					        <%	else	%>
							        <br><br><br>
					        <%	end if	
					        end if	%>
			        </td>
		        </tr>		
		        <tr>
			        <td ALIGN="CENTER" colspan="3"><%=firmante4Ds%></td>
			        <td ALIGN="CENTER" colspan="3"><%=firmante5Ds%></td>
		        </tr>
        <%      end if %>		
	        </table>
		</td>
	</tr>			
</table>
<input type="hidden" name="errFirma" id="errFirma">
</form>
</body>
</html>