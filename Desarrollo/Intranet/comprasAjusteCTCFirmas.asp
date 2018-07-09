<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosHKEY.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->

<%
'------------------------------------------------------------------------------------------------------------
'Dibuja dentro de la cabecera del contrato una linea con los datos de la obra
Function drawRowObra()
	Dim auxDsAreaObra, auxDsDetalleObra
 %>
	<tr>
		<td class="reg_Header_nav"><% =GF_TRADUCIR("Obra") %></td>
		<% if CTC_idObra > 0 then %>
		<td class="reg_Header_navdos" colspan="3" >
			<table class="reg_Header_navdos" width="100%" cellspacing="0" cellpadding="0" border="0">
				<tr>
					<td width="6%"></td>
					<td width="6%"></td>
					<td width="88%"></td>
				</tr>
				<tr>
					<td colspan="3" style="cursor:pointer;" onClick="abrirObra(<% =CTC_idObra %>);" >&nbsp;
						<b><% =CTC_obraCD & " - " & CTC_obraDS %></b>
					</td>
				</tr>
				<% if CTC_areaObra > 0 then
					Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rsAFE, "TBLBUDGETOBRAS_GET_BY_PARAMETERS", CTC_idObra & "||" & CTC_areaObra & "||0||1||")
					auxDsAreaObra = ""
					if not rsAFE.EoF then auxDsAreaObra = rsAFE("DSBUDGET")	%>
				<tr>
					<td></td>
					<td colspan="2">&nbsp;<img src="images\zwe0.gif"><b><% =CTC_areaObra & "-" & auxDsAreaObra%></b></td>
				</tr>
				<%end if%>
				<% if CTC_detalleObra > 0 then 
					Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rsAFE, "TBLBUDGETOBRAS_GET_BY_PARAMETERS", CTC_idObra & "||" & CTC_areaObra & "||" & CTC_detalleObra & "||1||")
					auxDsDetalleObra = ""
					if not rsAFE.EoF then auxDsDetalleObra = rsAFE("DSBUDGET") %>
				<tr>
					<td></td>
					<td></td>
					<td>&nbsp;<img src="images\zwe0.gif"><b><% =CTC_detalleObra & "-" & auxDsDetalleObra%></b></td>
				</tr>
				<%end if%>
			</table>			
		</td>				
		<% else %>
		<td class="reg_Header_navdos" colspan="3"><%=GF_TRADUCIR("Sin Obra")%></td>
		<% end if %>		
	</tr><%
End function
'------------------------------------------------------------------------------------------------------------
Function mostrarAjustes(pIdAjuste)
	dim rsAJU, rsCtc
	Set rsAJU = getRsAJU(pIdAjuste)
	
	if (not rsAJU.eof) then
	    Call loadContrato(rsAJU("idContrato"), "")
	    simboloMoneda = getSimboloMoneda(CTC_cdMoneda)
%>	
        <table width="60%" align="center" cellpadding="2" cellspacing="1" class="reg_Header" border="0">			        
            <tr><td colspan="4" class="reg_header_nav">AJUSTE DE CONTRATO</td></tr>
            <tr><td colspan="4"><% call showErrors() %></td></tr>
			<tr>
				<td width="8%" class="reg_Header_nav round_border_top_left"><% =GF_TRADUCIR("Contrato") %></td>
				<td width="20%" class="reg_Header_navdos" >&nbsp;<b><% =CTC_cdContrato %></b></td>				
				<td width="8%" class="reg_Header_nav round_border_top_left"><% =GF_TRADUCIR("Total Contrato") %></td>
				<td width="20%" class="reg_Header_navdos">&nbsp;<b><% =simboloMoneda & " " & GF_EDIT_DECIMALS(CTC_TotalImporte, 2) %></b></td>				
			</tr>
            <tr>
                <td class="reg_Header_nav round_border_top_left"><% =GF_TRADUCIR("Titulo contrato") %></td>
				<td class="reg_Header_navdos" >&nbsp;<b><% =CTC_Titulo %></b></td>				
				<td class="reg_Header_nav"><% =GF_TRADUCIR("Tipo") %></td>
				<td class="reg_Header_navdos" >&nbsp;<b><% =getDsTipoCTC(CTC_tipo) %>&nbsp;</b></td>
            </tr>
			<tr>
				<td class="reg_Header_nav"><% =GF_TRADUCIR("Fondo de Reparo") %></td>
				<td class="reg_Header_navdos" >&nbsp;<b><% =CTC_FReparo %>&nbsp;%</b></td>
                <td class="reg_Header_nav"><% =GF_TRADUCIR("Fecha Vto") %></td>
				<td class="reg_Header_navdos">
					<div id="issuedateDiv" class="labelStyle">
					<%  if(CDbl(CTC_fechaVto) = 0)then
							Response.Write "Sin fecha definida"
						else
							Response.Write GF_FN2DTE(CTC_fechaVto)							
						end if
					%></div>
					<input type="hidden" id="issuedate" name="issuedate" value="<% =CTC_fechaVto %>" />
				</td>								
			</tr>
			<tr>					
				<td class="reg_Header_nav"><% =GF_TRADUCIR("Responsable") %></td>
				<td class="reg_Header_navdos" >&nbsp;<b><% =CTC_dsResponsable %></b></td>
                <td class="reg_Header_nav"><% =GF_TRADUCIR("Supervisor") %></td>
				<td class="reg_Header_navdos" >&nbsp;<b><% =CTC_dsSupervisor %></b></td>
			</tr>	
			<tr>			
				<td class="reg_Header_nav"><% =GF_TRADUCIR("Pedido de Precio") %></td>
				<% if (pct_idPedido > 0) then %>
				<td class="reg_Header_navdos" >&nbsp;<b><% =pct_cdPedido & " --> " & Left(pct_tituloPedido, 30) & "..."%></b></td>
				<%  else     %>
                <td class="reg_Header_navdos" >&nbsp;<b>-</b></td>
				<% end if 				
				if (CTC_tipo = CTC_TIPO_UNITARIO) then
				%>
				<td class="reg_Header_nav"><% =GF_TRADUCIR("Valor Unitario") %></td>
				<td class="reg_Header_navdos"><% =simboloMoneda & " " & GF_EDIT_DECIMALS(CTC_valorUnitario, 2) %></td>		
				<% else %>
                <td class="reg_Header_navdos" colspan="2"></td>
                <% end if %>		
			</tr>
			<tr>
				<td class="reg_Header_nav"><% =GF_TRADUCIR("Proveedor") %></td>
				<td class="reg_Header_navdos" colspan="3" >&nbsp;<b><% =CTC_idProveedor &"-"& CTC_dsProveedor %></b></td>
			</tr>	
			<% Call drawRowObra() %>
		</table>	
		<br />
	<table align="center" width="60%" class="reg_Header">		
		<tr>
		    <td colspan="4" class="reg_header_nav">DATOS DEL AJUSTE</td>
		</tr>
		<tr class="reg_Header_nav">
			<td align="center"><% =GF_TRADUCIR("Id") %></td>
			<td align="center"><% =GF_TRADUCIR("Fecha") %></td>
			<td align="center"><% =GF_TRADUCIR("Tipo") %></td>						
			<td align="center"><% =GF_TRADUCIR("Importe") %></td>
		</tr>		

<%
	reg=0	
		while (not rsAJU.eof)
			reg = reg + 1
			%>							
			<tr class="reg_Header_navdos">
				<td align="center"><%=rsAJU("IDAJUSTE")%></td>				
				<td align="center"><%=GF_FN2DTE(Left(rsAJU("MOMENTO"), 8)) %></td>
				<td align="center">
				<%
				    if (rsAJU("TIPOAJUSTE") = CTC_AJUSTE_GENERAL) then %>
		                Ajuste del Presupuesto del Contrato
		        <%  else    %>
		                Ajuste del Valor Unitario
		        <%  end if  %>		        
				</td>
<%				if (CTC_cdMoneda = MONEDA_PESO) then		%>
				<td align="RIGHT"><% =getSimboloMoneda(MONEDA_PESO) & "&nbsp;" & GF_EDIT_DECIMALS(cDbl(rsAJU("IMPORTEPESOS")),2) %></td>
<%				else			%>				
				<td align="RIGHT"><% =getSimboloMoneda(MONEDA_DOLAR) & "&nbsp;" & GF_EDIT_DECIMALS(cDbl(rsAJU("IMPORTEDOLARES")),2) %></td>
<%				end if		
%>				
            </tr>
            <tr class="reg_Header_nav">
		        <td colspan="4" align="center"><% =GF_TRADUCIR("Observaciones") %></td>
	        </tr>
	        <tr class="reg_Header_navdos">
				<td colspan="4"><% =rsAJU("OBSERVACIONES")%></td>				
			</tr>
			<%
			rsAJU.movenext
		wend
	else %>
		<tr><td class="TDNOHAY" colspan="8"><% =GF_TRADUCIR("No se encontraron datos para mostrar") %></td></tr>
<%	end if %>
	</table>
<%
End Function
'----------------------------------------------------------------------------------
Function getRsAJU(pIdAjuste)
	dim strSQL, conn, rs
	strSQL = "Select * from TBLOBRACTCAJUSTES WHERE IDAJUSTE=" & pIdAjuste
	'Response.Write strSQL
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)							
	
	Set getRsAJU = rs
End Function
'------------------------------------------------------------------------------------------------------
Function cargarFirmas(pIdAjuste)
	Dim rsFirmas, connFirmas, strSQL
   
	Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rsFirmas, "TBLOBRACTCAJUSTESFIRMAS_GET_BY_IDAJUSTE", pIdAjuste)
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
		end if				
		rsFirmas.MoveNext()
	wend			
	
End Function
'***********************************************************************************
'*******	                     COMIENZO DE LA PAGINA                      ********
'***********************************************************************************
Dim idAjuste, newTotalPesos, newTotalDolares, saldoPesos, saldoDolares, ajustesPesos, ajustesDolares
Dim firmante1Cd, firmante2Cd, flagYaFirmo
Dim firmante1Ds, firmante2Ds
Dim firmante1Sec, firmante2Sec, firmante3Sec
Dim firmante1Tx, firmante2Tx, rolUsuario
Dim firmante3Cd, firmante3Ds, firmante3Tx
Dim firmante1Rol, firmante2Rol, firmante3Rol

idAjuste = GF_Parametros7("idAjuste",0,6)
errFirma = GF_PARAMETROS7("errFirma","",6)
if (errFirma <> "") then Call setError(errFirma)
if (idAjuste <> 0) then Call cargarFirmas(idAjuste)
rolUsuario = getRolFirma(session("Usuario"), SEC_SYS_COMPRAS)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title><% =GF_TRADUCIR("Sistema de Compras - Ajuste de Contrato") %></title>
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

    var link = "comprasFirmarAjusteCTC.asp?idAjuste=<%=idAjuste%>&secuencia=";
    var hkey0 = new Hkey('hk0', link + "<%=firmante1Sec%>", '<% =HKEY() %>', 'check_callback()');
	var hkey1 = new Hkey('hk1', link + "<%=firmante2Sec%>", '<% =HKEY() %>', 'check_callback()');
	var hkey2 = new Hkey('hk2', link + "<%=firmante3Sec%>", '<% =HKEY() %>', 'check_callback()');
	

	function check_callback(resp) {	
		if (resp != "<% =RESPUESTA_OK %>") document.getElementById("errFirma").value = resp;
		document.getElementById("frmSel").submit();
	}
	
	function bodyOnLoad(){		
		hkey0.start();
		hkey1.start();
		hkey2.start();		
    }    
</script>
</head>
<body onLoad="bodyOnLoad()">
<form method="post" id="frmSel" action="comprasAjusteCTCFirmas.asp?idAjuste=<%=idAjuste%>">
<div id="toolbar"></div><br>
<% Call mostrarAjustes(idAjuste)	
   flagYaFirmo = false		
%>
<table class="reg_header" align="center" width="60%" border="0" >
	<tr>
		<td colspan="6" class="reg_header_nav recuadroRound"><% =GF_TRADUCIR("Firmas") %></td>				
	</tr>
	<tr>
		<td colspan="6">	
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
			</table>
		</td>
	</tr>			
</table>
<input type="hidden" name="errFirma" id="errFirma">
</form>
</body>
</html>