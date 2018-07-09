<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosAFE.asp"-->
<!--#include file="Includes/procedimientosHKEY.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<%
dim errFirma,errDs

errFirma = GF_PARAMETROS7("errFirma","",6)
errDs = errMessage(errFirma)
'-----------------------------------------------------------------------------------------------
function showControl(pTipo, pId, pName, pValor, pOnClick, pCompararCon, pEtiqueta, pInitState)
dim myChecked, myImg, rtrn
if instr(pCompararCon, pValor) > 0 then
	myChecked = "CHECKED"
	if ucase(pTipo) = "CHECKBOX" then
		myImg = myImgChecked
	else
		if ucase(pName) = "CUMPLIMIENTOS" then
			if instr(afe_Tipo,AFE_TIPO_CUMPIMIENTO)>0 then 
				myImg = myImgRChecked
			else
				myImg = myImgRCheckedD 
			end if	
		else
			myImg = myImgRChecked 
		end if
	end if	
else
	myChecked = "" 
	if ucase(pTipo) = "CHECKBOX" then
		myImg = myImgUnchecked
	else
		if ucase(pName) = "CUMPLIMIENTOS" then
			if instr(afe_Tipo,AFE_TIPO_CUMPIMIENTO)>0 then 
				myImg = myImgRUnchecked 
			else
				myImg = myImgRUncheckedD 
			end if	
		else
			myImg = myImgRUnchecked 
		end if
	end if	
end if
rtrn = myImg
showControl = rtrn & UCASE(GF_Traducir(pEtiqueta))
end function
'-----------------------------------------------------------------------------------------------
Function showCategoria(pId, pValor, pOnClick, pCompararCon, pInitState)
	showCategoria = showControl("radio", pId, "categoria", pValor, pOnClick, pCompararCon, getDescripcionCategoriaAFE(pValor), pInitState)
End Function
'-----------------------------------------------------------------------------------------------
Function showTipo(pId, pValor, pOnClick, pCompararCon, pInitState)
	showTipo = showControl("checkbox", pId, "tipo", pValor, pOnClick, pCompararCon, getDescripcionTipoAFE(pValor), pInitState)
End Function
'---------------------------------------------------------------
'Muestra la firma o la opción de firmar dependiendo del usuario de la session.
Function showFirma(pCdUser, pHKey, pDsUser, pEstadoFirma, pRol, pHkNro)
	'response.write pCdUser & ", " & pHKey & ", " & pDsUser & ", " & pEstadoFirma & ", " & pRol & ", " & pHkNro & "<br>"
	Dim ret, rol 
	ret = "<br><br><br>"
	if (Trim(pHkey) <> "") then
		'1) SI LA FIRMA YA FUE REGISTRADA, MUESTRO EL FIRMANTE
		ret = "<img src='images/firmas/" & obtenerFirma(pCdUser) & "'><br>"
	else		
		if (pEstadoFirma = afe_Confirmado) then
			'2)SI EL ESTADO PROXIMO A FIRMAR COINCIDE CON EL ESTADO DEL AFE, VERIFICO SI EL USUARIO ES EL QUE DEBE FIRMAR O SI PERTENECE AL ROL INDICADO
			if (CDbl(pRol) = 0) then
				if (UCase(Session("Usuario")) = Trim(UCase(pCdUser))) then ret = "<div id='" & pHkNro & "'></div>"
			else
				if (Cdbl(getRolFirma(UCase(session("Usuario")), SEC_SYS_COMPRAS)) = Cdbl(pRol)) then ret = "<div id='" & pHkNro & "'></div>"
			end if
		end if
	end if
	ret = ret & pDsUser    
	showFirma = ret
End Function
'****************************************************************
'*******	         COMIENZO DE LA PAGINA               ********
'****************************************************************
dim idPedido, idProveedor, dsProveedor, divisionObra, dsDivision, dsObra, idObra 
dim categoria, tipo, cumplimientos, ques1, ques2, ques3, ques4, ques5, ques6, ques7, ques8, ques9, ques10, ques11
dim montoLocal, myImgRChecked, myImgRUnchecked, myImgRCheckedD, myImgRUncheckedD
dim idAFE, cdCuenta, myOnload, commentAux, index, mySplit, flagFile, statusAfeCompl
dim dirFileFinal, file, totalImpObra, totalAfesObra,g_UsuarioAFirmar,g_RolAFirmar

idAFE    = GF_Parametros7("idAFE"   ,0 ,6)
errFirma = GF_PARAMETROS7("errFirma","",6)

if (errFirma <> "") then Call setError(errFirma)

'Si no hay ni obra ni pedido, todo depende de los permisos de AFE
Call comprasControlAccesoCM(RES_AFE)	

Call GP_CONFIGURARMOMENTOS()

'Recuperar parametros querystring	

statusAfeCompl = "text"
'Fin Recuperar parametros form
dim myImgUnchecked, myImgChecked
myImgUnchecked = "&nbsp;<img src='images/icon_unchecked.gif'>&nbsp;"
myImgChecked = "&nbsp;<img src='images/icon_checked.gif'>&nbsp;"
myImgRChecked = "&nbsp;<img src='images/radio_chk1.gif'>&nbsp;"
myImgRUnchecked = "&nbsp;<img src='images/radio_chk0.gif'>&nbsp;"
myImgRCheckedD = "&nbsp;<img src='images/radio_chk1_dis.gif'>&nbsp;"
myImgRUncheckedD = "&nbsp;<img src='images/radio_chk0_dis.gif'>&nbsp;"

Call readAFE(idAFE, 0, 0)
	
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title>Sistema de Compras - AFE</title>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/MagicSearch.css" type="text/css">
<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript" src="scripts/formato.js"></script>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/MagicSearchObj.js"></script>
<script type="text/javascript" src="scripts/hkey.js"></script>
</head>
<script>
	<%if errFirma <> "" then%>
		//alert('<%=errDs%>')
	<%end if%>
	
	var link = "comprasFirmarAFE.asp?IDAFE=<%=idAFE%>";	
    var hkey1 = new Hkey('hk1', link, '<% =HKEY() %>', 'check_callback()');
	var hkey2 = new Hkey('hk2', link, '<% =HKEY() %>', 'check_callback()');
	var hkey3 = new Hkey('hk3', link, '<% =HKEY() %>', 'check_callback()');
	var hkey4 = new Hkey('hk4', link, '<% =HKEY() %>', 'check_callback()');
	var hkey5 = new Hkey('hk5', link, '<% =HKEY() %>', 'check_callback()');
	var hkey6 = new Hkey('hk6', link, '<% =HKEY() %>', 'check_callback()');	
	var hkey7 = new Hkey('hk7', link, '<% =HKEY() %>', 'check_callback()');
	var hkey8 = new Hkey('hk8', link, '<% =HKEY() %>', 'check_callback()');
	
	function bodyOnLoad(){
		var tb = new Toolbar('toolbar', 8, 'images/compras/');
		tb.draw();
        document.getElementById("estimacion").innerHTML = document.getElementById("hiddEstimacion").innerHTML;
        document.getElementById("hiddEstimacion").innerHTML = "";
        document.getElementById("requerido").innerHTML = document.getElementById("hiddRequiere").innerHTML;
        document.getElementById("hiddRequiere").innerHTML = "";
        document.getElementById("revisionTecnica").innerHTML = document.getElementById("hiddRevisionTecnica").innerHTML;
        document.getElementById("hiddRevisionTecnica").innerHTML = "";
        document.getElementById("gerente").innerHTML = document.getElementById("hiddGerente").innerHTML;
        document.getElementById("hiddGerente").innerHTML = "";
        document.getElementById("coordinador").innerHTML = document.getElementById("hiddCoordinador").innerHTML;
        document.getElementById("hiddCoordinador").innerHTML = "";
        document.getElementById("controller").innerHTML = document.getElementById("hiddController").innerHTML;
        document.getElementById("hiddController").innerHTML = "";
        document.getElementById("tesoreria").innerHTML = document.getElementById("hiddFinanzas").innerHTML;
        document.getElementById("hiddFinanzas").innerHTML = "";
        document.getElementById("director").innerHTML = document.getElementById("hiddDirector").innerHTML;
        document.getElementById("hiddDirector").innerHTML = "";
        hkey1.start();
		hkey2.start();
		hkey3.start();
		hkey4.start();
		hkey5.start();
		hkey6.start();		
		hkey7.start();
		hkey8.start(); 
	}

	function check_callback(resp) {		
		if (resp != "<% =RESPUESTA_OK %>") document.getElementById("errFirma").value = resp;		
		document.getElementById("frmSel").submit();
	}	

	function checkCancel_callback(resp) {		
		if (resp != "<% =RESPUESTA_OK %>") {
			document.getElementById("errFirma").value = resp;		
			document.getElementById("frmSel").submit();
		} else {
			window.close();
		}
	}
	
</script>
<body id="mainBody" onload="bodyOnLoad();">

<div id="toolbar"></div>
<BR>

<form method="post" id="frmSel">

<input type="hidden" name="idAFE" 	 value="<%=idAFE%>">
<input type="hidden" name="errFirma" value="<%=errFirma%>" id="errFirma">


<table border=0 width="1000px" align="center">
	<tr>
		<td colspan="3"><% Call showErrors() %></td>
	</tr>
	<tr>
		<td valign="top" rowspan="3">
			<img width="120" height="30" src="images/logo1.gif">
		</td>
		<td valign="top" align="center" rowspan="2">
			<font class="BIG"><%=GF_Traducir("AUTORIZACION PARA GASTOS (AFE)")%></font>
		</td>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td align="right"><b>AFE NO.:</b> <font style="font-size:10pt; font-weight"><%=afe_CdAFE%></font></b></td>
	</tr>		
</table>
<br>
<table border="0" width="1000px" align="center">
	<tr>
		<td width="50%">
			<!--Seccion 1-->
			<table width="100%" align="center" border="1" cellpadding="0" cellspacing="0" rules="GROUP">
				<tr>
					<td WIDTH="60%">
						<table width="100%">
							<tr>
								<td>						
									<b><%=UCASE(GF_TRADUCIR("EMPRESA"))%></b>
								</td>
							</tr>

							<tr>
								<td align="center">						
									<%=getDescripcionProveedor(CD_TOEPFER)%>
									<input type="hidden" name="dsProveedor" value="<%=afe_DsProveedor%>">
									<input type="hidden" name="idProveedor" value="<%=afe_IdProveedor%>">
								</td>
							</tr>
						</table>			
					</td>
					<td  valign=top>	
						<table width="100%">
							<tr>
								<td>		
									<b><%=UCASE(GF_TRADUCIR("DIVISION"))%></b>		
								</td>
							</tr>
							<tr>
								<td align="center">						
									<%=ucase(afe_ObraDivDs)%>
									<input type="hidden" name="idDivision" value="<%=afe_IdDivision%>">
									<input type="hidden" name="obraDivDS"  value="<%=afe_ObraDivDs%>">
								</td>
							</tr>				
						</table>			
					</td>		
				</tr>				
			</table>	
			<!--Fin Seccion 1-->
		</td>
		<td rowspan="2">
			<table width="100%" border=1 cellpadding="1" cellspacing="2" rules="groups">
			<!--Seccion 3-->
				<tr>
					<th class="reg_header_nav" colspan="2">
						<B><%=UCASE(GF_TRADUCIR("TIPO"))%></B>
					</th>
				</tr>
				<tr>
					<td>
						<%= showTipo("tipo", AFE_TIPO_MEJORA, "", afe_Tipo, "")%>
					</td>
					<td>
						<%= showTipo("tipo", AFE_TIPO_REPUESTOS, "", afe_Tipo, "")%>		
					</td>	 
				</tr>
				<tr>
					<td>
						<%= showTipo("tipo", AFE_TIPO_COMUNICACIONES, "", afe_Tipo, "")%>						
					</td>
					<td>
						<%= showTipo("tipo", AFE_TIPO_DESVIO, "", afe_Tipo, "")%>						
					</td>	 
				</tr>	
				<tr>	
					<td>
						<%= showTipo("tipo", AFE_TIPO_CAPACIDAD, "", afe_Tipo, "")%>				
					</td>	 
					<td>
						<%= showTipo("tipo", AFE_TIPO_MANTENIMIENTO, "", afe_Tipo, "")%>				
					</td>	
				</tr>
				<tr> 
					<td>
						<%= showTipo("tipo", AFE_TIPO_VEHICULOS, "", afe_Tipo, "")%>				
					</td>	 
					<td>
						<%= showTipo("tipo", AFE_TIPO_CAMBIO_OBJETIVO, "", afe_Tipo, "")%>						
					</td>
				</tr>				
				<tr>	
					<td>
						<%= showTipo("tipoEsp", AFE_TIPO_CUMPIMIENTO, "HabilitarCumplimientos(this,'cumplimientos');", afe_Tipo, "")%>	
					</td>
					<td>
						<%= showControl("radio", "cumplimientos3", "cumplimientos", AFE_TIPO_CUMPLIMIENTO_NC, "", afe_TipoCC, getDescripcionTipoAFE(AFE_TIPO_CUMPLIMIENTO_NC), "disabled")%>
						<br>
						<%= showControl("radio", "cumplimientos1", "cumplimientos",  AFE_TIPO_CUMPLIMIENTO_MA, "", afe_TipoCC, getDescripcionTipoAFE(AFE_TIPO_CUMPLIMIENTO_MA), "disabled")%>
						<br>
						<%= showControl("radio", "cumplimientos2", "cumplimientos", AFE_TIPO_CUMPLIMIENTO_SEG, "", afe_TipoCC, getDescripcionTipoAFE(AFE_TIPO_CUMPLIMIENTO_SEG), "disabled")%>
					</td>	 					
				</tr>		
				<tr>			 
					<td>
						<%= showTipo("TIPO_OTROS", AFE_TIPO_OTROS, "HabilitarSocio(this,'tipoOtros')", afe_Tipo, "")%>								
						<% if printIt then 
								Response.write UCASE(afe_TipoOtros)
							else %>
								<input type="text" size="15" disabled name="tipoOtros" id="tipoOtros" value="<%=afe_TipoOtros%>">
						<% end if %>		
					</td>		
					<td>&nbsp;</td>
				</tr>
			<!--Fin Seccion 3-->	
			</table>		
		</td>
	</tr>
	<tr>
		<td>		
			<!--Seccion 2--> 
			<table width="100%" border=1 cellpadding="1" cellspacing="2" rules="groups">
				<tr>
					<th class="reg_header_nav" colspan="2">
						<B><%=UCASE(GF_TRADUCIR("CATEGORIA"))%></B>
					</th>
				</tr>
				<tr>
					<td>
						<%= showCategoria("categoria", AFE_CATEGORIA_CAPITAL, "DeshabilitarExtras(1)", afe_categoria, "")%>
					</td>	 
					<td>
						<%= showCategoria("categoria", AFE_CATEGORIA_GASTOS, "DeshabilitarExtras(1)", afe_categoria, "")%>
					</td>
				</tr>
				<tr>
					<td>
						<%= showCategoria("categoria", AFE_CATEGORIA_INVERSIONES, "DeshabilitarExtras(1)", afe_categoria, "")%>			
					</td>	 
					<td>
						<%= showCategoria("categoria", AFE_CATEGORIA_SERVICIOS, "DeshabilitarExtras(1)", afe_categoria, "")%>						
					</td>	 
				</tr>	
				<tr>	
					<td>
						<%= showCategoria("categoria", AFE_CATEGORIA_ALQUILER, "DeshabilitarExtras(1)", afe_categoria, "")%>						
					</td>	 
					<td>
						
					</td>	 
				</tr>
				<tr>
					<td COLSPAN="2">
				<%= showCategoria("CAT_AFEC", AFE_CATEGORIA_COMPLEMENTARIO, "DeshabilitarExtras(1);HabilitarSocio(this,'nroAFEComplID')", afe_categoria, "")%>			
							<%
							if printIt then
								Response.write getCdAFE(afe_NroAFEComplID)
							else
								%>
								<select disabled name="nroAFEComplID" id="nroAFEComplID">				
									<option value="0">N/A
									<%	
										if (afe_idObra > 0) then
											Set afeRaiz = readAFEObra(afe_idObra, AFE_RAIZ)
										else
											Set afeRaiz = readAFEPedido(afe_IdPedido, AFE_RAIZ)
										end if
										while (not afeRaiz.eof)
												if (afe_NroAFEComplID = afeRaiz("IDAFE")) then
												%>
													<option value="<% = afeRaiz("IDAFE") %>" selected="true"><% = afeRaiz("CDAFE") %>
												<%
												else
												%>
													<option value="<% = afeRaiz("IDAFE") %>"><% = afeRaiz("CDAFE") %>
												<%
												end if
											afeRaiz.MoveNext()
										wend
										%>
								</select>
							<%
							end if 
							%>										
					</td>	 
				</tr>
				<tr>
					<td COLSPAN="2">
						<%= showCategoria("CAT_OTROS", AFE_CATEGORIA_OTROS, "DeshabilitarExtras(1);HabilitarSocio(this,'catOtros')", afe_categoria, "")%>						
						<% if printIt then 
								Response.write UCASE(afe_CatOtros)
							else %>
								<input disabled type="text" size="50" name="catOtros" id="catOtros" value="<%=afe_CatOtros%>">
						<% end if %>
					</td>
				</tr>
			<!--Fin Seccion 2-->				
			</table>			
		</td>
	</tr>
<table>
<br>
<table align="center" width="1000px">
	<tr>
		<td>		
			<b><%=GF_TRADUCIR("PROYECTO/PARTIDA NO.")%></b>		
		</td>
	</tr>
	<tr>
		<td align="center">
			<table border="1" width="100%" rules="ALL">
				<tr>
					<td align="center">
						<%=GF_TRADUCIR("Ptda. Presup.")%>
					</td>
					<td colspan="2" align="center">
						<%=GF_TRADUCIR("Detalle")%>
					</td>
				</tr>
				<tr>
					<td align="center" width="80%">
						<%if (afe_ObraCD = "") then 
							Response.Write "-"
						else							
							Response.Write afe_ObraCD & "-" & getDescripcionObra(afe_idObra)
						end if	
						%>
					</td>
					<td align=center>
						<% =afe_IDArea %>
					</td>
					<td align=center>
						<% =afe_IDDetalle %>
					</td>
				</tr>
			</table>
			<input type="hidden" name="cdObra" value="<%=afe_ObraCD%>">
		</td>
	</tr>				
</table>
<br>
<table  align="center" width="1000px" border="0" cellpadding="0" cellspacing="0" rules="GROUP">
	<tr>		
		<td  colspan="2">		
			<table width="100%">
				<tr>
					<td>		
						<b><%=GF_TRADUCIR("TITULO DEL AFE:")%></b>		
					</td>
				</tr>
				<tr>
					<td align="left">			
							
						<% Response.write ucase(afe_titulo) %>							
									
					</td>
				</tr>				
			</table>			
		</td>
	</tr>
</table>
<br>
<!--Seccion 4-->
<table  width="1000px" align="center" border="0" cellpadding="0" cellspacing="0" rules="GROUP">
	<tr>
		<td>
			<b><%=UCASE(GF_TRADUCIR("DESCRIPCION:"))%></b>		
		</td>
	</tr>			
	<tr>
		<td>
			<textarea style="border: 0px solid" COLS="160" ROWS=10><%=replace(afe_Descripcion,ENTER_SYMBOL,chr(10))%></textarea>			
		</td>
	</tr>
</table>
<br>
<!--Seccion 5-->
<table width="1000px" align="center" border="1" cellpadding="0" cellspacing="0" rules="all">
	<tr>
		<td WIDTH="25%">
			<b><%=GF_TRADUCIR("TOTAL x PRESUPUESTO (US$)")%></b>
		</td>
		<td align="center" WIDTH="25%">
			<%	totalImpObra = cdbl(calcularCostoEstimadoObra(MONEDA_DOLAR, afe_IdObra, afe_IDArea, afe_IDDetalle))%>
			<% =GF_EDIT_DECIMALS(totalImpObra,2) %>
		</td>
		<td WIDTH="25%">
			<b><%=GF_TRADUCIR("TOTAL AFEs PREVIOS APROBADOS (US$)")%></b>		
		</td>
		<td align="center" WIDTH="25%">
			<% totalAfesObra = totalizarAFESObra(MONEDA_DOLAR, afe_IdObra, afe_IDArea, afe_IDDetalle, false) %>
			<% =GF_EDIT_DECIMALS(totalAfesObra,2) %>
		</td>	
	</tr>	
	<tr>
		<td WIDTH="25%">
			<b><%=GF_TRADUCIR("IMPORTE DEL AFE (US$)")%></b>
		</td>
		<td align="center" WIDTH="25%">
			<% =GF_EDIT_DECIMALS(cdbl(afe_ImporteDolares),2) %>			
		</td>
		<td WIDTH="25%">
			<b><%=GF_TRADUCIR("MONTO TOTAL EN MONEDA LOCAL")%></b>		
		</td>
		<td align="center" WIDTH="25%">
			<% =GF_EDIT_DECIMALS(cdbl(afe_ImportePesos),2) %>
		</td>	
	</tr>	
	<tr>
		<td WIDTH="25%">
			<b><%=GF_TRADUCIR("TIPO DE CAMBIO")%></b>
		</td>
		<td align="center" WIDTH="25%">
			<%=afe_TipoCambio%>			
		</td>
		<td WIDTH="25%">
			<b><%=GF_TRADUCIR("CÓDIGO DE MONEDA")%></b>		
		</td>
		<td align="center" WIDTH="25%"><%=getSimboloMonedaLetras(MONEDA_PESO)%></td>	
	</tr>	
</table>
<br>
<table width="1000px" align="center" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td width="500px">
			<table width="100%" align="center" border="1" cellpadding="0" cellspacing="0" rules="all">
				<tr>
					<td width="25%"><%=GF_TRADUCIR("NPV")%></b></td>
					<td width="25%" align="center">
						<% 
						if cdbl(afe_NPV) = 0 then
							Response.Write "NA"
						else
							Response.Write GF_EDIT_DECIMALS(cdbl(afe_NPV),2)
						end if							
						%>
					</td>
					<td width="25%"><%=GF_TRADUCIR("IRR")%></b></td>
					<td width="25%" align="center">
						<% 
						if cdbl(afe_Irr) = 0 then
							Response.Write "NA"
						else
							Response.Write GF_EDIT_DECIMALS(cdbl(afe_Irr),2)
						end if
						%>
					</td>		
				</tr>
				<tr>
					<td width="25%"><%=GF_TRADUCIR("ROIC")%></b></td>
					<td width="25%" align="center">
						<% 
						if cdbl(afe_ROIC) = 0 then
							Response.Write "NA"
						else
							Response.Write GF_EDIT_DECIMALS(cdbl(afe_ROIC),2)
						end if							
						%>	
					</td>
					<td width="25%"><%=GF_TRADUCIR("PAYBACK")%></b></td>
					<td width="25%" align="center">
						<% 
						if cdbl(afe_PAYBACK) = 0 then
							Response.Write "NA"
						else
							Response.Write GF_EDIT_DECIMALS(cdbl(afe_PAYBACK),2)
						end if									
						%>						
					</td>		
				</tr>
			</table>
		</td>
		<td>&nbsp;</td>
	</tr>
</table>
<br>
<table width="1000px" align="center" border="0" cellpadding="0" cellspacing="0" rules="all">
	<tr>
		<td>
			<table align="center" width="80%" border="1" cellpadding="0" cellspacing="0">
				<tr>
					<td COLSPAN="2" ALIGN="CENTER">
						<B><%=GF_TRADUCIR("REVISION Y APROBACION")%></B>
					</td>
				</tr>
				<tr>
					<td WIDTH="50%">
						<table width="100%" border=0 cellpadding=0 cellspacing=0>
							<tr>
								<td WIDTH="100%">
									<b><%=GF_TRADUCIR("ESTIMACIÓN HECHA POR")%></b>
								</td>
							</tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td align="center">
                                <div id="estimacion"></div>
							</td></tr>
						</table>					
					</td>					
					<td WIDTH="50%">
						<table width="100%" border=0 cellpadding=0 cellspacing=0>
							<tr>
								<td WIDTH="100%">
									<b><%=GF_TRADUCIR("GERENTE PUERTO")%></b>
								</td>
							</tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td align="center">
							    <div id="gerente"></div>
							</td></tr>
						</table>							
					</td>
				</tr>
				<tr>
					<td>
						<table width="100%" border=0 cellpadding=0 cellspacing=0>
							<tr>
								<td WIDTH="100%">
									<b><%=GF_TRADUCIR("GASTO REQUERIDO POR")%></b>
								</td>
							</tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td align="center">
							    <div id="requerido"></div>
							</td></tr>
						</table>							
					</td>
					<td>
						<table width="100%" border=0 cellpadding=0 cellspacing=0>
							<tr>
								<td WIDTH="100%">
									<b><%=GF_TRADUCIR("COORDINADOR PUERTO")%></b>
								</td>
							</tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td align="center">
							    <div id="coordinador"></div>
							</td></tr>
						</table>							
					</td>
				</tr>

				<tr>
					<td>
						<table width="100%" border=0 cellpadding=0 cellspacing=0>
							<tr>
								<td WIDTH="100%">
									<b><%=GF_TRADUCIR("REVISIÓN TÉCNICA")%></b>
								</td>
							</tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td align="center">
							    <div id="revisionTecnica"></div>
							</td></tr>
						</table>							
					</td>
					<td>
						<table width="100%" border=0 cellpadding=0 cellspacing=0>
							<tr>
								<td WIDTH="100%">
									<b><%=GF_TRADUCIR("DIRECTOR LOCAL")%></b>
								</td>
							</tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td align="center">
							    <div id="director"></div>
							</td></tr>
						</table>							
					</td>
				</tr>
				<tr>
					<td>
						<table width="100%" border=0 cellpadding=0 cellspacing=0>
							<tr>
								<td>
									<b><%=GF_TRADUCIR("CONTROLLER LOCAL")%></b>	
								</td>
							</tr>
							<tr><td>&nbsp;</td></tr>
							<tr>
								<td align = 'center'>
									<div id="controller"></div>
								</td>
							</tr>
						</table>
					</td>
					<td>
						<b><%=GF_TRADUCIR("BUSINESS DEVELOPMENT TEAM")%></b>		
						<BR><BR><BR><BR>&nbsp;<DIV align="center"></div>
					</td>
				</tr>
				<tr>
					<td>
						<table width="100%" border=0 cellpadding=0 cellspacing=0>
							<tr>
								<td WIDTH="100%">
									<b><%=GF_TRADUCIR("CFO / TESORERÍA")%></b>
								</td>
							</tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td align="center">
							    <div id="tesoreria"></div>
							</td></tr>
						</table>	
					</td>
					<td>
						<b><%=GF_TRADUCIR("MANAGING DIRECTOR CONTROLLING/ACCOUNTING OF ACTI")%></b>		
						<BR><BR><BR><BR>&nbsp;<DIV align="center"></div>
					</td>
				</tr>
			</table>	
		</td>
	</tr>
    
    <%  strSQL = "SELECT ESTADOACTUAL,DATOAUXILIAR AS ROL FROM TBLESTADOSTRANSICION WHERE IDSISTEMA ="&SEC_SYS_COMPRAS&" AND TIPOOBJETO = '"&RES_AFE&"' AND DATOAUXILIAR2 = '"&afe_IdDivision & afe_isCFO & "' AND EVENTO = 1  ORDER BY isnumeric(ESTADOACTUAL), ESTADOACTUAL "        
        'response.Write strSQL & "<br>"
        Call executeQueryDb(DBSITE_SQL_INTRA, rsAFE, "OPEN", strSQL) %>
        <div id="hiddEstimacion" style="display:none;">
        <%  'PREPARA AFE
	'response.write "ACA->" & rsAFE("ESTADOACTUAL") & "<br>"
            if (not rsAFE.Eof) then
                Response.Write showFirma(afe_PreparedByCD ,afe_PreparedByHkey ,afe_PreparedBy ,rsAFE("ESTADOACTUAL"),rsAFE("ROL"),"hk1")
                rsAFE.MoveNext()
            end if %>
        </div>
        <div id="hiddRequiere">
        <%  'PREPARA AFE
	'response.write "ACA->" & rsAFE("ESTADOACTUAL") & "<br>"
            if (not rsAFE.Eof) then
                Response.Write showFirma(afe_RequestedByCD ,afe_RequestedByHkey ,afe_RequestedBy ,rsAFE("ESTADOACTUAL"),rsAFE("ROL"),"hk2")
                rsAFE.MoveNext()
            end if %>
        </div>
        <div id="hiddRevisionTecnica" style="display:none;">
        <%  'REVISION TECNICA
	'response.write "ACA->" & rsAFE("ESTADOACTUAL") & "<br>"
            if (not rsAFE.Eof) then
                Response.Write showFirma(afe_EngReviewCD ,afe_EngReviewHkey ,afe_EngReview ,rsAFE("ESTADOACTUAL"),rsAFE("ROL"),"hk3")
                rsAFE.MoveNext()
            end if %>
        </div>
        <div id="hiddGerente" style="display:none;">
        <%  'GERENTE DE PUERTO
	'response.write "ACA->" & rsAFE("ESTADOACTUAL") & "<br>"
            if (afe_IdDivision <> DIVSION_EXPORTACION) then
                if (not rsAFE.Eof) then
                    Response.Write showFirma(afe_OfficerCD ,afe_OfficerHkey ,afe_Officer ,rsAFE("ESTADOACTUAL"),rsAFE("ROL"),"hk4")
                    rsAFE.MoveNext()
                end if
            end if %>
        </div>
        <div id="hiddCoordinador" style="display:none;">
        <%  'COORDINADOR DE PUERTO
	'response.write "ACA->" & rsAFE("ESTADOACTUAL") & "<br>"
            if (not rsAFE.Eof) then
                 Response.Write showFirma(afe_VicePresidentCD ,afe_VicePresidentHkey ,afe_VicePresident ,rsAFE("ESTADOACTUAL"),rsAFE("ROL"),"hk5")
                 rsAFE.MoveNext()
            end if %>
        </div>
        <div id="hiddController" style="display:none;">
        <%  'CONTROLLER
	'response.write "ACA->" & rsAFE("ESTADOACTUAL") & "<br>"
            if (not rsAFE.Eof) then
                 Response.Write showFirma(afe_controllerCD ,afe_controllerHkey ,afe_controller ,rsAFE("ESTADOACTUAL"),rsAFE("ROL"),"hk6")
                 rsAFE.MoveNext()
            end if %>
        </div>
        <div id="hiddFinanzas" style="display:none;">
        <%  'FINANZAS	
	'response.write "ACA->" & rsAFE("ESTADOACTUAL") & "<br>"
            if (afe_isCFO = TIPO_AFIRMACION) then
                 if (not rsAFE.Eof) then
                    Response.Write showFirma(afe_cfoCD ,afe_cfoHkey ,afe_cfo ,rsAFE("ESTADOACTUAL"),rsAFE("ROL"),"hk7")
                    rsAFE.MoveNext()
		'JAS-- else   
		'JAS--	rsAFE.MoveNext() 
                 end if             
            end if %>
        </div>
        <div id="hiddDirector" style="display:none;">
        <%  'DIRECTOR
            if (not rsAFE.Eof) then	    
                 Response.Write showFirma(afe_PresidentCD ,afe_PresidentHkey ,afe_President ,rsAFE("ESTADOACTUAL"),rsAFE("ROL"),"hk8")
                 rsAFE.MoveNext()
            end if %>
        </div>
</table>
</form>
</body>
</html>

