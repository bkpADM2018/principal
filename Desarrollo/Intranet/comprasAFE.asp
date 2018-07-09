<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->
<!--#include file="Includes/procedimientosAFE.asp"-->
<!--#include file="Includes/procedimientosHKEY.asp"-->
<!-- #include file="Includes/procedimientosUser.asp"-->


<%
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

	if printIt then
		rtrn = myImg
	else
		rtrn = "<input STYLE='cursor:pointer;border:none;' id='" & pId & "' type='" & pTipo & "' value='" & pValor & "' name='" & pName & "' " & myChecked & " onclick=" & pOnClick & ">"
	end if
	showControl = rtrn & UCASE(pEtiqueta)
end function
'-----------------------------------------------------------------------------------------------
Function showCategoria(pId, pValor, pOnClick, pCompararCon, pInitState)
	showCategoria = showControl("radio", pId, "categoria", pValor, pOnClick, pCompararCon, getDescripcionCategoriaAFE(pValor), pInitState)
End Function
'-----------------------------------------------------------------------------------------------
Function showTipo(pId, pValor, pOnClick, pCompararCon, pInitState)
	showTipo = showControl("checkbox", pId, "tipo", pValor, pOnClick, pCompararCon, getDescripcionTipoAFE(pValor), pInitState)
End Function
'****************************************************************
'Se verifican si estan todos los datos necesarios para cargar un AFE.
Function checkCargaAFE(idPedido, idAFE, idObra)
	dim strSQL, rs, conn
		
	checkCargaAFE=false
	if (idAFE = 0) then
		'1 - Se checkea que exista una cotizacion cargada.
		strSQL = "Select * from TBLDATOSOBRAS where IDOBRA=" & idObra	
		Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if (not rs.eof) then		
			checkCargaAFE = true
		else		
			strSQL = "Select * from TBLPCPCABECERA where IDPEDIDO=" & idPedido
			Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
			if (not rs.eof) then		
				checkCargaAFE = true			
			end if
		end if
	else
		checkCargaAFE = true
	end if
End Function 
'****************************************************************
'*******	         COMIENZO DE LA PAGINA               ********
'****************************************************************
dim idPedido, idProveedor, dsProveedor, divisionObra, dsDivision, dsObra, idObra 
dim categoria, tipo, cumplimientos, ques1, ques2, ques3, ques4, ques5, ques6, ques7, ques8, ques9, ques10, ques11
dim montoLocal, printIt, myImgRChecked, myImgRUnchecked, myImgRCheckedD, myImgRUncheckedD
dim idAFE, cdCuenta, myOnload, commentAux, index, mySplit, flagFile, statusAfeCompl
dim dirFileFinal, file, blnGrabo, nroAFEAnula

idObra = GF_Parametros7("idObra",0,6)	
idPedido = GF_Parametros7("idPedido",0,6)
idAFE = GF_Parametros7("idAFE",0,6)

if (idPedido > 0) then	'Si indico un pedido, veo si puedo acceder al pedido
	Call comprasControlAccesoCM(RES_CC)		
elseif (idObra > 0) then 'Si mando una obra, chequeo permisos de obra
	Call comprasControlAccesoCM(RES_OBR)	
else	'Si no hay ni obra ni pedido, todo depende de los permisos de AFE
	Call comprasControlAccesoCM(RES_AFE)	
end if

Call GP_CONFIGURARMOMENTOS()

'SE VERIFICA SI SE POSEEN TODOS LOS DATOS NECESARIOS PARA CARGAR UN AFE O VISUALIZAR UNO EXISTENTE.
if (checkCargaAFE(idPedido, idAFE, idObra)) then

	'Recuperar parametros querystring	
	accion = GF_Parametros7("accion","",6)
	printIt = true
	statusAfeCompl = "text"
	dirFileFinal = ""
	'Fin Recuperar parametros form
	dim myImgUnchecked, myImgChecked
	myImgUnchecked = "&nbsp;<img src='images/icon_unchecked.gif'>&nbsp;"
	myImgChecked = "&nbsp;<img src='images/icon_checked.gif'>&nbsp;"
	myImgRChecked = "&nbsp;<img src='images/radio_chk1.gif'>&nbsp;"
	myImgRUnchecked = "&nbsp;<img src='images/radio_chk0.gif'>&nbsp;"
	myImgRCheckedD = "&nbsp;<img src='images/radio_chk1_dis.gif'>&nbsp;"
	myImgRUncheckedD = "&nbsp;<img src='images/radio_chk0_dis.gif'>&nbsp;"	
	Call readAFE(idAFE, idObra, idPedido)
	if (idAFE > 0) then idObra = afe_idObra		
	if (esEditable(idAFE)) then printIt = false
	blnGrabo=false
	if accion <> "" then	
		if ControlAFE(afe_NroAFEComplID, divisionObraID, afe_Categoria, afe_CatOtros, afe_Tipo, afe_TipoOtros, afe_TipoCC, afe_Descripcion, afe_Titulo, afe_IDArea,afe_IDDetalle,afe_IdObra, afe_PreparedByCD, afe_RequestedByCD, afe_EngReviewCD) then
			'Puede grabar
			commentAux = replace(commentAux,"'","*")
			'commentAux = replace(afe_Descripcion,chr(10),ENTER_SYMBOL)
			'if (left(commentAux,4) = ENTER_SYMBOL) then
			'	commentAux = mid(commentAux,5,len(commentAux))
			'end if
			'afe_Descripcion = commentAux
			nroAFEAnula = 0
			Call addAFE(idAFE, afe_cdAFE, afe_IdObra, idPedido, afe_ObraCuentaDS, afe_NroAFEComplID, afe_titulo, afe_IdDivision, afe_Departamento, afe_Categoria, afe_CatOtros, afe_Tipo, afe_TipoOtros, afe_TipoCC, afe_Descripcion, afe_ImportePesos, afe_ImporteDolares, afe_TipoCambio, afe_NPV, afe_IRR, afe_ROIC, afe_PAYBACK, afe_PreparedByCD, afe_RequestedByCD, afe_EngReviewCD, afe_isCFO, afe_IdArea, afe_IdDetalle, nroAFEAnula)			
			blnGrabo = true
		end if						
	end if

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
		
	function HabilitarCumplimientos(pObj, pSoc){
		if (pObj.checked==true){
			document.getElementById(pSoc + '1').disabled = false;
			document.getElementById(pSoc + '2').disabled = false;
			document.getElementById(pSoc + '3').disabled = false;
		}
		else{
			document.getElementById(pSoc + '1').disabled = true;
			document.getElementById(pSoc + '2').disabled = true;
			document.getElementById(pSoc + '3').disabled = true;
		}
	}	
	function HabilitarSocio(pObj, pSoc){						
		if (pObj.checked==true){
			document.getElementById(pSoc).disabled = false;
		}
		else{
			document.getElementById(pSoc).disabled = true;
		}
	}	
	function DeshabilitarExtras(pOpc){
			document.getElementById("nroAFEComplID").selectedIndex = 0;
			document.getElementById("nroAFEComplID").disabled = true;
			document.getElementById("catOtros").value = "";
			document.getElementById("catOtros").disabled = true;
	}
	function bodyOnLoad(){
		<% if (blnGrabo) then %>
			window.close()
		<% end if %>
		var tb = new Toolbar('toolbar', 8, 'images/compras/');
		tb.addButton("save-16x16.png", "<%=GF_Traducir("Guardar")%>", "submitForm()");			
		tb.draw();
		
		var myObj = document.getElementById("tipoEsp");
			HabilitarCumplimientos(myObj,"cumplimientos");
			HabilitarSocio(document.getElementById("CAT_AFEC"), "nroAFEComplID");
			HabilitarSocio(document.getElementById("CAT_OTROS"), "catOtros");
			HabilitarSocio(document.getElementById("TIPO_OTROS"), "tipoOtros");

			var msSign0 = new MagicSearch("", "preparedBy", 30, 2, "comprasStreamElementos.asp?tipo=personas");
			msSign0.setToken(";");
			msSign0.onBlur = seleccionarFirmante;
			msSign0.setValue('<% =afe_PreparedBy %>');	
			
			var msSign2 = new MagicSearch("", "requestedBy", 30, 2, "comprasStreamElementos.asp?tipo=personas");
			msSign2.setToken(";");
			msSign2.onBlur = seleccionarFirmante;
			msSign2.setValue('<% =afe_RequestedBy %>');	
			
			var msSign3 = new MagicSearch("", "engReview", 30, 2, "comprasStreamElementos.asp?tipo=personas");
			msSign3.setToken(";");
			msSign3.onBlur = seleccionarFirmante;
			msSign3.setValue('<% =afe_EngReview %>');	
			
	}

function seleccionarFirmante(ms) {
	var desc = ms.getSelectedItem();
	if (desc.indexOf('-') != -1) {
		var arr = desc.split('-');
		document.getElementById(ms.nameDIV + "CD").value = arr[0];
		ms.setValue(arr[1]);
	} else {
		if (desc == "") document.getElementById(ms.nameDIV + "CD").value = "";
	}		
}
function sumarTotal(current) {
		var tipoCambio = <% =getTipoCambio(MONEDA_DOLAR, "") %>;
		var totalDolares = 0;		
		var totalPesos = 0;
		var objP = document.getElementById("importePesos");
		var objD = document.getElementById("importeDolares");
		if (current == "D") {			
			objP.value = objD.value.replace(/,/,".") * tipoCambio;						
		} else {
			objD.value = objP.value.replace(/,/,".") / tipoCambio;						
		}		
		objP.value = editarImporte(objP.value);			
		objD.value = editarImporte(objD.value);
	}	
	function submitForm() {					
		document.getElementById("accion").value="<% =ACCION_GRABAR %>";
		document.getElementById("confirmado").value="<% =AFE_NO_CONFIRMADO %>";
		document.getElementById("frmSel").submit();
	}				
</script>
<body id="mainBody" onLoad="bodyOnLoad();">
<div id="toolbar"></div>
<BR>
<form method="post" id="frmSel">

<input type="hidden" name="idAFE" 	 value="<%=idAFE%>">
<input type="hidden" name="cdAFE" 	 value="<%=afe_CdAFE%>">
<input type="hidden" name="idPedido" value="<%=afe_IdPedido%>">
<input type="hidden" name="idObra"   value="<%=afe_IdObra%>">

<input type="hidden" name="confirmado" 		id="confirmado" 	 value="<%=afe_Confirmado%>">
<input type="hidden" name="accion" 			id="accion" 	   	 value="">
<input type="hidden" name="preparedByCD" 	id="preparedByCD" 	 value="<%=afe_PreparedByCD%>">
<input type="hidden" name="requestedByCD" 	id="requestedByCD" 	 value="<%=afe_RequestedByCD%>">
<input type="hidden" name="engReviewCD" 	id="engReviewCD" 	 value="<%=afe_EngReviewCD%>">
<input type="hidden" name="officerCD" 		id="officerCD" 		 value="<%=afe_OfficerCD%>">
<input type="hidden" name="vicePresidentCD" id="vicePresidentCD" value="<%=afe_VicePresidentCD%>">
<input type="hidden" name="presidentCD" 	id="presidentCD" 	 value="<%=afe_PresidentCD%>">
<input type="hidden" name="controllerCD" 	id="controllerCD" 	 value="<%=afe_ControllerCD%>">
<input type="hidden" name="cfoCD" 			id="cfoCD" 			 value="<%=afe_CFOCD%>">

<table border=0 width="1000px" align="center">
	<tr>
		<td colspan=3><% Call showErrors() %></td>
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
		<td align="right"><b>AFE NO.: <%=afe_CdAFE%> </b></td>
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
										
									<%	if (idObra > 0) then
											Set afeRaiz = readAFEObra(idObra, AFE_RAIZ)
										else
											Set afeRaiz = readAFEPedido(idPedido, AFE_RAIZ)
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
							Response.Write afe_ObraCD & "-" & getDescripcionObra(idObra)
						end if	
						%>
					</td>
					<td align=center>
						<%if ((afe_ObraCD <> "") and (not isInversion(afe_IdObra))) then%><input type="text" size="4" id="idArea" name="idArea" onKeyPress="return controlIngreso(this, event, 'N')" value="<% =afe_IDArea %>"><%end if%>
					</td>
					<td align=center>
						<%if ((afe_ObraCD <> "") and (not isInversion(afe_IdObra))) then%><input type="text" size="4" id="idDetalle" name="idDetalle" onKeyPress="return controlIngreso(this, event, 'N')" value="<% =afe_IDDetalle %>"></div><%end if%>
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
						<% if printIt then 
								Response.write ucase(afe_titulo)
							else %>
								<input maxlength="150" size="193" type="text" name="afe_titulo" value="<%=afe_titulo%>">
						<% end if %>							
									
					</td>
				</tr>				
			</table>			
		</td>
	</tr>
</table>
<br>
<!--Seccion 4-->
<table  align="center" width="1000px" border="0" cellpadding="0" cellspacing="0" rules="GROUP">
	<tr>
		<td>
			<b><%=UCASE(GF_TRADUCIR("DESCRIPCION:"))%></b>		
		</td>
	</tr>			
	<tr>
		<td>
			<% if printIt then 
					Response.write replace(afe_Descripcion," ","&nbsp;")
				else 
					'afe_Descripcion = replace(afe_Descripcion,ENTER_SYMBOL,chr(10))
					%>
					<textarea COLS="160" ROWS="13" name="descripcion"><%=afe_Descripcion%></textarea>
			<% end if %>			
		</td>
	</tr>
</table>			
<!--Fin Seccion 4-->
<br>
<!--Seccion 5-->
<table width="1000px" align="center" border="1" cellpadding="0" cellspacing="0" rules="all">
	<tr>
		<td WIDTH="25%">
			<b><%=GF_TRADUCIR("GASTO TOTAL (USD)")%></b>
		</td>
		<td align="center" WIDTH="25%">
			<input style="text-align:right;" size="12" type="text" id="importeDolares" name="importeDolares" value="<%=afe_importeDolares/100%>" onKeyPress="return controlIngreso(this, event, 'I')" onBlur="sumarTotal('D')">								
		</td>
		<td WIDTH="25%">
			<b><%=GF_TRADUCIR("MONTO TOTAL EN MONEDA LOCAL")%></b>		
		</td>
		<td align="center" WIDTH="25%">
			<input style="text-align:right;" size="12" type="text" id="importePesos" name="importePesos" value="<%=afe_importePesos/100%>" onKeyPress="return controlIngreso(this, event, 'I')" onBlur="sumarTotal('P')">								
		</td>	
	</tr>	
	<tr>
		<td WIDTH="25%">
			<b><%=GF_TRADUCIR("TIPO DE CAMBIO")%></b>
		</td>
		<td align="center" WIDTH="25%">
			<%=afe_TipoCambio%>
			<input type="hidden" name="tipoCambio" value="<%=afe_TipoCambio%>">
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
					<td width="25%">NPV</b></td>
					<td width="25%" align="center">
						<% 
						if printIt then 
							if cdbl(afe_NPV) = 0 then
								Response.Write "NA"
							else
								Response.Write GF_EDIT_DECIMALS(cdbl(afe_NPV),2)
							end if							
						else %>
							<input style="text-align:right;" type="text" size="5" name="Arr" value="<%=GF_EDIT_DECIMALS(cdbl(afe_NPV),2)%>" onKeyPress="return controlIngreso(this, event, 'I')">
						<% end if %>	
					</td>
					<td width="25%">IRR</b></td>
					<td width="25%" align="center">
							<% 
							if printIt then 
								if cdbl(afe_Irr) = 0 then
									Response.Write "NA"
								else
									Response.Write GF_EDIT_DECIMALS(cdbl(afe_Irr),2)
								end if
							else %>
								<input style="text-align:right;" type="text" size="5" name="Irr" value="<%=GF_EDIT_DECIMALS(cdbl(afe_Irr),2)%>" onKeyPress="return controlIngreso(this, event, 'I')">
							<% end if %>								
					</td>		
				</tr>
				<tr>
					<td width="25%">ROIC</b></td>
					<td width="25%" align="center">
						<% 
						if printIt then 
							if cdbl(afe_ROIC) = 0 then
								Response.Write "NA"
							else
								Response.Write GF_EDIT_DECIMALS(cdbl(afe_ROIC),2)
							end if							
						else %>
							<input style="text-align:right;" type="text" size="5" name="RONA" value="<%=GF_EDIT_DECIMALS(cdbl(afe_ROIC),2)%>" onKeyPress="return controlIngreso(this, event, 'I')">
						<% end if %>	
					</td>
					<td width="25%">PAYBACK</b></td>
					<td width="25%" align="center">
						<% 
						if printIt then 
							if cdbl(afe_PAYBACK) = 0 then
								Response.Write "NA"
							else
								Response.Write GF_EDIT_DECIMALS(cdbl(afe_PAYBACK),2)
							end if									
						else %>
							<input style="text-align:right;" type="text" size="5" name="PAYBACK" value="<%=GF_EDIT_DECIMALS(cdbl(afe_PAYBACK),2)%>" onKeyPress="return controlIngreso(this, event, 'I')">
						<% end if %>	
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
					<td align="center"><B><%=GF_TRADUCIR("FIRMAS A DEFINIR")%></B></td>
					<td align="center"><B><%=GF_TRADUCIR("FIRMAS YA DEFINIDAS POR EL SISTEMA")%></B></td>
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
							<% 
							if printIt then 
								Response.Write afe_PreparedBy
							else %>
								<div id="preparedBy"></div>
							<% end if %>
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
							<% 
							Response.Write afe_Officer %>								
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
							<% 
							if printIt then 
								Response.Write afe_RequestedBy
							else %>
								<div id="requestedBy"></div>
							<% end if %>								
							</td></tr>
						</table>							
					</td>
					<td>
						<table width="100%" border=0 cellpadding=0 cellspacing=0>
							<tr>
								<td WIDTH="100%">
									<b><%=GF_TRADUCIR("COORDINADOR PUERTOS")%></b>
								</td>
							</tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td align="center">
							<%  Response.Write afe_VicePresident %>
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
							<% 
							if printIt then 
								Response.Write afe_EngReview
							else %>
								<div id="engReview"></div>
							<% end if %>								
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
							<% Response.Write afe_President %>
							</td></tr>
						</table>							
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
							<% 
							if printIt then 
								Response.Write afe_cfo
							else %>
								<input type="checkbox" id="chkFinanzas" name="chkFinanzas" <% If(afe_isCFO = TIPO_AFIRMACION) then %> checked <% end if %> /> Si, el AFE debe ser firmado por Finanazas.
							<% end if %>
							</td></tr>
						</table>	
					</td>
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
									<% Response.Write afe_controller %>
								</td>
							</tr>
						</table>
					</td>
				</tr>	
				<tr>
					<td>&nbsp;</td>
					<td>
						<b><%=GF_TRADUCIR("Business Development Team")%></b>		
						<BR>&nbsp;<DIV align="center"></div>
					</td>
				</tr>
				<tr>
					<td>&nbsp;</td>
					<td>
						<b><%=GF_TRADUCIR("MANAGING DIRECTOR, ACTI HH")%></b>		
						<BR>&nbsp;<DIV align="center"></div>
					</td>
				</tr>
			</table>	
		</td>
	</tr>
</table>	
<!--Fin Seccion 5-->
</form>
</body>
</html>
<% else %>
	<html>
	<head>
	<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
	<style type="text/css">
	.titleStyle {
		font-weight: bold;
		font-size: 20px;
	}
	</style>
	</head>
	<body>
	<% call GF_TITULO2("kogge64.gif","CARGA DE AFE") %>
	<div id="toolbar"></div>
	<br>
	<div align="center"><h4><% =GF_TRADUCIR("Para poder cargar un AFE debe primero carga la cotización elegida para el pedido")	%>.</h4></div>
	<div align="center"><h4><% =GF_TRADUCIR("Complete este procedimiento haciendo")	%> <a href="comprasPIC.asp?idPedido=<% =idPedido %>">click</a> <% =GF_TRADUCIR("aqui")%>.</h4></div>
	</body>
	</html>
<% end if %>
