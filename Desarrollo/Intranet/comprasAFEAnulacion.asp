<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientossql.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosAFE.asp"-->
<!-- #include file="Includes/procedimientosUser.asp"-->
<%
Call comprasControlAccesoCM(RES_AFE)

'-----------------------------------------------------------------------------------------------
Function getEstadoAFE(estado)
	Dim aux
	Select case (estado)
		case AFE_NO_CONFIRMADO:
			aux = "No Confirmado"
		case AFE_APROBADO:
			aux = "Aprobado"
		case AFE_ANULADO:
			aux = "Anulado"
		case AFE_ANULACION:
			aux = "Anulación"
        case else
			aux = "En Firma"  
	end Select
	getEstadoAFE = aux
End Function
'-----------------------------------------------------------------------------------------------
Function anularAFE(idAFE, pComentario)
	Dim strSQL, conn, rsAFE, auxIsAnulacion, auxIdAFE, auxCdAFE, rsAFECompl, AFEsAnulados, hayCompl, complAnulados
	
	'Se anulan los complementarios.
	hayCompl = false

	if (afe_NroAFEComplID = 0) then
		Set rsAFECompl = listaAFESComplementarios(idAFE)
		if (not rsAFECompl.eof) then
			While (not rsAFECompl.eof)
				strSQL = "UPDATE tbldatosafe SET confirmado = '" & AFE_ANULADO & "' WHERE IDAFE = " & rsAFECompl("IDAFE")
				Call executeQueryDb(DBSITE_SQL_INTRA, rsAFE "UPDATE", strSQL)
				complAnulados = complAnulados & rsAFECompl("CDAFE") & ", "
				hayCompl = true
				rsAFECompl.MoveNext
			Wend
		end if
	end if

	'Se crea el nuevo AFE anulación
	AFEsAnulados = "AFE ANULADO " & afe_CdAFE
	if (hayCompl) then
		AFEsAnulados = AFEsAnulados & " Y SUS COMPLEMENTARIOS " & complAnulados
		AFEsAnulados = left(AFEsAnulados, len(AFEsAnulados)-2)
	end if

	nroAFEAnula = idAFE
	afe_Descripcion = AFEsAnulados & "  -  " & pComentario
    Call addAFE(auxIdAFE, auxCdAFE, afe_IdObra, afe_IdPedido, afe_ObraCuentaDS, afe_NroAFEComplID, afe_Titulo, afe_IdDivision, afe_Departamento, afe_Categoria, afe_CatOtros, afe_Tipo, afe_TipoOtros, afe_TipoCC, afe_Descripcion, afe_ImportePesos, afe_ImporteDolares, afe_TipoCambio, afe_NPV, afe_IRR, afe_ROIC, afe_PAYBACK, afe_PreparedByCD, afe_RequestedByCD, afe_EngReviewCD, afe_isCFO, afe_IDArea,afe_IDDetalle,nroAFEAnula)
	
	strSQL = "UPDATE tbldatosafe SET confirmado = '" & AFE_ANULACION & "' WHERE IDAFE = " & auxIdAFE
	Call executeQueryDb(DBSITE_SQL_INTRA, rsAFE "UPDATE", strSQL)

	'Se da de baja el AFE.
	strSQL = "UPDATE tbldatosafe SET confirmado = '" & AFE_ANULADO & "' WHERE IDAFE = " & idAFE
	Call executeQueryDb(DBSITE_SQL_INTRA, rsAFE "UPDATE", strSQL)
	
End Function
'-----------------------------------------------------------------------------------------------
'**********************************************************
'*****************	COMIENZO DE PAGINA	*******************
'**********************************************************
Dim idAFE, accion, comentario, dsUsuario, rs

idAFE = GF_PARAMETROS7("idAFE", 0, 6)
accion = GF_PARAMETROS7("accion", "", 6)
comentario = GF_PARAMETROS7("comentario", "", 6)

Call readAFE(idAFE, 0, 0)

if (accion = ACCION_GRABAR) then
	Call anularAFE(idAFE, comentario)
	accion = ACCION_CERRAR
end if

%>
<html>
<head>
<title>Anulación de AFEs</title>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">

<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript">
	var popUpAnulacion = getObjPopUp('popUpAnularAFE');
	function bodyOnLoad() {
		<% if (accion = ACCION_CERRAR) then %>
			cerrar();
		<% end if %>
		var tb = new Toolbar('toolbar', 5, 'images/almacenes/');
		tb.addButton("accept-16x16.png", "Confirmar", "submitInfo()");
		tb.draw();
		document.getElementById("comentario").focus();
	}

	function cerrar() {
		popUpAnulacion.hide();
	}

	function submitInfo() {
		if (document.getElementById("comentario").value == '') {
			alert('Debe ingresar el motivo de la anulación del AFE');
		} else {
			document.getElementById("accion").value = '<% =ACCION_GRABAR %>';
			document.getElementById("frmAFE").submit();
		}
	}
</script>
</head>
<body onLoad="bodyOnLoad()">
<div id="toolbar"></div>
<br>
<form id="frmAFE" name="frmAFE">
	<table class="reg_Header" align="center" width="90%">
		<tr>
			<td align="right" colspan="4" style="font-size:14px; font-weight:bold;"><% =GF_TRADUCIR("AFE: ") & afe_CdAFE  %></td>
		</tr>
		<tr>
			<td colspan="4" class="reg_Header_nav"><% =GF_TRADUCIR("Datos del AFE") %></td>
		</tr>
		<tr>
			<td class="reg_Header_navdos" width="15%"><% =GF_TRADUCIR("Titulo")  %></td>
			<td colspan="3"><% =afe_ObraDS  %></td>
		</tr>
		<tr>
			<td class="reg_Header_navdos" width="15%"><% =GF_TRADUCIR("División")  %></td>
			<td width="30%"><% =afe_ObraDivDS  %></td>
			<td class="reg_Header_navdos" width="15%"><% =GF_TRADUCIR("Responsable")  %></td>
			<% dsUsuario = getUserDescription(afe_cdUsuario) %>
			<td><% =dsUsuario  %></td>
		</tr>
		<tr>
			<td class="reg_Header_navdos" width="15%"><% =GF_TRADUCIR("Importe")  %></td>
			<td><% =getSimboloMoneda(MONEDA_DOLAR) & " " & GF_EDIT_DECIMALS(afe_ImporteDolares,2)  %></td>
			<td class="reg_Header_navdos" width="15%"><% =GF_TRADUCIR("Estado")  %></td>
			<td><% =getEstadoAFE(afe_Confirmado)  %></td>
		</tr>
		<tr>
			<td colspan="4" class="reg_Header_nav"><% =GF_TRADUCIR("AFEs Complementarios") %></td>
		</tr>
		<tr>
			<td colspan="4">
				<table width="100%">
				<%	if(afe_NroAFEComplID > 0) then %>
						<tr><td><% =GF_TRADUCIR("Este AFE es complementario del AFE: ") & getCdAFE(afe_NroAFEComplID) %></td></tr>
				<%	else %>
				<%		Set rs = listaAFESComplementarios(idAFE) %>
				<%		if (not rs.eof) then %>
							<tr><td><% =GF_TRADUCIR("Este AFE contiene uno o mas AFEs complementarios que se anularan automaticamente junto con el mismo.") %></td></tr>
				<%			While (not rs.eof) %>
								<tr><td><% =rs("CDAFE") %></td></tr>
				<%				rs.MoveNext %>
				<%			Wend %>
							</td>
				<%		else %>
							<tr><td><% =GF_TRADUCIR("No se encontraron AFEs complementarios.") %></td></tr>
				<%		end if %>
				<%	end if %>
				</table>
			</td>
		</tr>
		<tr>
			<td colspan="4" class="reg_Header_nav"><% =GF_TRADUCIR("Comentario") %></td>
		</tr>
		<tr>
			<td colspan="4" style="color:red;" align="center"><% =GF_TRADUCIR("* Ingrese el motivo de anulación del AFE") %></td>
		</tr>
		<tr>
			<td colspan="4" align="center">
				<textarea id="comentario" name="comentario" cols="50" rows="4" maxlength="3000" value="<% =comentario %>"></textarea>
			</td>
		</tr>
	</table>
	<input type="hidden"  id="idAFE" name="idAFE" value="<% =idAFE %>">
	<input type="hidden"  id="cdAFE" name="cdAFE" value="<% =afe_CdAFE %>">
	<input type="hidden"  id="idObra" name="idObra" value="<% =afe_IdObra %>">
	<input type="hidden"  id="idPedido" name="idPedido" value="<% =afe_IdPedido %>">
	<input type="hidden"  id="nroAFEComplID" name="nroAFEComplID" value="<% =afe_NroAFEComplID %>">
	<input type="hidden"  id="idDivision" name="idDivision" value="<% =afe_IdDivision %>">
	<input type="hidden"  id="categoria" name="categoria" value="<% =afe_Categoria %>">
	<input type="hidden"  id="catOtros" name="catOtros" value="<% =afe_CatOtros %>">
	<input type="hidden"  id="tipo" name="tipo" value="<% =afe_Tipo %>">
	<input type="hidden"  id="tipoOtros" name="tipoOtros" value="<% =afe_TipoOtros %>">
	<input type="hidden"  id="cumplimientos" name="cumplimientos" value="<% =afe_TipoCC %>">
	<input type="hidden"  id="descripcion" name="descripcion" value="<% =afe_Descripcion %>">
	<input type="hidden"  id="importePesos" name="importePesos" value="<% =cDbl(afe_ImportePesos)/100 %>">
	<input type="hidden"  id="importeDolares" name="importeDolares" value="<% =cDbl(afe_ImporteDolares)/100 %>">
	<input type="hidden"  id="tipoCambio" name="tipoCambio" value="<% =afe_TipoCambio %>">
	<input type="hidden"  id="Arr" name="Arr" value="<% =cDbl(afe_Arr)/100 %>">
	<input type="hidden"  id="Irr" name="Irr" value="<% =cDbl(afe_Irr)/100 %>">
	<input type="hidden"  id="Q1" name="Q1" value="<% =afe_Preg1 %>">
	<input type="hidden"  id="Q2" name="Q2" value="<% =afe_Preg2 %>">
	<input type="hidden"  id="Q3" name="Q3" value="<% =afe_Preg3 %>">
	<input type="hidden"  id="Q4" name="Q4" value="<% =afe_Preg4 %>">
	<input type="hidden"  id="Q4T" name="Q4T" value="<% =afe_Preg4T %>">
	<input type="hidden"  id="Q5" name="Q5" value="<% =afe_Preg5 %>">
	<input type="hidden"  id="Q6" name="Q6" value="<% =afe_Preg6 %>">
	<input type="hidden"  id="Q7" name="Q7" value="<% =afe_Preg7 %>">
	<input type="hidden"  id="Q8" name="Q8" value="<% =afe_Preg8 %>">
	<input type="hidden"  id="Q9" name="Q9" value="<% =afe_Preg9 %>">
	<input type="hidden"  id="Q10" name="Q10" value="<% =afe_Preg10 %>">
	<input type="hidden"  id="Q11" name="Q11" value="<% =afe_Preg11 %>">
	<input type="hidden"  id="confirmado" name="confirmado" value="<% =afe_Confirmado %>">
	<input type="hidden"  id="obraCuentaDS" name="obraCuentaDS" value="<% =afe_ObraCuentaDS %>">
	<input type="hidden"  id="idArea" name="idArea" value="<% =afe_IDArea %>">
	<input type="hidden"  id="idDetalle" name="idDetalle" value="<% =afe_IDDetalle %>">
	<input type="hidden"  id="RONA" name="RONA" value="<% =cDbl(afe_RONA)/100 %>">
	<input type="hidden"  id="PAYBACK" name="PAYBACK" value="<% =cDbl(afe_PAYBACK)/100 %>">
	<input type="hidden"  id="dsObra" name="dsObra" value="<% =afe_ObraDS %>">
	<input type="hidden"  id="preparedByCD" name="preparedByCD" value="<% =afe_PreparedByCD %>">
	<input type="hidden"  id="requestedByCD" name="requestedByCD" value="<% =afe_RequestedByCD %>">
	<input type="hidden"  id="engReviewCD" name="engReviewCD" value="<% =afe_EngReviewCD %>">
	<input type="hidden"  id="officerCD" name="officerCD" value="<% =afe_OfficerCD %>">
	<input type="hidden"  id="vicePresidentCD" name="vicePresidentCD" value="<% =afe_VicePresidentCD %>">
	<input type="hidden"  id="presidentCD" name="presidentCD" value="<% =afe_PresidentCD %>">
	<input type="hidden"  id="controllerCD" name="controllerCD" value="<% =afe_ControllerCD %>">
	<input type="hidden"  id="departamento" name="departamento" value="<% =afe_Departamento %>">
	<input type="hidden"  id="accion" name="accion" value="">
</form>
<br>
</body>
</html>