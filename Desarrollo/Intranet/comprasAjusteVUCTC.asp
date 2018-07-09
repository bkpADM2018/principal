<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<%
'***********************************************************************************
'*******	                     COMIENZO DE LA PAGINA                      ********
'***********************************************************************************

Dim idContrato, idArticulo, accion, newTotal, saldoPesos, saldoDolares, ajuPesos, ajuDolares, newTotalPesos, newTotalDoalres
Dim guardado, aprobado, idArea, idDetalle, abrevUnidad, ajuCantidad, saldoCantidad, estado, ajusteImporte, saldo, aju
Dim myObraIdDivision, myObraDsDivision, member2Cd, CTC_ImporteAsignado
Dim myImporteCTC

guardado = false
idContrato = GF_Parametros7("idContrato",0,6)
accion = GF_PARAMETROS7("accion","",6)
newTotal = GF_PARAMETROS7("newTotal",2,6)
CTC_observaciones = GF_PARAMETROS7("CTC_observaciones","",6)
member2Cd = GF_PARAMETROS7("member2Cd","",6)

'Se controla que se haya indicado un contrato y que este existe.
if idContrato <> 0 then
	Set rsContrato = readCTC(idContrato)
	CTC_cdMoneda = rsContrato("CDMONEDA")
	if (rsContrato.eof) then Response.Redirect "comprasAccesoDenegado.asp"
else
	Response.Redirect "comprasAccesoDenegado.asp"
end if

CTC_cdResponsable = rsContrato("CDRESPONSABLE")
CTC_ContratoPesos	= Cdbl(rsContrato("IMPORTEUNITARIOPESOS"))
CTC_ContratoDolares = Cdbl(rsContrato("IMPORTEUNITARIODOLARES"))
myImporteCTC = CTC_ContratoPesos
if (CTC_cdMoneda = MONEDA_DOLAR) then myImporteCTC = CTC_ContratoDolares
CTC_tipoCambio = getTipoCambio(MONEDA_DOLAR,"")
CTC_idDivision = rsContrato("IDDIVISION")
'El importe ajustado estará en la moneda en la cual se nomine el contrato!!!. En la pantalla se mostrará solo uno de los campos, segun corresponda.
if ((newTotal = 0) and (accion = "")) then
	newTotal = myImporteCTC/100
end if

ajusteImporte = (newTotal*100) - myImporteCTC
saldo = myImporteCTC- CTC_ImporteAsignado + ajusteImporte


if Controlar then
	if accion = ACCION_GRABAR then
		guardado = true
		ajuPesos = ajusteImporte
		ajuDolares = Round(ajusteImporte / CTC_tipoCambio, 0)
		newTotalPesos = Round(newTotal*100, 0)
		newTotalDolares = Round(newTotalPesos / CTC_tipoCambio, 0)
		if (CTC_cdMoneda = MONEDA_DOLAR) then 
		    ajuDolares = ajusteImporte
		    ajuPesos = Round(ajusteImporte * CTC_tipoCambio, 0)
		    newTotalDolares = Round(newTotal*100, 0)
		    newTotalPesos = Round(newTotalDolares * CTC_tipoCambio, 0)
		end if
		'GUARDAR <CABECERA> DEL AJUSTE
		strSQL = "SELECT * FROM TBLOBRACTCAJUSTES WHERE IDCONTRATO=" & idContrato & " AND APLICADO='" & TIPO_NEGACION & "' AND TIPOAJUSTE='" & CTC_AJUSTE_UNITARIO & "'"  				
		Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if not rs.eof then
			myIdAjuste = rs("IDAJUSTE")			
			strSQL = "UPDATE TBLOBRACTCAJUSTES SET APLICADO='" & TIPO_NEGACION & "', IMPORTEPESOS=" & ajuPesos & ", IMPORTEDOLARES=" & ajuDolares & ", OBSERVACIONES='" & CTC_observaciones & "', CDUSUARIO='" & session("usuario") & "', MOMENTO=" & session("MmtoSistema") & " WHERE IDAJUSTE=" & myIdAjuste			
			Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
		else
			strSQL = "INSERT INTO TBLOBRACTCAJUSTES(IDCONTRATO, IMPORTEPESOS, IMPORTEDOLARES, OBSERVACIONES, APLICADO, CDUSUARIO, MOMENTO, TIPOAJUSTE) VALUES(" & idContrato & "," & ajuPesos & "," & ajuDolares & ", '" & CTC_observaciones & "','" & TIPO_NEGACION & "','" & session("usuario") & "'," & session("MmtoSistema") & ", '" & CTC_AJUSTE_UNITARIO & "')"				
			Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXECUTE", strSQL)
			strSQL = "SELECT * FROM TBLOBRACTCAJUSTES WHERE IDCONTRATO=" & idContrato & " AND APLICADO='" & TIPO_NEGACION & "' AND TIPOAJUSTE='" & CTC_AJUSTE_UNITARIO & "'" 
			Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
			myIdAjuste = rs("IDAJUSTE")
		end if	
		
		'Response.Write strSQL

		'GUARDAR <FIRMAS> DEL AJUSTE
		'Si el precio del contrato aumento, se debe autorizar el ajuste.
		'Si se esta disminuyendo el importe del contrato, se da por aprobado automaticamente el ajuste.		
		if (CDbl(ajusteImporte) < 0)  then					            
            myEstado = ESTADO_CTC_AUTORIZADO
            if (saldo <= 0) then myEstado = ESTADO_CTC_FINALIZADO
            'Se actualiza el estado del contrato.
			strSQL="Update TBLOBRACONTRATOS set IMPORTEUNITARIOPESOS=" & newTotalPesos & ", IMPORTEUNITARIODOLARES=" & newTotalDolares & ", ESTADO=" & myEstado & " where IDCONTRATO=" & idContrato
			Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
			
			'ACTUALIZAR ESTADO DE AJUSTE
			strSQL = "UPDATE TBLOBRACTCAJUSTES SET APLICADO='" & TIPO_AFIRMACION & "' WHERE IDAJUSTE=" & myIdAjuste
			'Response.Write strSQLAux
            Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
			
			Call setInfo(CTZ_AJU_APROBADO)		
		else 			
			'EL Contrato PASA A ESTADO EN AJUSTE
			strSQL = "UPDATE TBLOBRACONTRATOS SET ESTADO=" & ESTADO_CTC_EN_AJUSTE & " WHERE IDCONTRATO=" & idContrato			
			Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
		    'Se cargan las firmas.
		    Call addAJUCTCFirmas(myIdAjuste, CTC_idDivision, rsContrato("CDRESPONSABLE"), member2Cd)		    		    
		    Call setWarning(CTZ_AJU_NO_APROBADO)
		end if
			        			
		
	end if		
end if
'----------------------------------------------------------------------------------------------------------------------------------
function Controlar
	if accion = "" then 
		accion = ACCION_CONTROLAR
	else	
		
		if  (cdbl(newTotal)*100 = cdbl(myImporteCTC)) then Call setError(CTZ_AJU_IMP_IGUALES)		
		
		if len(CTC_observaciones) < 1 then Call setError(COMENTARIO_REQUERIDO) 
		
		if (member2Cd = "") then Call setError(AUTORIZANTE_NO_EXISTE)
		
        'Controlo que la partida presupuestaria del contrato no este en proceso de reasignacion/ajuste
        'if (tieneReasignacionAjusteActivo(CTC_idObra, CTC_areaObra, CTC_detalleObra)) then Call setError(BUDGET_REASIGNACION_EN_PROCESO)
	end if
	if not hayError() then Controlar = true
end function
'----------------------------------------------------------------------------------------------------------------------------------
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title><% =GF_TRADUCIR("Sistema de Compras - Ajuste CTC") %></title>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/iwin.css" type="text/css">
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
<script type="text/javascript" src="scripts/calendar.js"></script>
<script type="text/javascript" src="scripts/formato.js"></script>
<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/calendar-1.js"></script>
<script type="text/javascript" src="scripts/iwin.js"></script>
<script type="text/javascript">		
	var refpopupAjuCTC;
	function bodyOnLoad(){
		var tb = new Toolbar('toolbar', 4, 'images/compras/');
		<% if not guardado then %>
		idBtnGuardar = tb.addButtonSAVE("Guardar", "submitInfo('<% =ACCION_GRABAR %>')");
		idBtnControl = tb.addButtonCONFIRM("Controlar",  "submitInfo('<% =ACCION_CONTROLAR %>')");			
		<% end if %>	
		tb.draw();
	}
	
	function seleccionAutorizante() {
        if (document.getElementById("cmbUsrAut")) {
	        var e = document.getElementById("cmbUsrAut");
            document.getElementById("member2Cd").value = e.options[e.selectedIndex].value;
        }            
    }
	    
	function submitInfo(acc){
		document.getElementById("accion").value = acc;
		document.getElementById("frmSel").submit();
	}		
	function sumarTotal() {				
		var objP = document.getElementById("newTotal");	
		objP.value = editarImporte(objP.value);
		if (objP.value == 0) objP.value = "";
	}	
</script>
</head>
<body onLoad="bodyOnLoad()">
<form method="post" id="frmSel" action="comprasAjusteVUCTC.asp?idContrato=<%=idContrato%>">
<div id="toolbar"></div><br>
<table class="reg_header" align="center" width="95%" border="0">				
	<tr>
		<td colspan="3"><% call showErrors() %></td>
	</tr>
	<tr>
		<td align="right" class="numberStyle" colspan="3"><% =GF_TRADUCIR("Contrato:") %>&nbsp;<% =rsContrato("CDCONTRATO") %></td>				
	</tr>
	<tr>
		<td class="reg_header_nav recuadroRound" colspan="3"><% =GF_TRADUCIR("Datos del Contrato") %></td>
	</tr>
	<tr>
		<td></td>						
		<td align="center" width="20%" id="tituloPesos"><b><u><% =GF_TRADUCIR(getNombreMoneda(CTC_cdMoneda))   %></u></b></td>
	</tr>
	<tr>
		<td class="reg_header_nav recuadroRound"><% =GF_TRADUCIR("Valor Unitario del Contrato") %></td>				
		<td align="right"><font size="+1"><b><%=getSimboloMoneda(CTC_cdMoneda) & " " & GF_EDIT_DECIMALS(myImporteCTC,2)%></b></font></td>			
	</tr>
	<tr>
		<td COLSPAN="3" align="right"><HR></td>	
	</tr>			
	<tr>
		<td class="reg_header_nav recuadroRound"><% =GF_TRADUCIR("Ajuste") %></td>						
		<td align="right"><font size="+1"><b><%=getSimboloMoneda(CTC_cdMoneda) & " " & GF_EDIT_DECIMALS(ajusteImporte,2)%></b></font></td>
	</tr>
	<tr>
		<td class="reg_headeAr_nav recuadroRound"></td>				
		<td align="right"><HR></td>	
		<td align="right"><HR></td>
	</tr>
	<%if not guardado then %>
	<tr>
		<td class="reg_header_nav recuadroRound"><% =GF_TRADUCIR("Nuevo Valor Unitario") %></td>				
		<td align="right"><input style="text-align:right;" type="text" size="25" onBlur="sumarTotal()" name="newTotal" id="newTotal" onkeypress="return controlIngreso(this, event, 'I');" value="<%=newTotal%>"></td>		
	</tr>
	<tr>
		<td COLSPAN="3" align="right"><HR></td>	
	</tr>
	<%end if %>
	<tr>
		<td class="reg_header_nav recuadroRound" colspan="3"><% =GF_TRADUCIR("Justificación del Ajuste") %></td>				
	</tr>
	<tr>		
			<%if guardado then			
				Response.Write "<td colspan='3' align=left>" & CTC_observaciones
			else%>	
				<td colspan="3" align=center>
					<textarea name="CTC_observaciones" id="CTC_observaciones" cols="100"><%=CTC_observaciones%></textarea>				
			<%end if%>	
		</td>
	</tr>
	<tr>
		<td class="reg_header_nav" colspan="4"><% =GF_TRADUCIR("Firmantes") %></td>
	</tr>
	<tr>
		<td align="center">
		    <table border="0" cellpadding="0" cellspacing="0" width="100%">
		        <tr>
		            <td width="50%" align="center">Responsable</td>
		            <td width="50%" align="center">Autorizante</td>
		        </tr>
		        <tr>
		            <td width="50%" align="center"><b><% =getUserDescription(CTC_cdResponsable) %></b></td>
		            <td width="50%" align="center"><% Call dibujarComboGte(CTC_cdResponsable, member2Cd) %></td>
		        </tr>
		    </table>					
		</td>
	</tr>
</table>
<input type="hidden" name="accion" id="accion">
<input type="hidden" id="member2Cd" name="member2Cd" value="<%=member2Cd%>">
</form>
</body>
</html>