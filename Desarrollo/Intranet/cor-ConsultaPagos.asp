<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientospaginacion.asp"-->
<!--#include file="Includes/procedimientosAS400.asp"-->
<!--#Include File="Includes/ExternalFunctions.ASP" -->
<!--#include file="Includes/cor-IncludeCTO.asp"-->
<!--#include file="Includes/cor-IncludePC.asp"-->
<%
Dim strErrorMsg,p_intDia,p_intMes,p_intAnio,strDS
Dim strTitulo,strLinkPagina,intMostrar, intIndex, unitDest
Dim cmbProducto, intProducto, aniosql, diasql, messql, retValue
Dim accion, rsTipo, oConn
Dim rs, conn, strCampoOrden

accion = GF_Parametros7("accion","",6)

strErrorMsg= ""

%>
<html>
<head>
	<title>TOEPFER INTERNATIONAL - Impresion de Retenciones</title>
	<link href="CSS/ActisaIntra-1.css" rel="stylesheet" type="text/css">
	<script language="javascript" src="Scripts/script_fechas.js"></script>
	<script language="javascript">
		function buscar() {
			var KCPRO = document.getElementById('KCPRO').value;
			var Minuta = document.getElementById('minuta').value;
			var dia = document.getElementById('dia').value;
			var mes = document.getElementById('mes').value;
			var anio = document.getElementById('anio').value;
			var slcComprobante = document.getElementById('cmbKCComprobante');
			var kcComprobante = slcComprobante.options[slcComprobante.selectedIndex].value;
			var nroOrden = document.getElementById('NroOrden').value;
			var nroRet = document.getElementById('nroRet').value;
			var slcTipoRet = document.getElementById('cmbTipoRet');
			var kcTipoRet = slcTipoRet.options[slcTipoRet.selectedIndex].value;
			var params = '?KCPRO=' + KCPRO + '&minuta=' + Minuta;
				params = params + '&dia=' + dia + '&mes=' + mes + '&anio=' + anio;
				params = params + '&tipoCbte=' + kcComprobante + '&nroOrden=' + nroOrden;
				params = params + '&nroRet=' + nroRet + '&tipoRet=' + kcTipoRet;
			document.getElementById('ifrmResultados').src = 'cor-ResultadosPagos.asp' + params;
		}

		function expandirIFrame(p_alto) {
			document.getElementById('ifrmResultados').height = p_alto + 'px';
		}

		function administrarTeclas() {
			if (document.all){
				if (window.event.keyCode == 13)
					buscar();
			} else {
				if (window.event.which == 13)
					buscar();
			}

		}
	</script>
</head>
<body>
<%	call GF_TITULO2("kogge64.gif","Sistema Web - Impresión de Retenciones")	%>
	<input type="hidden" name="strTipo" value="<%=strTipo%>">
	<table width="60%" cellspacing="0" cellpadding="0" align="center" border="0">
		<input type="hidden" name="accion" id="accion" value="">
		<tr>
			<td width="8"><img src="images/marco_r1_c1.gif"></td>
			<td width="25%"><img src="images/marco_r1_c2.gif" width="100%" height="8"></td>
			<td width="8"><img src="images/marco_r1_c3.gif"></td>
			<td width="73%"><td>
			<td></td>
			</tr>
		<tr>
			<td width="8"><img src="images/marco_r2_c1.gif"></td>
			<td align=center valign="center"><font class="big" color="#517b4a"><% =GF_Traducir("Busqueda")%></font></td>
			<td width="8"><img src="images/marco_r2_c3.gif"></td>
			<td></td>
			<td></td>
			</tr>
		<tr>
			<td><img src="images/marco_r2_c1.gif" height="8"  width="8"></td>
			<td></td>
			<td valign="top" align="right"><img src="images/marco_r1_c2.gif" height="8" width="2"></td>
			<td><img src="images/marco_r1_c2.gif" width="100%" height="8"></td>
			<td width="8"><img src="images/marco_r1_c3.gif"></td>
			</tr>
		<tr>
			<td height="100%"><img src="images/marco_r2_c1.gif" height="100%" width="8"></td>
			<td colspan="3">
				<table width="100%" align="center" border="0">
				<tr>
					<td align="right" width="30%"><% =GF_Traducir("Cod. Proveedor")%>:</td>
					<td><input type="text" id="KCPRO" size="5" maxLength="5" value="" name="txtKcPro"></td>
				</tr>
				<tr>
					<td align="right" width="20%"><% =GF_Traducir("Minuta")%>:</td>
					<td><input type="text" id="minuta" size="6" maxLength="6" value="" name="txtMinuta"></td>
					<td align="right" width="20%"><% =GF_Traducir("Tipo Cbte.")%>:</td>
					<td>
<%						strSQL="Select * from MG where MG_KM='SF' order by MG_DS asc"
						call GF_BD_CONTROL(rs,oConn,"OPEN",strSQL)
%>						<select id="cmbKCComprobante" name="cmbKCComprobante">
							<option SELECTED value="">- <% =GF_TRADUCIR("Todos") %> -
<%							while (not rs.eof)	%>
								<option value="<% =rs("MG_KC") %>"><% =GF_TRADUCIR(rs("MG_DS")) %>
<%								rs.MoveNext
							wend
%>						</select>
					</td>
				</tr>
				<tr>
					<td align="right" width="20%"><% =GF_Traducir("Nro. de Orden")%>:</td>
					<td><input type="text" id="NroOrden" size="4" maxLength="4" value="" name="txtNroOrden"></td>
					<td align="right"><% =GF_Traducir("Fecha Pago")%>:</td>
					<td>
						<input type="text" id="dia"  size="2" maxLength="2" value="<% =p_intDia %>" name="txtDia"  onBlur="javascript:ControlarDia(this);"> /
						<input type="text" id="mes"  size="2" maxLength="2" value="<% =p_intMes %>" name="txtMes"  onBlur="javascript:ControlarMes(this);"> /
						<input type="text" id="anio" size="4" maxLength="4" value="<% =p_intAnio%>" name="txtAnio" onBlur="javascript:ControlarAnio(this);">
					</td>
				</tr>
				<tr>
                    <td align="right" width="20%"><%=GF_Traducir("Nro. Ret:")%></td>
                    <td><input type="text" size="10" value="<%=nroRet%>" id="nroRet" name="nroRet"></td>
					<td align="right" width="20%"><% =GF_Traducir("Tipo de Retención")%>:</td>
					<td>
						<select id="cmbTipoRet" name="cmbTipoRet">
							<option value="0" SELECTED>- <% =GF_TRADUCIR("Todos") %> -
<%							strSQL="Select * from tblconceptopago where CDCONCEPTO not in ('W','100', '101', '102')"
							call GF_BD_AS400(rsTipo,conn,"OPEN",strSQL)
							while (not rsTipo.eof)
								if (left(rsTipo("CDCONCEPTO"),1) = p_tipoRet) then
%>									<option value="<% =left(rsTipo("CDCONCEPTO"),1) %>" SELECTED><% =GF_TRADUCIR(rsTipo("DSCONCEPTO")) %>
<%								else	%>
									<option value="<% =left(rsTipo("CDCONCEPTO"),1) %>"><% =GF_TRADUCIR(rsTipo("DSCONCEPTO")) %>
<%								end if
								rsTipo.MoveNext
							wend
%>						</select>
					</td>
				</tr>
				<tr>
					<td colspan="4" align="center"><br><input type="button" value="<% =GF_Traducir("Buscar")%>..." onClick="buscar();" id=button1 name=button1></td>
				</tr>
			</table>
		</td>
			<td height="100%"><img src="images/marco_r2_c3.gif" width="8" height="100%"></td>
		</tr>
		<tr>
			<td width="8"><img src="images/marco_r3_c1.gif"></td>
			<td width="100%" align=center colspan="3"><img src="images/marco_r3_c2.gif" width="100%" height="8"></td>
			<td width="8"><img src="images/marco_r3_c3.gif"></td>
		</tr>
	</table>
	<div id="divAdobe" style="visibility:hidden;position:absolute;"><img src="images/get_adobe_reader.gif" onClick="javascript:window.open('http://www.adobe.com/products/acrobat/readstep2.html');" style="cursor:hand;"></div>
	<br>
<%	if (strErrorMsg = "") then	%>
		<iframe id="ifrmResultados" src="" frameborder=0 width="100%"></iframe>
<%	else	%>
		<table width="60%" cellspacing="0" cellpadding="0" align="center" border="0">
			<tr>
				<td width="8"><img src="images/marco_r1_c1.gif"></td>
				<td width="100%"><img src="images/marco_r1_c2.gif" width="100%" height="8"></td>
				<td width="8"><img src="images/marco_r1_c3.gif"></td>
			</tr>
			<tr>
				<td width="8" height="100%"><img src="images/marco_r2_c1.gif" width="8" height="100%"></td>
				<td align="center" class="TDERROR"><% =GF_TRADUCIR(strErrorMsg) %></td>
				<td width="8" height="100%"><img src="images/marco_r2_c3.gif" width="8" height="100%"></td>
			</tr>
			<tr>
				<td width="8"><img src="images/marco_r3_c1.gif"></td>
				<td width="100%"><img src="images/marco_r3_c2.gif" width="100%" height="8"></td>
				<td width="8"><img src="images/marco_r3_c3.gif"></td>
			</tr>
		</table>
		<script language="javascript">check_qnt=<% =intIndex%>;</script>
<%	end if	%>
</BODY>
<script language="javascript">
	var divAdobe = document.getElementById('divAdobe');
	divAdobe.style.top = '160px';
	divAdobe.style.left = window.screen.width - 300 + 'px';
	divAdobe.style.visibility = 'visible';
</script>
</HTML>

