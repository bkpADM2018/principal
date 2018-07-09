<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosAS400.asp"-->
<!--#include file="Includes/procedimientosmail.asp"-->
<!--#include file="Includes/procedimientosCupos.asp"-->
<!--#include file="Includes/procedimientosUnificador.asp"-->
<%

Dim idProveedor, fecha
Dim nroContrato, cdProducto, cdSucursal, anioCosecha
Dim pdf_accion
Dim periodo, eMails

idProveedor = GF_PARAMETROS7("id",0,6)
usr = GF_PARAMETROS7("usr","",6)
accion = GF_PARAMETROS7("accion","",6)
fecha = GF_PARAMETROS7("fecha",0,6)
nroContrato = GF_Parametros7("nroContrato", 0, 6)
cdProducto = GF_Parametros7("cdProducto", 0, 6)
cdSucursal = GF_Parametros7("cdSucursal", 0, 6)
cdOperacion = GF_Parametros7("cdOperacion", 0, 6)
anioCosecha = GF_Parametros7("anioCosecha", 0, 6)
pdf_accion = GF_Parametros7("pdf_accion", 0, 6)
xls_accion = GF_Parametros7("xls_accion", 0, 6)
periodo = GF_Parametros7("periodo", 0, 6)
eMails = GF_PARAMETROS7("mails","",6)
descripcionMail = GF_PARAMETROS7("descripcionMail","",6)
if (descripcionMail = "") then descripcionMail="Ver cupos asignados en el archivo adjunto"

'filtros
fltrCdProducto = GF_PARAMETROS7("fltrCdProducto","",6)
fltrCdSucursal = GF_PARAMETROS7("fltrCdSucursal","",6)
fltrCdOperacion = GF_PARAMETROS7("fltrCdOperacion","",6)
fltrNroContrato = GF_PARAMETROS7("fltrNroContrato","",6)
fltrAnioCosecha = GF_PARAMETROS7("fltrAnioCosecha","",6)
fltrPuerto = GF_PARAMETROS7("fltrPuerto",0,6)
fltrCorredor = GF_PARAMETROS7("fltrCorredor",0,6)
fltrVendedor = GF_PARAMETROS7("fltrVendedor",0,6)

chkEnviados = GF_PARAMETROS7("chkEnviados",0,6)

'Se determina el proveedor verdadero para cuando es mercado a termino
if (idProveedor = MERCADO_A_TERMINO) then	
	strSQL="Select * from MERFL.MER311FH where CPRORH=" & cdProducto & " and CSUCRH=" & cdSucursal & " and COPERH=" & cdOperacion & " and NCTORH=" & nroContrato & " and ACOSRH=" & anioCosecha	
	Call executeQuery(rsCto, "OPEN", strSQL)
	if (not rsCto.eof) then idProveedor = rsCto("CCORRH")	
end if

if (accion = ACCION_GRABAR) then 	
	Call grabarMails(emails, idProveedor)
end if
%>
<html>
<head>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript" src="scripts/iwin.js"></script>

<script type="text/javascript">
var refPopUpEnviarMail;

function validarEmail(valor) {	
	var filter=/^([\w-]+(?:\.[\w-]+)*)@((?:[\w-]+\.)*\w[\w-]{0,66})\.([a-z]{2,6}(?:\.[a-z]{2})?)$/i		
	if (!filter.test(valor)){		
			alert("La dirección de email " + valor + " es incorrecta.\n" + "Ingrese los mails separados por punto y coma (;)");
			return false;
		}
	return true;		
}
	
function bodyOnLoad() {
    refPopUpEnviarMail = startIWin('popupEnviarMail');    
	<% if (CLng(idProveedor) = PROV_ID_ADM) then %>
	    adjuntarADM();
    <% else %>	    
        adjuntarXLS();
    <% end if %>	    

}

function adjuntarXLS(){
	document.getElementById("frmSel").action = 'cuposPorProveedorPrintXLS.asp';	
	document.getElementById("adjuntoXLS").checked = 'checked';
	document.getElementById("adjuntoPDF").checked = '';
	document.getElementById("adjuntoADM").checked = '';
}
function adjuntarPDF(){
	document.getElementById("frmSel").action = 'cuposPorProveedorPrint.asp';
	document.getElementById("adjuntoXLS").checked = '';
	document.getElementById("adjuntoPDF").checked = 'checked';
	document.getElementById("adjuntoADM").checked = '';
}
function adjuntarADM() {
    document.getElementById("frmSel").action = 'cuposPorProveedorPrintADM.asp';
    document.getElementById("adjuntoADM").checked = 'checked';
    document.getElementById("adjuntoPDF").checked = '';
    document.getElementById("adjuntoXLS").checked = '';
}
function agregarMail(){
	var strMails = document.getElementById('mails').value;
	strMails = strMails.replace(/\n/gi, "");
	var arrayMails = strMails.split(";");
	if (arrayMails.length < 10 ){
		var mailsCorrectos = true;
		for (i=0;i<arrayMails.length;i++){	
			if (!validarEmail(arrayMails[i].toString().toLowerCase())) {
				mailsCorrectos = false;
			}
		}
		if (mailsCorrectos){
			document.getElementById("frmSel").action = 'cuposGetDescripcionMail.asp';
			document.getElementById("accion").value= '<% =ACCION_GRABAR %>';
			document.getElementById("frmSel").submit();
		}
	}else{
		alert("solo se permite ingresar hasta diez e-mails por proveedor");
	}		
}

</script>
</head>
<body onLoad="bodyOnLoad()">
<form name="frmSel" id="frmSel" method="post" action="">
<table width="100%" align=center>
	<tr>
		<td class="title_sec_section"><img align="absMiddle" src="images/cupos/Mail-32x32.png"> <% =GF_TRADUCIR("Enviar Mail") & " a " & getDescripcionProveedor(idProveedor)%></td>
	</tr>
	<tr>
		<td>
			<table width="100%">				
				<tr><td colspan="3">
					<table class="reg_Header" id="TAB3" align="center" border="0">	
						<tr>
							<td class="reg_Header_navdos">
								<% =GF_TRADUCIR("Mails") %>
							</td>				
							<td align="center">
								<textarea id="mails" name="mails" style="text-align: left" wrap="soft" rows="2" cols="30"><% =getStringMailsProveedor(idProveedor) %></textarea>
							</td>
							<td align="center">
								<img src="images/cupos/Guardar.gif" onclick="javascript:agregarMail();"style="cursor:pointer" title="<%=GF_Traducir("Agregar Mail")%>"></img>
							</td>									
						</tr>					
					</table>
				</td></tr>
				<tr>
					<td width="10%" class="reg_header" align="center"><% =GF_TRADUCIR("Adjunto") & ":" %></td>
					<td width= "90%">
						<input type="radio" name="adjuntoXLS" id="adjuntoXLS" onclick="adjuntarXLS();" /><img src="images/cupos/excel.gif" />&nbsp; Excel
						<input type="radio" name="adjuntoPDF" id="adjuntoPDF" onclick="adjuntarPDF();" /><img src="images/cupos/pdf.gif" />&nbsp; PDF
						<input type="radio" name="adjuntoADM" id="adjuntoADM" onclick="adjuntarADM();" /><img src="images/cupos/document-16.png" />&nbsp; TXT
					</td>
				</tr>
				<tr>
					<td colspan="2" class="reg_header" align="center"><textarea id="descripcionMail" name="descripcionMail" style="text-align: left" wrap="soft" rows="10" cols="92"><%=descripcionMail%></textarea>
					</td> 
				</tr>											
				
			</table>
		</td>
	</tr>	
	<tr>
		<td align="center">
			<table>	
				<tr>
					<td>
						<input type="submit" id="aceptar" name="aceptar" value="<% =GF_TRADUCIR("Aceptar") %>" />
						<input type="button" value="<% =GF_TRADUCIR("Cancelar") %>" onClick="refPopUpEnviarMail.hide()" id="button1" name="button1" />
					</td>
				</tr>
			</table>
		</td>		
	</tr>
</table>
<input type="hidden" name="accion" id="accion" value="<% =ACCION_GRABAR %>">
<input type="hidden" name="id" id="id" value="<% =idProveedor %>">
<input type="hidden" name="fecha" id="fecha" value="<% =fecha %>">
<input type="hidden" name="usr" id="usr" value="<% =usr %>">
<input type="hidden" name="nroContrato" id="nroContrato" value="<% =nroContrato %>">
<input type="hidden" name="cdProducto" id="cdProducto" value="<% =cdProducto%>">
<input type="hidden" name="cdOperacion" id="cdOperacion" value="<% =cdOperacion %>">
<input type="hidden" name="cdSucursal" id="cdSucursal" value="<% =cdSucursal %>">
<input type="hidden" name="anioCosecha" id="anioCosecha" value="<% =anioCosecha %>">
<input type="hidden" name="fltrNroContrato" id="fltrNroContrato" value="<% =fltrNroContrato %>">
<input type="hidden" name="fltrCdProducto" id="fltrCdProducto" value="<% =fltrCdProducto%>">
<input type="hidden" name="fltrCdOperacion" id="fltrCdOperacion" value="<% =fltrCdOperacion %>">
<input type="hidden" name="fltrCdSucursal" id="fltrCdSucursal" value="<% =fltrCdSucursal %>">
<input type="hidden" name="fltrAnioCosecha" id="fltrAnioCosecha" value="<% =fltrAnioCosecha %>">
<input type="hidden" name="fltrPuerto" id="fltrPuerto" value="<% =fltrPuerto %>">
<input type="hidden" name="fltrCorredor" id="fltrCorredor" value="<% =fltrCorredor %>">
<input type="hidden" name="fltrVendedor" id="fltrVendedor" value="<% =fltrVendedor %>">
<input type="hidden" name="pdf_accion" id="pdf_accion" value="<% =pdf_accion%>">
<input type="hidden" name="xls_accion" id="xls_accion" value="<% =xls_accion%>">
<input type="hidden" name="chkEnviados" id="chkEnviados" value="<% =chkEnviados%>">
<input type="hidden" name="periodo" id="periodo" value="<% =periodo%>">
</form>
</body>
</html>