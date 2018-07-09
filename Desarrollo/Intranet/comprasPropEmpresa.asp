<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosAS400.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosSql.asp"-->
<!--#include file="Includes/procedimientosEmpresas.asp"-->
<%
Call comprasControlAccesoCM(RES_ADM)
'------------------------------------------------------------------------------------------
Function controlCUIT(cuit1, cuit2, cuit3)
	
	dim suma, cuitCompelto, digito, cuitDigits(11), x, resto
	
	controlCUIT = RESPUESTA_OK
	'Controlo las longitudes
	if (len(cuit1) <> 2) then controlCUIT = CUIT_ERRONEO
	if (len(cuit2) <> 8) then controlCUIT = CUIT_ERRONEO
	if (len(cuit3) <> 1) then controlCUIT = CUIT_ERRONEO	
	if (controlCUIT = RESPUESTA_OK) then
		'Se controla el CUIT
		cuitCompleto = cuit1 & cuit2 & cuit3				
		For x = 0 to 10
			cuitDigits(x) = Left(cuitCompleto,1)			
			cuitCompleto = Right(cuitCompleto, 10-x)			
		next
		suma = cuitDigits(0) * 5
		suma = suma + cuitDigits(1) * 4
		suma = suma + cuitDigits(2) * 3
		suma = suma + cuitDigits(3) * 2
		suma = suma + cuitDigits(4) * 7
		suma = suma + cuitDigits(5) * 6
		suma = suma + cuitDigits(6) * 5
		suma = suma + cuitDigits(7) * 4
		suma = suma + cuitDigits(8) * 3
		suma = suma + cuitDigits(9) * 2			
		resto = suma mod 11
		if (resto > 0) then
			digito = 11 - resto				
		else
			digito = 0
		end if				
		if (digito <> Cint(cuitDigits(10))) then controlCUIT = CUIT_ERRONEO
	end if
End Function
'------------------------------------------------------------------------------------------
Function controlEMAIL(email)
	
	controlEMAIL = RESPUESTA_OK
	if ((Instr(1, email, "@") = 0) and (email <> "")) then
		controlEMAIL = EMAIL_ERRONEO
	end if
	
End Function
'------------------------------------------------------------------------------------------
Function controlar(idEmpresa, dsEmpresa, cuit1, cuit2, cuit3, email, estado, tipoDoc)
	Dim strSQL, rs, conn, cuit
	
	cuit = cuit1 & cuit2 & cuit3
	controlar = controlCUIT(cuit1, cuit2, cuit3)
	if (controlar = RESPUESTA_OK) then
		controlar = controlEMAIL(email)
		if (controlar = RESPUESTA_OK) then
			if (tipoDoc <> TIPO_CUIT_EX_83) then
				strSQL="Select * from VWEMPRESAS where cuit=" & cuit & " and tipodocumento=" & tipoDoc
				Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
				if (not rs.eof) then
					'Compraro todas las empresas con igual CUIT para saber si la que
					'quiero modificar es una de ellas.
					esEmpresa = false
					while not rs.eof
						if (CLng(rs("idEmpresa")) = CLng(idEmpresa)) then esEmpresa = true
						'Si la empresa esta registrada en el AS400, es única y esta dada de baja se permite crear
						'la nueva dado que luego si resulta ganadora se reemplazará automaticamente por los datos 
						'ya guardados en el AS400.
						if (CLng(rs("idEmpresa")) < 100000) then
							if (rs("ESTADO") = "*") then esEmpresa = true
						end if
						rs.MoveNext()
					wend
					if (not esEmpresa) then
						controlar = EMPRESA_EXISTE
					end if
				end if
			end if
		end if
	end if
	if (dsEmpresa = "") then controlar = DESCRIPCION_VACIA
	if ((estado < PROV_ACTIVO) or (estado > PROV_PROHIBIDO_PAGOS)) then controlar = VALOR_NO_VALIDO

End Function
'------------------------------------------------------------------------------------------
Function accionGrabarEMAIL(idEmpresa, email)
	Dim strSQL, rs, conn
	
	strSQL="Select * from TBLMAILSCOMPRAS where IDEMPRESA=" & idEmpresa
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (rs.eof) then
		'Es una unidad nueva
		strSQL="Insert into TBLMAILSCOMPRAS(IDEMPRESA, EMAIL)"		
		strSQL= strSQL & " values(" & idEmpresa & ", '" & email & "')"
	else
		'Es una modificacion
		strSQL="Update TBLMAILSCOMPRAS Set EMAIL='" & email & "' where IDEMPRESA=" & idEmpresa
	end if
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
End Function
'------------------------------------------------------------------------------------------
Function accionGrabarEstado(cuitEmpresa, estado, motivo)
	Dim strSQL, rs, conn

	'grabo el estado en la lista negra, si es mayor a 0 (ACTIVA)
	strSQL="Select * from TBLESTADOEMPRESAS where CUIT=" & cuitEmpresa
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (estado = PROV_ACTIVO) then
		if (not rs.eof) then
			strSQL = "Delete from TBLESTADOEMPRESAS "
			strSQL = strSQL & " where CUIT=" & cuitEmpresa
			Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
		end if
	else		
		if (not rs.eof) then
			strSQL = "Update TBLESTADOEMPRESAS Set ESTADO=" & estado & ", MOTIVO='" & motivo
			strSQL = strSQL & "' where CUIT=" & cuitEmpresa	
		else
			strSQL = "Insert into TBLESTADOEMPRESAS(CUIT, ESTADO, MOTIVO) "
			strSQL = strSQL & "values (" & cuitEmpresa & ", " & estado & ", '" & motivo & "')"
		end if
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
	end if

End Function
'------------------------------------------------------------------------------------------
Function accionGrabar(idEmpresa, dsEmpresa, cuit1, cuit2, cuit3, email, estado, tipoDoc, motivo)
	Dim strSQL, rs, conn, cuitEmpresa
	
	cuitEmpresa = cuit1 & cuit2 & cuit3
	
	accionGrabar = false
	if (idEmpresa = 0) then
		'Es una unidad nueva
		strSQL="Insert into TBLEMPRESAS(DSEmpresa, CUIT, TIPODOCUMENTO)"
		strSQL= strSQL & " values('" & dsEmpresa & "', " & cuitEmpresa & ", " & tipoDoc & ")"
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
		strSQL = "Select MAX(IDEMPRESA) MAXIMO from TBLEMPRESAS"
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if (not rs.eof) then idEmpresa = rs("MAXIMO")
	else
		'Es una modificacion
		strSQL="Update TBLEMPRESAS Set DSEMPRESA='" & dsEmpresa & "', CUIT=" & cuitEmpresa & ", TIPODOCUMENTO=" & tipoDoc
		strSQL = strSQL & " where IDEMPRESA=" & idEmpresa
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
	end if	
	'grabo mail
	Call accionGrabarEMAIL(idEmpresa, email)
	'grabo estado
	Call accionGrabarEstado(cuitEmpresa, estado, motivo)
	accionGrabar = true
	
End Function
'------------------------------------------------------------------------------------------
Function accionConsulta(idEmpresa, ByRef dsEmpresa, ByRef cuit1, ByRef cuit2, ByRef cuit3, ByRef email, ByRef estado, ByRef tipoDoc, ByRef motivo)
	
	Dim strSQL, rs, conn, cuitTemp, cuit
	
	cuit = cuit1 & cuit2 & cuit3
	if (cuit = "") then cuit = getEnterpriseCUIT(idEmpresa)
	strSQL="Select * from VWEMPRESAS where IDEMPRESA=" & idEmpresa
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then				
		dsEmpresa = Trim(rs("DSEMPRESA"))
		cuitTemp = Trim(rs("CUIT"))
		cuit1 = left(cuitTemp, 2)
		cuit2 = mid(cuitTemp, 3, 8)
		cuit3 = right(cuitTemp, 1)
		tipoDoc = cDBl(rs("TIPODOCUMENTO"))
		strSQL="Select * from TBLMAILSCOMPRAS where IDEMPRESA=" & idEmpresa
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if (not rs.eof) then email=rs("EMAIL")
		strSQL="Select * from TBLESTADOEMPRESAS where CUIT=" & cuit
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if (not rs.eof) then 
			estado=rs("ESTADO")
			motivo=rs("MOTIVO")
		end if	
	end if
End Function
'***************************************************
'******   COMIENZO DE LA PAGINA
'***************************************************
Dim accion, errMsg, idEmpresa, dsEmpresa, cuit1, cuit2, cuit3, email, estado, motivo

idEmpresa = GF_PARAMETROS7("idEmpresa",0,6)
cuit1 = GF_PARAMETROS7("cuit1","",6)
cuit2 = GF_PARAMETROS7("cuit2","",6)
cuit3 = GF_PARAMETROS7("cuit3","",6)
dsEmpresa = UCase(GF_PARAMETROS7("descripcion","",6))
email = GF_PARAMETROS7("email","",6)
estado = GF_PARAMETROS7("estado", 0, 6) 'REFERIDO A LISTA NEGRA
tipoDoc = GF_PARAMETROS7("tipoDoc",0,6)
accion = GF_PARAMETROS7("accion","",6)
motivo = GF_PARAMETROS7("motivo","",6)

Call GP_ConfigurarMomentos
if (accion = ACCION_GRABAR) then
	errMsg = controlar(idEmpresa, dsEmpresa, cuit1, cuit2, cuit3, email, estado, tipoDoc)
	if (errMsg = RESPUESTA_OK) then
		Call accionGrabar(idEmpresa, dsEmpresa, cuit1, cuit2, cuit3, email, estado, tipoDoc, motivo)
		accion = ACCION_CERRAR
	else
		setError(errMsg)
	end if
else
	Call accionConsulta(idEmpresa, dsEmpresa, cuit1, cuit2, cuit3, email, estado, tipoDoc, motivo)
end if
if (accion = "") then accion = ACCION_GRABAR
 %>
<html>
<head>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css"	 type="text/css">

<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript" src="scripts/formato.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.5.1.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/botoneraPopUp.js"></script>
<script type="text/javascript">
var refPopUpEmpresa;

function checkComment() {
	var cmb = document.getElementById("motivo");
}

function codigoOnBlur(ref) {
	if (ref.value == "") {
		document.getElementById("aceptar").disabled = true;
	} else {
		document.getElementById("aceptar").disabled = false;
	}			
}

function submitir() {
	document.getElementById("frmSel").submit();
}

function EmpresaOnLoad() {
	var botones = new botonera("botones");
	
	<% if ((idEmpresa >= 100000) or (idEmpresa = 0)) then %>
		document.getElementById("descripcion").focus();
	<% else %>
		document.getElementById("email").focus();
	<% end if %>
	refPopUpEmpresa = getObjPopUp('popupEmpresa');
	<% if (accion = ACCION_CERRAR) then %>
		refPopUpEmpresa.hide();
	<% end if 	
	   if (not isAuditor(SIN_DIVISION)) then %>
		botones.addbutton("Aceptar","submitir()");
	<% end if%>
	botones.show();
}
</script>
</head>
<body onLoad="EmpresaOnLoad()">
<form name="frmSel" id="frmSel" method="post" action="comprasPropEmpresa.asp">
<table  width="100%">
	<tr>
		<td class="title_sec_section" colspan="2"><img align="absMiddle" src="images/compras/Company-32x32.png"> <% =GF_TRADUCIR("Propiedades de Empresa") %></td>
	</tr>
	<tr>
		<td colspan="2"><% call showErrors() %></td>
	</tr>
	<tr>
		<td></td>
		<td>
			<table width="100%" align="center">				
				<tr>
					<td width="10%" class="reg_header"><% =GF_TRADUCIR("Empresa") %></td>
					<td><% =idEmpresa %></td>
				</tr>
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("Descripcion") %></td>
					<td>
						<% if ((idEmpresa >= 100000) or (idEmpresa = 0)) then %>
						<input type="text" id="descripcion" name="descripcion" maxlength="50" size="30" value="<% =dsEmpresa %>" onkeypress="return controlSalto(this, event)">
						<% else %>
						<% =dsEmpresa %>
						<input type="hidden" id="descripcion" name="descripcion" value="<% =dsEmpresa %>">
						<% end if %>
					</td>
				</tr>
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("Tipo Doc.") %></td>
					<td>
						<!-- tipo de cuit, nacional o extranjero -->
						<select id="tipoDoc" name="tipoDoc">
							<option value="<% =TIPO_CUIT_80		%>" <% if (tipoDoc = TIPO_CUIT_80)		then Response.Write "selected" %>><% =GF_TRADUCIR("CUIT (80)")	 %></option>
							<option value="<% =TIPO_CUIT_EX_83	%>" <% if (tipoDoc = TIPO_CUIT_EX_83)	then Response.Write "selected" %>><% =GF_TRADUCIR("CUIT EX (83)")%></option>
						</select>
					</td>
				</tr>
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("CUIT") %></td>
					<td>
						<% if ((idEmpresa >= 100000) or (idEmpresa = 0)) then %>
						<input type="text" id="cuit1" name="cuit1" maxlength="2" size="2" value="<% =cuit1 %>" onKeyPress="return controlDatos(this, event, 'N')">-
						<input type="text" id="cuit2" name="cuit2" maxlength="8" size="8" value="<% =cuit2 %>" onKeyPress="return controlDatos(this, event, 'N')">-
						<input type="text" id="cuit3" name="cuit3" maxlength="1" size="1" value="<% =cuit3 %>" onKeyPress="return controlDatos(this, event, 'N')">
						<% else %>
						<% =cuit1 %>-<% =cuit2 %>-<% =cuit3 %>
						<input type="hidden" id="cuit1" name="cuit1" value="<% =cuit1 %>">
						<input type="hidden" id="cuit2" name="cuit2" value="<% =cuit2 %>">
						<input type="hidden" id="cuit3" name="cuit3" value="<% =cuit3 %>">
						<% end if %>
					</td>
				</tr>		
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("E-MAIL") %></td>
					<td>
						<input type="text" id="email" name="email" maxlength="256" size="30" value="<% =email %>">						
					</td>
				</tr>
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("Estado") %></td>
					<td>
						<!-- ESTADO DE LA EMPRESA EN LA LISTA NEGRA -->
						<select id="estado" name="estado">
							<option value="<% =PROV_ACTIVO %>" <% if (estado = PROV_ACTIVO) then Response.Write "selected" %>><% =GF_TRADUCIR("Activa") %></option>
							<option value="<% =PROV_PROHIBIDO_PEDIDOS %>" <% if (estado = PROV_PROHIBIDO_PEDIDOS) then Response.Write "selected" %>><% =GF_TRADUCIR("No puede recibir pedidos") %></option>
							<option value="<% =PROV_PROHIBIDO_PAGOS %>" <% if (estado = PROV_PROHIBIDO_PAGOS) then Response.Write "selected" %>><% =GF_TRADUCIR("No puede recibir pedidos ni pagos ") %></option>
						</select>
					</td>
				</tr>
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("Motivo (Justificación de ingreso a la lista negra.)") %></td>
					<td>
						<textarea name="motivo" id="motivo" rows="5" cols="30"><% =motivo %></textarea>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr><td>&nbsp;</td><tr>
</table>
<div id="botones"></div>
<input type="hidden" name="accion" value="<% =ACCION_GRABAR %>">
<input type="hidden" name="idEmpresa" value="<% =idEmpresa %>">
</form>
</body>
</html>