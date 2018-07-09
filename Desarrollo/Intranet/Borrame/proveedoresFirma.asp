<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosHKEY.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosProveedores.asp"-->

<%
'------------------------------------------------------------------------------------------------------
Function cargarFirmas(pIdEmpresa)
	Dim rsFirmas, connFirmas, strSQL
	strSQL = "Select * from TOEPFERDB.TBLEMPRESASFIRMAS where IDEMPRESA=" & pIdEmpresa & " order by SECUENCIA"
	Call executeQuery(rsFirmas,"OPEN", strSQL)									
	while not rsFirmas.eof
		select case cint(rsFirmas("SECUENCIA"))
			case cint(FIRMA_ROL_LEGALES)
				member1Cd = rsFirmas("CDUSUARIO")
				member1 = getUserDescription(member1Cd)
				if (rsFirmas("HKEY") <> "") then member1Firma = armarTextoFirma(rsFirmas("HKEY"), rsFirmas("FECHAFIRMA"))
		end select
		rsFirmas.movenext
	wend
End Function
'------------------------------------------------------------------------------------------------------
Function puedeFirmarAutorizaciones()
	Dim rol,rtrn,rsRegistros
	rtrn = false
    rol = getRolFirma(session("Usuario"), SEC_SYS_PROVEEDORES)    
	if (rol = FIRMA_ROL_LEGALES) then
		rtrn = true
	else
		setError(USUARIO_NO_AUTORIZADO)
	end if	
	puedeFirmarAutorizaciones = rtrn
End Function
'****************************************************************************
'							INICIO DE LA PAGINA
'****************************************************************************
	dim idProveedor,member1Cd, member1, member1Firma
	idProveedor = GF_PARAMETROS7("idProveedor",0,6)
	errFirma = GF_PARAMETROS7("errFirma","",6)
	if (errFirma <> "") then Call setError(errFirma)
	Call cargarFirmas(idProveedor)
	existeProv = loadDataDB(idProveedor)
%>
<html>
<head>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/hkey.js"></script>
<script type="text/javascript">
	// Se determina el explorador.	
	isFirefox=true; //FF
	if (navigator.userAgent.indexOf("MSIE")>=0) isFirefox=false; //IE
	var link = "proveedoresFirmarAutorizacion.asp?idProveedor=<% =idProveedor %>&secuencia=<%=FIRMA_ROL_LEGALES%>";
	var hkey0 = new Hkey('hk0', link, '<% =HKEY() %>', 'actualizar_callback()');
	
	function actualizar_callback(resp) {	
		document.getElementById("frmSel").submit();
	}
	function bodyOnLoad(){
		hkey0.start();
	}
</script>

</head>

<body onLoad="bodyOnLoad()">
<form name="frmSel" id="frmSel" method="POST">
<table width="40%" align="center" class="reg_header">
    <tr>
    	<td colspan="2">
        	 <%=showErrors()%>
        </td>
    </tr>
	<tr>
		<td colspan="2" class="reg_header_nav">
	    	<%=GF_Traducir("Autorizar nuevo proveedor")%>
		</td>
	</tr>
	<tr>
		<td width="20%" class="reg_header_navdos"><%=GF_Traducir("Nro. Proveedor")%></td>
		<td><%=idProveedor%></td>
	</tr>
	<tr>
	    <td class="reg_header_navdos"><%=GF_Traducir("Nombre")%></td>
		<td><%=razsoc%></td>
	</tr>
	<tr>
	    <td class="reg_header_navdos"><%=GF_Traducir("Nombre Ampliado")%></td>
	    <td><%=nomamp%></td>
	</tr>
	<tr>
	    <td class="reg_header_navdos"><%=getDsTipoDoc(tipdoc)%></td>
	    <td><%=nrodoc%></td>
	</tr>	
	<tr>
		<td colspan="2" height="100" align="center" class="recuadro round_border_bottom_left">
				<%	
					if (member1Firma  <> "" ) then %>
		            <img src="images/firmas/<% =obtenerFirma(member1Cd) %>"><br>
		            <% =member1Firma %>
		        <%	else
		                if (puedeFirmarAutorizaciones()) then	%>
		                    <br><div id="hk0"></div><br>
		            <%	else	%>
			            <%=showErrors()%>
		            <%	end if	
		        end if	%>
		        ________________________________________<br>
		      <%if (member1Firma <> "") then%>
				  	<%=member1%>
		      <%else%>
					<br>
		      <%end if%>
		</td>
	</tr>	
</table>
<input type="hidden" name="errFirma" id="errFirma">
</form>
</body>
</html>
