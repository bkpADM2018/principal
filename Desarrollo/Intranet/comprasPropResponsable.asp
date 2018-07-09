<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/GF_MGSRADD.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosSql.asp"-->
<%
Call comprasControlAccesoCM(RES_ADM)

dim idResponsable, accion, strSQL, cdResponsable, dsResponsable, rsRegistro, myHKEY
dim myChecked, myValue, myBaseValue, myPageValue
Dim myChkExportacionAj, myChkArroyoAj, myChkPiedrabuenaAj, myChkTransitoAj
dim chkAsientoEXP, chkAsientoARR, chkAsientoBBA, chkAsientoTRA
dim	modificaArt, confirmaContratos
Dim chkExportacionAjPto, chkArroyoAjPto, chkPiedrabuenaAjPto, chkTransitoAjPto

idResponsable = GF_PARAMETROS7("idResponsable",0,6)
accion = GF_PARAMETROS7("accion","",6)

myChkExportacionAj = GF_PARAMETROS7("chkExportacionAj",0,6)
myChkArroyoAj = GF_PARAMETROS7("chkArroyoAj",0,6)
myChkPiedrabuenaAj = GF_PARAMETROS7("chkPiedrabuenaAj",0,6)
myChkTransitoAj = GF_PARAMETROS7("chkTransitoAj",0,6)

chkAsientoEXP = GF_PARAMETROS7("chkAsientoEXP",0,6)
chkAsientoARR = GF_PARAMETROS7("chkAsientoARR",0,6)
chkAsientoBBA = GF_PARAMETROS7("chkAsientoBBA",0,6)
chkAsientoTRA = GF_PARAMETROS7("chkAsientoTRA",0,6)

modificaArt = GF_PARAMETROS7("modificaArt",0,6)
confirmaContratos = GF_PARAMETROS7("confirmaContratos",0,6)

chkExportacionAjPto = GF_PARAMETROS7("chkExportacionAjPto",0,6)
chkArroyoAjPto = GF_PARAMETROS7("chkArroyoAjPto",0,6)
chkPiedrabuenaAjPto = GF_PARAMETROS7("chkPiedrabuenaAjPto",0,6)
chkTransitoAjPto = GF_PARAMETROS7("chkTransitoAjPto",0,6)

myHKEY = GF_PARAMETROS7("hkey","",6)

myHKEYOriginal = GF_PARAMETROS7("HKEYOriginal","",6)

'Obtener datos profesional
strSQL="Select * from WFPROFESIONAL where IdProfesional=" & idResponsable
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
if not rs.eof then
	cdResponsable = UCase(rs("CDUSUARIO"))
	dsResponsable = rs("NOMBRE")
end if

strSQL="Select * from TBLREGISTROFIRMAS where CDUSUARIO='" & UCase(cdResponsable) & "'"
Call executeQueryDB(DBSITE_SQL_INTRA, rsRegistro, "OPEN", strSQL)

if ucase(accion) = "GRABAR" then
	if (rsRegistro.eof) then
		strSQL= "Insert into TBLREGISTROFIRMAS(HKEY, AJARROYO, AJTRANSITO, AJPIEDRABUENA, AJEXPORTACION, ASARROYO, ASTRANSITO, ASPIEDRABUENA, ASEXPORTACION, MODIFICAARTICULOS, CDUSUARIO, CONFIRMACONTRATOS, AJPTOARROYO, AJPTOTRANSITO, AJPTOPIEDRABUENA, AJPTOEXPORTACION) " &_
		        "values ('" & myHKEY & "'," & myChkArroyoAj & "," & myChkTransitoAj & "," & myChkPiedrabuenaAj & "," & myChkExportacionAj & "," & chkAsientoEXP & "," & chkAsientoARR & "," & chkAsientoBBA & "," & chkAsientoTRA & "," & modificaArt & ",'" & UCase(cdResponsable) & "'," & confirmaContratos & ","& chkArroyoAjPto &","& chkTransitoAjPto &","& chkPiedrabuenaAjPto &","& chkExportacionAjPto &")"		
	else
		strSQL="Update TBLREGISTROFIRMAS set HKEY='" & myHKEY & "', AJARROYO=" & myChkArroyoAj & ", AJTRANSITO=" & myChkTransitoAj & ", AJPIEDRABUENA=" & myChkPiedrabuenaAj & ", AJEXPORTACION=" & myChkExportacionAj & ",ASARROYO=" & chkAsientoARR & ", ASTRANSITO=" & chkAsientoTRA & ", ASPIEDRABUENA=" & chkAsientoBBA & ", ASEXPORTACION=" & chkAsientoEXP & ", MODIFICAARTICULOS=" & modificaArt & ", CONFIRMACONTRATOS=" & confirmaContratos & ",AJPTOARROYO=" & chkArroyoAjPto& ", AJPTOTRANSITO=" & chkTransitoAjPto & ", AJPTOPIEDRABUENA=" & chkPiedrabuenaAjPto & ", AJPTOEXPORTACION=" & chkExportacionAjPto &"  where CDUSUARIO='" & UCase(cdResponsable) & "'"
	end if
	Call executeQueryDB(DBSITE_SQL_INTRA, rsRegistro, "EXEC", strSQL)
else
	if (not rsRegistro.eof) then
		myHKEY = rsRegistro("HKEY")
				
		myChkExportacionAj = rsRegistro("AJEXPORTACION")
		myChkArroyoAj      = rsRegistro("AJARROYO")
		myChkPiedrabuenaAj = rsRegistro("AJPIEDRABUENA")
		myChkTransitoAj    = rsRegistro("AJTRANSITO")
		
		chkAsientoEXP = rsRegistro("ASEXPORTACION") 
		chkAsientoARR = rsRegistro("ASARROYO")
		chkAsientoBBA = rsRegistro("ASPIEDRABUENA")
		chkAsientoTRA = rsRegistro("ASTRANSITO")
		
		chkExportacionAjPto = rsRegistro("AJPTOEXPORTACION")
		chkArroyoAjPto		= rsRegistro("AJPTOARROYO") 
		chkPiedrabuenaAjPto = rsRegistro("AJPTOPIEDRABUENA")
		chkTransitoAjPto	= rsRegistro("AJPTOTRANSITO")
		modificaArt = rsRegistro("MODIFICAARTICULOS")
		confirmaContratos = rsRegistro("CONFIRMACONTRATOS")
	end if
end if
%>
<html>
<head>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript">
function responsableOnLoad() {
	refPopUpResponsable = getObjPopUp('popupResponsable');
	<% if (accion = ACCION_CERRAR or accion = ACCION_GRABAR) then %>
		<% if (myHKEY <> myHKEYOriginal) then %>
			window.location="comprasmotivoCambioHKEY.asp?idResponsable=<%=idResponsable%>&HkeyNew=<%=myHKEY%>&HkeyOld=<%=myHKEYOriginal%>"
		<% else %>
			refPopUpResponsable.hide();
		<% end if %>
	<% end if %>
	document.getElementById("rolFirma").focus();
	pngfix();
}
function detectarCambioHKEY(valor){
	document.getElementById("cambioHKEY").value = valor;
	if (valor == "TRUE"){
		document.getElementById("HKEYNew").value = '<%=myHKEY%>';
		document.getElementById("HKEYOld").value = '<%=myHKEYOriginal%>';
	}
}
</script>
</head>
<body onLoad="responsableOnLoad()">
<form name="frmSel" method="post" action="comprasPropResponsable.asp">


<input type="hidden" name="HKEYNew" id="HKEYNew" value="">
<input type="hidden" name="HKEYOld" id="HKEYOld" value="">
<input type="hidden" name="cambioHKEY" id="cambioHKEY" value="">


<table align="center" border=0 width="100%">
	<tr>
		<td class="title_sec_section" align="left" colspan="2">
			<img align="absMiddle" src="images/compras/users-32x32.png">
			<b><% =UCase(cdResponsable) %> - <% =UCase(dsResponsable) %></b>		</td>
	</tr>	
	<tr>
		<td colspan="2"><% call showErrors() %></td>
	</tr>
	<tr><td>&nbsp;</td><tr>
	<tr>
		<td><b><%= GF_Traducir("Llave (HARD KEY)")%></b></td>
		<td><INPUT type="text" size="10" maxlength="8" name="hkey" value="<% =myHKEY  %>" onBlur="detectarCambioHKEY(this)"></td>
	</tr>	
	<tr>
	  <td colspan="2" align="left"><b><%= GF_Traducir("Ajuste de Stock")%></b></td>
    </tr>
	<tr>
	  <td colspan="2" align="center"><table width="80%" border="0" cellpadding="0" cellspacing="0" class="reg_header">
        <tr>
          <td align="center" class="reg_header_nav round_border_top_left"><%= GF_Traducir("Exportación")%></td>
          <td align="center" class="reg_header_nav"><%= GF_Traducir("Arroyo")%></td>
          <td align="center" class="reg_header_nav"><%= GF_Traducir("Piedrabuena")%></td>
          <td align="center" class="reg_header_nav round_border_top_right"><%= GF_Traducir("Tránsito")%></td>
        </tr>
        <tr>
          <td align="center" width="25%" >
       		  <% 	
					myChecked = ""
					if myChkExportacionAj = 1 then myChecked = "Checked"
				%>
              <input style="border:none;cursor:pointer;" type="checkbox" name="chkExportacionAj" value="1" <%=myChecked%>>          </td>
          <td align="center" width="25%">
       		  <% 
					myChecked = ""
					if myChkArroyoAj = 1 then myChecked = "Checked"
				%>
       		  <input style="border:none;cursor:pointer;" type="checkbox" name="chkArroyoAj" value="1" <%=myChecked%>>          </td>
          <td align="center" width="25%">
       		  <% 
					myChecked = ""
					if myChkPiedrabuenaAj = 1 then myChecked = "Checked"
				%>
       		  <input style="border:none;cursor:pointer;" type="checkbox" name="chkPiedrabuenaAj" value="1" <%=myChecked%>>          </td>
          <td align="center" width="25%">
       		  <% 
					myChecked = ""
					if myChkTransitoAj = 1 then myChecked = "Checked"
				%>
              <input style="border:none;cursor:pointer;" type="checkbox" name="chkTransitoAj" value="1" <%=myChecked%>>          </td>
        </tr>
      </table></td>
    </tr>
	<tr>
	  <td colspan="2" align="left"><b><%= GF_Traducir("Responsable de Asientos Contables")%></b></td>
    </tr>
	<tr>
		<td colspan="2" align="center">
			<table width="80%" border="0" cellpadding="0" cellspacing="0" class="reg_header">
				<tr>
					<td align="center" class="reg_header_nav round_border_top_left"><%= GF_Traducir("Exportación")%></td>
					<td align="center" class="reg_header_nav"><%= GF_Traducir("Arroyo")%></td>
					<td align="center" class="reg_header_nav"><%= GF_Traducir("Piedrabuena")%></td>
					<td align="center" class="reg_header_nav round_border_top_right"><%= GF_Traducir("Tránsito")%></td>
			    </tr>
				<tr>
					<td align="center" width="25%">
						<input style="border:none;cursor:pointer;" type="checkbox" name="chkAsientoEXP" value="1" <%if chkAsientoEXP = 1 then Response.Write "Checked"%>>          
					</td>
					<td align="center" width="25%">
       					<input style="border:none;cursor:pointer;" type="checkbox" name="chkAsientoARR" value="1" <%if chkAsientoARR = 1 then Response.Write "Checked"%>>
					</td>          
					<td align="center" width="25%">
       					<input style="border:none;cursor:pointer;" type="checkbox" name="chkAsientoBBA" value="1" <%if chkAsientoBBA = 1 then Response.Write "Checked"%>>          
       				</td>
					<td align="center" width="25%">
						<input style="border:none;cursor:pointer;" type="checkbox" name="chkAsientoTRA" value="1" <%if chkAsientoTRA = 1 then Response.Write "Checked"%>>          
					</td>
				</tr>
			</table>
		</td>
    </tr>
    
	<tr>
	  <td colspan="2" align="left"><b><%= GF_Traducir("Modificación de Articulos")%></b></td>
    </tr>
	<tr>
	  <td colspan="2" align="center"><table width="80%" border="0" cellpadding="0" cellspacing="0" class="reg_header">
        <tr>
          <td align="center" class="reg_header_nav round_border_top_left"><%= GF_Traducir("Autorizado")%></td>
          <td align="center" class="reg_header_nav"><%= GF_Traducir("Denegado")%></td>
        </tr>
        <tr>
          <td align="center" width="50%" >
              <input style="border:none;cursor:pointer;" type="radio" name="modificaArt" value="1" <%if modificaArt = 1 then Response.Write "Checked"%>></td>
          <td align="center" width="50%">
       		  <input style="border:none;cursor:pointer;" type="radio" name="modificaArt" value="0" <%if modificaArt = 0 then Response.Write "Checked"%>></td>
        </tr>
      </table></td>
    </tr>
    
	<tr>
	  <td colspan="2" align="left"><b><%= GF_Traducir("Confirmación de Contratos")%></b></td>
    </tr>
	<tr>
	  <td colspan="2" align="center"><table width="80%" border="0" cellpadding="0" cellspacing="0" class="reg_header">
        <tr>
          <td align="center" class="reg_header_nav round_border_top_left"><%= GF_Traducir("Autorizado")%></td>
          <td align="center" class="reg_header_nav"><%= GF_Traducir("Denegado")%></td>
        </tr>
        <tr>
          <td align="center" width="50%" >
              <input style="border:none;cursor:pointer;" type="radio" name="confirmaContratos" value="1" <%if confirmaContratos = 1 then Response.Write "Checked"%>></td>
          <td align="center" width="50%">
       		  <input style="border:none;cursor:pointer;" type="radio" name="confirmaContratos" value="0" <%if confirmaContratos = 0 then Response.Write "Checked"%>></td>
        </tr>
      </table></td>
    </tr>
	<tr>
	  <td colspan="2" align="left"><b><%= GF_Traducir("Ajuste de Puertos")%></b></td>
    </tr>
	<tr>
	  <td colspan="2" align="center">
		 <table width="80%" border="0" cellpadding="0" cellspacing="0" class="reg_header">
			<tr>
			  <td align="center" class="reg_header_nav round_border_top_left"><%= GF_Traducir("Exportación")%></td>
			  <td align="center" class="reg_header_nav"><%= GF_Traducir("Arroyo")%></td>
			  <td align="center" class="reg_header_nav"><%= GF_Traducir("Piedrabuena")%></td>
			  <td align="center" class="reg_header_nav round_border_top_right"><%= GF_Traducir("Tránsito")%></td>
			</tr>
			<tr>
			  <td align="center" width="25%" >
       			  <% 	
						myCheckedPto = ""
						if chkExportacionAjPto = 1 then myCheckedPto = "Checked"
					%>
			      <input style="border:none;cursor:pointer;" type="checkbox" name="chkExportacionAjPto" value="1" <%=myCheckedPto%>>
			  </td>
			  <td align="center" width="25%">
       			  <% 
						myCheckedPto = ""
						if chkArroyoAjPto = 1 then myCheckedPto = "Checked"
					%>
       			  <input style="border:none;cursor:pointer;" type="checkbox" name="chkArroyoAjPto" value="1" <%=myCheckedPto%>>
       		  </td>
			  <td align="center" width="25%">
       			  <% 
						myCheckedPto = ""
						if chkPiedrabuenaAjPto = 1 then myCheckedPto = "Checked"
					%>
       			  <input style="border:none;cursor:pointer;" type="checkbox" name="chkPiedrabuenaAjPto" value="1" <%=myCheckedPto%>>
       		  </td>
			  <td align="center" width="25%">
       			  <% 
						myCheckedPto = ""
						if chkTransitoAjPto = 1 then myCheckedPto = "Checked"
					%>
			      <input style="border:none;cursor:pointer;" type="checkbox" name="chkTransitoAjPto" value="1" <%=myCheckedPto%>>
			  </td>
			</tr>
	     </table>
      </td>
    </tr>
	<tr>	
		<td colspan="2" align="center">
			<table>	
				<tr><td><br>
					<%  if (not isAuditor(SIN_DIVISION)) then %>
					<input type="submit" id="aceptar" name="aceptar" value="<% =GF_TRADUCIR("Aceptar") %>" <% if (idResponsable = 0) then response.write "disabled=true" %>>					
					<%	end if %>
				</td></tr>
			</table>		</td>		
	</tr>	
</table>
<input type="hidden" name="accion" value="<% =ACCION_GRABAR %>">
<input type="hidden" name="idResponsable" value="<% =idResponsable %>">
<input type="hidden" name="HKEYOriginal" id="HKEYOriginal" value="<%=myHKEY%>">


</form>
</body>
</html>