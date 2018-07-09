<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosAS400.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosvalidacion.asp"-->
<!--#include file="Includes/ExternalFunctions.asp"-->
<%dim p_cuit1, p_cuit2, p_cuit3, rSocial
dim strSQL, rs, conn, codigoP, estado

p_cuit1 = trim(GF_Parametros7("p_cuit1","",6))
p_cuit2 = trim(GF_Parametros7("p_cuit2","",6))
p_cuit3 = trim(GF_Parametros7("p_cuit3","",6))
%>
<html>
<head>
    <title><%=GF_Traducir("Consulta de Proveedores")%></title>
    <link href="CSS/ActisaIntra-1.css" rel="stylesheet" type="text/css">
</head>

<body onKeyPress="javascript: if (window.event.keyCode == 13) document.forms[0].submit();">
<form method="POST" action="consultaProveedores.asp">
    <%call GF_TITULO_2(GF_Traducir("Consulta de Proveedores"))%>
        <table width="250" cellspacing="0" cellpadding="0" align="center" border="0">
          <tr>
              <td width="8"><img src="images/marco_r1_c1.gif"></td>
              <td><img src="images/marco_r1_c2.gif" width="100%" height="8"></td>
              <td width="8"><img src="images/marco_r1_c3.gif"></td>
          </tr>
          <tr>
              <td width="8" height="100%"><img src="images/marco_r2_c1.gif" width="8" height="100%"></td>
              <td align="center">
                <table align="center" border=0 cellpadding=0 cellspacing=7>
                    <tr height=30 valign="top">
                        <td align=center><font class="big"><b><% =GF_TRADUCIR("C.U.I.T. del Proveedor") %></b></font></td>
                    </tr>
                    <tr>
                        <td align="center">C.U.I.T. : <input type=text name="p_cuit1" id="p_cuit1" size=1 maxlength=2 value="<%=p_cuit1%>" tabindex=1>-<input type=text name="p_cuit2" id="p_cuit2" size=8 maxlength=8 value="<%=p_cuit2%>" tabindex=2>-<input type=text name="p_cuit3" id="p_cuit3" style="width:12px" maxlength=1 value="<%=p_cuit3%>" tabindex=3></td>
                    </tr>
                    <tr>
                        <td align="center"><input type="submit" value="<%=GF_Traducir("Consultar")%>" tabindex=3></td>
                    </tr>
                </table>
              </td>
              <td width="8" height="100%"><img src="images/marco_r2_c3.gif" width="8" height="100%"></td>
          </tr>
          <tr>
              <td width="8"><img src="images/marco_r3_c1.gif"></td>
              <td><img src="images/marco_r3_c2.gif" width="100%" height="8"></td>
              <td width="8"><img src="images/marco_r3_c3.gif"></td>
          </tr>
        </table>
    <% if (p_cuit1 & p_cuit2 & p_cuit3 <> "")then %>
        <br><br>
        <table width="60%" cellspacing="0" cellpadding="0" align="center" border="0">
          <tr>
              <td width="8"><img src="images/marco_r1_c1.gif"></td>
              <td width="100%"><img src="images/marco_r1_c2.gif" width="100%" height="8"></td>
              <td width="8"><img src="images/marco_r1_c3.gif"></td>
          </tr>
          <% if (GF_CONTROL_CUIT(p_cuit1 & p_cuit2 & p_cuit3)) then
					rSocial = GetDsEnterprise3(p_cuit1 & p_cuit2 & p_cuit3)
					codigoP = GetCDEnterprise3(p_cuit1 & p_cuit2 & p_cuit3)
					if ((rSocial <> DS_PROV_NO_EXISTE) and (codigoP < ID_PROV_MAX)) then %>
                        <tr>
                        <td width="8" height="100%"><img src="images/marco_r2_c1.gif" width="8" height="100%"></td>
						<td align="center"><b><% =rSocial %></b>&nbsp;<% =GF_Traducir("ha presentado toda la informacion impositiva correspondiente.") %></td>
						<td width="8" height="100%"><img src="images/marco_r2_c3.gif" width="8" height="100%"></td>
                        </tr>
				<%
					else %>
                        <tr>
                        <td width="8" height="100%"><img src="images/marco_r2_c1.gif" width="8" height="100%"></td>
						<td align="center"><B><% =GF_Traducir("El C.U.I.T. ingresado no se encuentra en nuestro Registro.") %></B></td>
						<td width="8" height="100%"><img src="images/marco_r2_c3.gif" width="8" height="100%"></td>
                        </tr>
                        <tr>
                        <td width="8" height="10"><img src="images/marco_r2_c1.gif" width="8" height="100%"></td>
                        <td width="100%"></td>
                        <td width="8" height="10"><img src="images/marco_r2_c3.gif" width="8" height="100%"></td>
                        </tr>
                        <tr>
                        <td width="8" height="100%"><img src="images/marco_r2_c1.gif" width="8" height="100%"></td>
                        <td>
                        <% =GF_TRADUCIR("Para Ingresar al registro se debe presentar la siguiente documentacion:") %><br><br>
						<img src="images/pfeil.gif">&nbsp;<% =GF_TRADUCIR("Constancia de Inscripción en IVA") %><br>
						<img src="images/pfeil.gif">&nbsp;<% =GF_TRADUCIR("Constancia de Inscripción en Ingresos Brutos") %><br>
						<img src="images/pfeil.gif">&nbsp;<% =GF_TRADUCIR("Copia de la consulta de la pagina de AFIP – RG 1394/02") %><br>
						<img src="images/pfeil.gif">&nbsp;<% =GF_TRADUCIR("Carta con el C.B.U.") %><br><br>
						<% =GF_TRADUCIR("Toda la documentación firmada en Original, en sobre dirigida al Sr. Ariel Marelli y/o Sr. Daniel Gonzalez") %>
						</td>
                        <td width="8" height="100%"><img src="images/marco_r2_c3.gif" width="8" height="100%"></td>
                        </tr>
				<% end if
				else
				'CUIT INVALIDO
				%>
				        <tr>
                        <td width="8" height="100%"><img src="images/marco_r2_c1.gif" width="8" height="100%"></td>
						<td align="center"><% =GF_Traducir("El numero de C.U.I.T. ingresado no es valido.") %></td>
						<td width="8" height="100%"><img src="images/marco_r2_c3.gif" width="8" height="100%"></td>
                        </tr>
                <%end if %>
          <tr>
              <td width="8"><img src="images/marco_r3_c1.gif"></td>
              <td width="100%"><img src="images/marco_r3_c2.gif" width="100%" height="8"></td>
              <td width="8"><img src="images/marco_r3_c3.gif"></td>
          </tr>
        </table>
    <%end if%>
</form>
</body>
<script language="javascript">
    document.getElementById('p_cuit1').focus();
</script>
</html>
