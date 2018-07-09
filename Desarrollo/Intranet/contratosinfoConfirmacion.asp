<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosAS400.asp"-->
<!--#include file="Includes/procedimientosMail.asp"-->

<%'
DIM g_intProducto, g_intSucursal, g_intOperacion, g_intNumero, g_intCosecha, g_intFechaConc, g_intUnitDest, g_intKCCOR, g_chrMrcConfirma, g_chrMrcRecibido, g_estado
DIM g_rsContratoSiguiente, g_intFechaConcertacion, cantMails, vectorMails(10), iMail, strTo

g_intProducto = GF_Parametros7("producto","",6)
g_intSucursal = GF_Parametros7("sucursal","",6)
g_intOperacion = GF_Parametros7("operacion","",6)
g_intNumero = GF_Parametros7("numero","",6)
g_intCosecha = GF_Parametros7("cosecha","",6)
g_intFechaConc = GF_Parametros7("fechaConc","",6)
g_intUnitDest = GF_Parametros7("unitDest","",6)
g_intFechaConcertacion = GF_Parametros7("FechaConcertacion","",6)
g_estado= GF_Parametros7("estado","",6)
g_intKCCOR = GF_Parametros7("corredor","",6)
g_chrMrcConfirma= "F"
g_chrMrcRecibido= "V"
cantMails = 0
%>
<html>
<head>
  <title></title>
  <LINK HREF="CSS/ActisaIntra-1.css" REL="stylesheet" TYPE="text/css">
  <link rel="stylesheet" href="../CSS/ActiSAIntra-1.css" type="text/css">
</head>
<body>
<br>
<table align=center border=0 width=90% cellspacing=0 cellpadding=0>
			 <tr height=8>
			     <td width="8"><img src="images/marco_r1_c1.gif" width="8" height="8"></td>
			     <td colspan=2 background="images/marco_r1_c2.gif"><img src="images/marco_r1_c2.gif" width="22" height="8"></td>
			     <td width="8"><img src="images/marco_r1_c3.gif" width="8" height="8"></td>
		     </tr>
		      <tr>
		          <td height="100%"><img src="images/marco_r2_c1.gif" width="8" height="100%"></td>
		          <%if g_estado = 2 or g_estado = 3 then%>
		              <td colspan=2 align="center" class="TDNOTICE"><b><% =GF_TRADUCIR("EL CONTRATO HA SIDO CONFIRMADO") %></b></td>
		          <%else%>
                      <td colspan=2  align="center" class="TDNOTICE"><b><% =GF_TRADUCIR("SE REQUIERE CONFIRMACION POR PARTE DE PERSONAL DE TOEPFER") %></b></td>
		          <%end if%>
		          <td height="100%"><img src="images/marco_r2_c3.gif" width="8" height="100%"></td>
		      </tr>
              <tr id="trLegal">
		          <td height="100%"><img src="images/marco_r2_c1.gif" width="8" height="100%"></td>
		          <td colspan=2 align="center">
                    <table>
                        <%if g_estado = 2 or g_estado = 3 then%>
                        <tr>
                            <td align="left" ><%=GF_Traducir("Contrato")%></td>
                            <td align="center">:</td>
                            <td align="left"><% =GF_EDIT_CONTRATO(g_intProducto,g_intSucursal,g_intOperacion,g_intNumero,g_intCosecha) %></td>
                        </tr>
						<tr><td colspan="3">&nbsp;</td></tr>	                        
                        <%end if%>
                    </table>
                  </td>
                  <td height="100%"><img src="images/marco_r2_c3.gif" width="8" height="100%"></td>
              </tr>
              
              <%
              'Response.write Session("KCOrganizacion")
              cantMails = obtenerMailConfirmaciones(Session("KCOrganizacion"), vectorMails)
              if cantMails = 0 then
              %>
              <tr id="trLegal">
		          <td height="100%"><img src="images/marco_r2_c1.gif" width="8" height="100%"></td>
		          <td colspan=2 align="left">
		         
                    <table width="100%">
                        <tr>
                            <td valign="top" align="left">
								
									<%
									response.write GF_Traducir("<b><u>Atención:</u></b> Si desea recibir esta información via e-mail por favor haga ") 
									Response.write "<a style='cursor:pointer;color:blue;' onclick='parent.configurarMail();'><u>" & GF_Traducir("click aqui") & "</u></a>"
									Response.Write GF_Traducir(" para configurar su dirección.")
								%>	
                            </td>
                        </tr>
                    </table>
                  </td>
                  <td height="100%"><img src="images/marco_r2_c3.gif" width="8" height="100%"></td>
              </tr>              
		      <%else%>
              <tr id="trLegal">
		          <td height="100%"><img src="images/marco_r2_c1.gif" width="8" height="100%"></td>
		          <td colspan=2 align="left">
		          <br>
                    <table width="100%">
                        <tr>
                            <td valign="top" align="left">
								
									<%
									response.write GF_Traducir("<b><u>Atención:</u></b> La dirección registrada para recibir esta información via e-mail es: ") 
								    iMail = 0
									while iMail < cantMails
									    strTo = strTo & vectorMails(iMail) & "; "
									    iMail = iMail + 1
									wend
									Response.write "'" & strTo & "'"
									Response.write "<br>Si desea modificarla por favor haga <a style='cursor:pointer;color:blue;' onclick='parent.configurarMail();'><u>" & GF_Traducir("click aqui.") & "</u></a>"
								%>	
                            </td>
                        </tr>
                    </table>
                  </td>
                  <td height="100%"><img src="images/marco_r2_c3.gif" width="8" height="100%"></td>
              </tr>              
		      <%end if%>
              <tr>
                 <td height="100%"><img src="images/marco_r2_c1.gif" width="8" height="100%"></td>
                 <td colspan=2>&nbsp;</td>
                 <td height="100%"><img src="images/marco_r2_c3.gif" width="8" height="100%"></td>
              </tr>		      
		      <tr height=8>
					<td height="100%"><img src="images/marco_r2_c1.gif" width="8" height="100%"></td>
                    <td colspan=2 align="center">
						<input type="button" name="Finalizar" value="<%=GF_Traducir("Finalizar")%>" onclick='javascript:parent.closePopUp();'>
					</td>
					<td height="100%"><img src="images/marco_r2_c3.gif" width="8" height="100%"></td>
		      </tr>

		      <tr height=8>
                    <td width="8"><img src="images/marco_r3_c1.gif" width="8" height="8"></td>
                    <td colspan=2 background="images/marco_r3_c2.gif"><img src="images/marco_r3_c2.gif" width="22" height="8"></td>
                    <td width="8"><img src="images/marco_r3_c3.gif" width="8" height="8"></td>
		      </tr>
	   </table>
