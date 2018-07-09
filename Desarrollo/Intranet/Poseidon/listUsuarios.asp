<!--#include file="../includes/procedimientos.asp"-->
<!--#include file="../includes/procedimientosUser.asp"-->
<!--#include file="../includes/procedimientosFechas.asp"-->
<!--#include file="../includes/procedimientosMG.asp"-->
<!--#include file="../includes/procedimientostraducir.asp"-->
<%
'ProcedimientoControl "GVPERMISOS"
%>
<HTML>
<HEAD>
   <TITLE>Listado de Usuarios en Planta </TITLE>
</HEAD>
<BODY>
<link href="../css/ActisaIntra-1.css" rel="stylesheet" type="text/css">
<div align="center">
   <table border="0" cellpadding="0" cellspacing="0" width="100%">
      <tr>
	     <td width="90%"><b><font class="Birthday">RMyD - Listado de Usuarios del Sistema</font></b></td>
	     <td align="center"><img width="40" height="40" SRC="Images/icon_user.gif"></td>
      </tr>
      <tr>
         <td>&nbsp;</td>
	     <td align="center" valign="top"><b><font name="Times New Roman" face="Verdana" size="4"><%Response.Write(request("Pto"))%></font></b></td>
      </tr>
   </table>
</div>
<hr>
<FORM action="listUsuarios.asp" method=POST id=form1 name=form1>
    <INPUT type="hidden" id="Pto" name="Pto" value=<%= Request("Pto")%>>
	<%
	strSQL = "Select a.cdusername, a.dsname, a.dslastname , a.icenabled, a.icexpire, (YEAR(a.dtexpiration)*10000 + Month(a.dtexpiration)*100 + DAY(a.dtexpiration)) as dtexpiration, a.crpassword, g.dsGroup,g.cdGroup,t.nutarjeta from dbo.Accounts a left join dbo.tarjetassupervisor t on t.cdusername=a.cdusername left join dbo.groups g on a.cdgroup=g.cdGroup  order by a.cdusername"
    if connect(Request("Pto")) then %>
       <table width="100%" border="0" align="left" cellpadding="0" cellspacing="0">
       <%
       Set Rs = connPorts.Execute(strsql)
       If Rs.Eof Then
	      Response.Write "No hay datos para la seleccion indicada."
       Else
          %>
          <tr>
             <td align="center" class="td-border">
                <table width="100%" id="tbl" class="table" border=0>
			       <tr>
              	      <th class="titu_header"><B><font>Usuario</font></B></th>
                      <th class="titu_header"><B><font>Nombre</font></B></th>
				      <th class="titu_header"><B><font>Apellido</font></B></th>
                      <th width="5%" class="titu_header"><B><font>Habil.</font></B></th>
                      <th width="5%" class="titu_header"><B><font>Expira</font></B></th>
                      <th class="titu_header"><B><font>Fecha Exp.</font></B></th>
  	                  <th class="titu_header"><B><font>Perfil</font></B></th>
   		           </tr>
                   <%
                   dim cont
                   cont = 0
                   while not Rs.EOF
                      if cont mod 2 = 0 then
                         MyBgColor = "white"
                      else
                         MyBgColor = "E0E0E0"
                      end if
                      cont = cont + 1
                      MyLastName = Trim(VerNull(Rs("dslastname")))
                      %>
                               <tr bgcolor="<%=MyBgColor%>">
                                  <TD> <%=Trim(VerNull(Rs("cdusername")))%> </TD>
                                  <TD> <% Response.Write Trim(VerNull(Rs("dsname")))%> </TD>
					   	          <TD> <% Response.Write MyLastName %> </TD>
					              <%
                                  if (VerNull(Rs("icenabled"))="S") then
                                     Response.Write "<TD align='center'> <img src='Images/icon_checked.gif'></TD>"
                                  else
                                     Response.Write "<TD align='center'> <img src='Images/icon_unchecked.gif'></TD>"
                                  end if
                                  if (VerNull(Rs("icexpire"))="S") then
                                     Response.Write "<TD align='center'> <img src='Images/icon_checked.gif'></TD>"
                                  else
                                     Response.Write "<TD align='center'> <img src='Images/icon_unchecked.gif'></TD>"
                                  end if
                                  %>
                                  <TD align="center">
                                     <% Response.Write GF_FN2DTE(Rs("dtexpiration")) %>
                                  </TD>
		      					  <TD> <% Response.Write Trim(VerNull(Rs("DsGroup")))%></TD>
                               </tr>
                      <%
                      Rs.MoveNext
                   wend
		           %>
                </table>
             </td>
		  </tr>
       <% 
       end if 
       %>
          <tr>
             <td colspan="7" align="right">&nbsp;</td>
          </tr>
          <tr>
             <td colspan="7" align="center">
                  <a class="simil" title="Imprimir Listado" onclick="JavaScript:window.print();">[Imprimir]</a>&nbsp;
                  <a class="simil" title="Volver al menu de Puertoe" onclick="JavaScript:history.back();">[Cancelar]</a>
             </td>
          </tr>
       </table>
    <%
    else
       Response.Write "<br><font color='red' size='+2'>No se pudo establecer la comunicación con la base de datos seleccionada.</font>"
    end if
    %>
</form>
</body>
</html>
