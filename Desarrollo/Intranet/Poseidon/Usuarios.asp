<!--#include file="../includes/procedimientos.asp"-->
<!--#include file="../includes/procedimientosUser.asp"-->
<!--#include file="../includes/procedimientosFechas.asp"-->
<!--#include file="../includes/procedimientosMG.asp"-->
<!--#include file="../includes/procedimientostraducir.asp"-->
<%
ProcedimientoControl "GVPERMISOS"
Dim Rs, strSQL
%>
<HTML>
<HEAD>
<TITLE>Consulta de Usuarios en Planta </TITLE>
<script language="JavaScript">
   function fcnResaltar(P_objFila)
   {
      P_objFila.bgColor='#9999FF';
      P_objFila.style.color= '#FFFFFF';
   }
   function fcnNormal(P_objFila)
   {
      P_objFila.bgColor='WhiteSmoke';
      P_objFila.style.color= '#000000';
   }
</script>
</HEAD>
<BODY>
<link href="../css/ActisaIntra-1.css" rel="stylesheet" type="text/css">
<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" width="100%">
		<tr>
			<td width="90%"><b><font class="Birthday">RMyD - Usuarios del Sistema</font></b></td>
			<td align="center"><img width="40" height="40" SRC="Images/icon_user.gif"></td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td align="center" valign="top"><b><font name="Times New Roman" face="Verdana" size="4"><%Response.Write(request("Pto"))%></font></b></td>
		</tr>
	</table>
</div>
<hr>
<form action="usuarios.asp" method=POST id=form1 name=form1>
<input type="hidden" id="Pto" name="Pto" value = <%= Request("Pto")%>>
<input type="hidden" name="actionABM" value="">
<%
strSQL = "Select a.cdusername, a.dsname, a.dslastname , a.icenabled, a.icexpire, (YEAR(a.dtexpiration)*10000 + Month(a.dtexpiration)*100 + DAY(a.dtexpiration)) as dtexpiration, a.crpassword, g.dsGroup,g.cdGroup,t.nutarjeta from dbo.Accounts a left join dbo.tarjetassupervisor t on t.cdusername=a.cdusername left join dbo.groups g on a.cdgroup=g.cdGroup  order by a.cdusername"
if connect(Request("Pto")) then %>
<table width="100%" border="0" align="left" cellpadding="0" cellspacing="0"><%
	Set Rs = connPorts.Execute(strsql)
    If Rs.Eof Then
		Response.Write "No hay datos para la seleccion indicada."
    else %>
		<tr>
			<td align="center" class="td-border">
				<div id="tbl-container">
					<table width="100%" id="tbl" class="table" border=0>
						<tr>
							<td class="titu_header"><B><font>Usuario			</font></B></td>
		                    <td class="titu_header"><B><font>Nombre				</font></B></td>
					        <td class="titu_header"><B><font>Apellido			</font></B></td>
		                    <td width="5%" class="titu_header"><B><font>Habil.	</font></B></td>
		                    <td width="5%" class="titu_header"><B><font>Expira	</font></B></td>
		                    <td class="titu_header"><B><font>Fecha Exp.			</font></B></td>
					        <td class="titu_header"><B><font>Perfil				</font></B></td>
						</tr>
		                <% 
		                while not Rs.EOF 
			                MyLastName = Trim(VerNull(Rs("dslastname"))) %>
							<tr onMouseOver="fcnResaltar(this)" onMouseOut="fcnNormal(this)">
								<td> <%=Trim(VerNull(Rs("cdusername")))%>			</td>
								<td> <% Response.Write Trim(VerNull(Rs("dsname")))%></td>
								<td> <% Response.Write MyLastName %>				</td>
						        <%
									if (VerNull(Rs("icenabled"))="S") then
									   response.write "<TD align='center'> <img src='Images/icon_checked.gif'></TD>"
									else
									   response.write "<TD align='center'> <img src='Images/icon_unchecked.gif'></TD>"
									end if
									if (VerNull(Rs("icexpire"))="S") then
									   response.write "<TD align='center'> <img src='Images/icon_checked.gif'></TD>"
									else
									   response.write "<TD align='center'> <img src='Images/icon_unchecked.gif'></TD>"
									end if
								%>
								<td align="center">
		                        <% Response.Write GF_FN2DTE(Rs("dtexpiration")) %>
		                        </td>
		  					    <td> <% Response.Write Trim(VerNull(Rs("DsGroup")))%>	</td>
							</tr>
							<%
		                      Rs.MoveNext
							wend
			                %>
					</table>
				</div>
			</td>
		</tr>
 <% end if %>        
	</table>
    <%
	else
		Response.Write "<br><font color='red' size='+2'>No se pudo establecer la comunicación con la base de datos seleccionada.</font>"
    end if
    %>
</form>
</body>
</html>
