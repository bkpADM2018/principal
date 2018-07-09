<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosseguridad.asp"-->
<%
'<!------------------------------------------------------------------->
'<!--                        INTRANET ACTISA                        -->
'<!--                                                               -->
'<!--               Programador: Ezequiel A. Bacarini               -->
'<!--               Fecha      : 16/01/2008                         -->
'<!--               Pagina     : AUPGetSectores.ASP                 -->
'<!--               Descripcion: Listado de personal por sector     -->
'<!------------------------------------------------------------------->
Dim strSQL, rsPersonas, oConn, MyClass
dim FrmDic
'Se crea el diccionario de parametros.
set FrmDic= CreateObject ("Scripting.Dictionary") 
For Each i in Request.QueryString
   FrmDic.Add  i,Request.QueryString(i).item
Next
 

function PrepareWord(pWord)
dim auxWord
	auxWord = replace(pWord,"á","a")
	auxWord = replace(auxWord,"é","e")
	auxWord = replace(auxWord,"í","i")
	auxWord = replace(auxWord,"ó","o")
	auxWord = replace(auxWord,"ú","u")
	auxWord = replace(auxWord,"ñ","n")
	PrepareWord = auxWord
end function
MyClass = "reg_header_nav"
if FrmDic("MODO") = "V" then MyClass = "TDERROR2"
%>
<table width="90%" border=0 align="center" class="reg_header" cellpadding="2" cellspacing="1">		
	<!--<tr class=reg_header_navdos>-->
	<tr class="<%=MyClass%>">
		<td align=center><b><%=GF_Traducir("Apellido")%></b></td>
		<td align=center><b><%=GF_Traducir("Nombre")%></b></td>	
		<td align=center><b><%=GF_Traducir("Legajo")%></b></td>
		<% if session("AUPUSER") = "ADMIN" then %>
		<td align=center><b>&nbsp;</b></td>
		<% end if %>
	</tr>
	<%
	strSQL= "select IdProfesional, Apellido, nombre, NroLegajo from Profesionales P inner join Personas Pe on Pe.idpersona=P.idProfesional where P.Sector=" & replace(FrmDic("SECTOR"),"SEC_","") & " and EgresoValido='" & FrmDic("MODO") & "' order by Apellido"
	'Response.Write strSQL
	call GF_BD_CONTROL (rsPersonas,oConn,"OPEN",strSQL)
	while not(rsPersonas.EOF)
			%>
			<tr style="cursor:pointer;" class="reg_header_navdos" onMouseOver="fcnResaltar(this)" onMouseOut="fcnNormal(this)" title="<%=GF_Traducir("Modificar el perfil de este usuario")%>">
			    <td onclick="javascript:loadAndSubmit(<%=rsPersonas("IdProfesional")%>);" align="center" width="30%"> 
					<font><%=PrepareWord(rsPersonas("Apellido"))%></font>
				</td>
			    <td onclick="javascript:loadAndSubmit(<%=rsPersonas("IdProfesional")%>);" align="center" width="40%"> 
					<font><%=PrepareWord(rsPersonas("Nombre"))%></font>
				</td>
				<td onclick="javascript:loadAndSubmit(<%=rsPersonas("IdProfesional")%>);" align="center" width="20%"> 
					<font><%=decrypt(rsPersonas("nroLegajo"))%></font>
				</td>
				<% if session("AUPUSER") = "ADMIN" then %>
				<td align=center width="2%" nowrap> 
					<a title="<%=GF_Traducir("Reporte para Auditoria por Usuario")%>" target="_new" href="AUPAuditoria.asp?pUser=<%=rsPersonas("IdProfesional")%>"><img src="Images/printer1.png"></a>
				</td>
				<% end if %>
			</tr>
			<%     
			rsPersonas.MoveNext 
	wend  
	call GF_BD_CONTROL (rsPersonas,oConn,"CLOSE",strSQL)	
	%>
</table>
