<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosUnificador.asp"-->
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
function IsToepfer(p_krempresa,cual)
	dim rtrn
	dim codigo
	
	select case cual
		case "KC"
			codigo = "07431"
		case "KR"
			codigo = "8163"
	end select
	
	if cstr(p_krempresa) = codigo  then
		rtrn = true
	else
		rtrn = false
	end if	
	IsToepfer = rtrn
end function
'-------------------------------------------------------------------
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
'-------------------------------------------------------------------
function LogoEmpresa(pCodigo)
	if IsToepfer(rsPersonas("empresa"),"KR") then%>
		<img src="Images/<%=pCodigo%>.png" width="20" height="20">
	<%else
	
		select case pCodigo
			
		end select
	end if
end function
'-------------------------------------------------------------------
Dim strSQL, rsPersonas, oConn, MyClass,MyActisaKR, MyActisaDS
dim FrmDic,mySub, myORDS, myORKC, myORKR
'Se crea el diccionario de parametros.
set FrmDic= CreateObject ("Scripting.Dictionary") 
For Each i in Request.QueryString
   FrmDic.Add  i,Request.QueryString(i).item
Next

MyClass = "reg_header_nav"
if FrmDic("MODO") = "V" then 
	MyClass = "TDBAJAS"
elseif FrmDic("MODO") = "E" then 
	MyClass = "TDEXTERNOS"
end if	
%>
<table width="90%" border=0 align="center" class="reg_header" cellpadding="2" cellspacing="1">		
	<!--<tr class=reg_header_navdos>-->

	<%
	'strSQL= "select IdProfesional, Apellido, nombre, NroLegajo from Profesionales P inner join Personas Pe on Pe.idpersona=P.idProfesional where P.Sector=" & replace(FrmDic("SECTOR"),"SEC_","") & " and EgresoValido='" & FrmDic("MODO") & "' order by Apellido"
	call GF_MGC ("OR", "07431", MyActisaKR, MyActisaDS)
	
	if FrmDic("MODO") = "V" or FrmDic("MODO") = "F" then 
		'mySub = " and EgresoValido='" & FrmDic("MODO") & "' and Empresa=" & MyActisaKR 
		mySub = " and EgresoValido='" & FrmDic("MODO") & "'"
	elseif FrmDic("MODO") = "E" then
		mySub = " and (Empresa<>" & MyActisaKR & " or Empresa is null)"
	end if
		
	strSQL= "select IdProfesional, Apellido, nombre, NroLegajo, Empresa from Profesionales P inner join Personas Pe on Pe.idpersona=P.idProfesional where P.Sector=" & replace(FrmDic("SECTOR"),"SEC_","") & mySub & " order by Apellido"
	'Response.Write strSQL
	call GF_BD_CONTROL (rsPersonas,oConn,"OPEN",strSQL)
	
	if not rsPersonas.EOF then
		%>
		<tr class="<%=MyClass%>">
			<td align=center><b><%=GF_Traducir("Empresa")%></b></td>
			<td align=center><b><%=GF_Traducir("Apellido")%></b></td>
			<td align=center><b><%=GF_Traducir("Nombre")%></b></td>	
			<td align=center><b>
			<% 
				if FrmDic("MODO") = "E" then 
					Response.Write GF_Traducir("Empresa")
				else
					Response.Write GF_Traducir("Legajo")
				end if 
			%>
			</b></td>
			<% if session("AUPUSER") = "ADMIN" then %>
			<td align=center><b>&nbsp;</b></td>
			<% end if %>
		</tr>
		<%
	end if	
	while not(rsPersonas.EOF)
			%>
			<tr style="cursor:pointer;" class="reg_header_navdos" onMouseOver="fcnResaltar(this)" onMouseOut="fcnNormal(this)" title="<%=GF_Traducir("Modificar el perfil de este usuario")%>">
			    <td align='center'>
						<%=LogoEmpresa(rsPersonas("Empresa"))%>
				</td>
				<td onclick="javascript:loadAndSubmit(<%=rsPersonas("IdProfesional")%>);" align="center" width="30%"> 
					<font><%=PrepareWord(rsPersonas("Apellido"))%></font>
				</td>
			    <td onclick="javascript:loadAndSubmit(<%=rsPersonas("IdProfesional")%>);" align="center" width="40%"> 
					<font><%=PrepareWord(rsPersonas("Nombre"))%></font>
				</td>
				<td onclick="javascript:loadAndSubmit(<%=rsPersonas("IdProfesional")%>);" align="center" width="20%"> 
				<b><font>
					<% 
					if FrmDic("MODO") = "E" then 
						myORKC = ""
						myORDS = ""
						myORKR = rsPersonas("Empresa")
						if isnull(rsPersonas("Empresa")) then myORKR = 0
						call GF_MGC("", "", clng(myORKR), myORDS) 
						Response.Write myORDS
					else
						Response.Write rsPersonas("nroLegajo")
					end if 
					%>
				</font></b></td>
				<% if session("AUPUSER") = "ADMIN" then %>
				<td align=center width="2%" nowrap> 
					<a title="<%=GF_Traducir("Reporte para Auditoria por Usuario")%>" target="_new" href="AUPAuditoria.asp?pUser=<%=rsPersonas("IdProfesional")%>"><img src="Images/printer1.gif"></a>
				</td>
				<% end if %>
			</tr>
			<%     
			rsPersonas.MoveNext 
	wend  
	call GF_BD_CONTROL (rsPersonas,oConn,"CLOSE",strSQL)	
	%>
</table>
