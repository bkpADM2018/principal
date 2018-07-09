<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosseguridad.asp"-->
<!--#include file="Includes/procedimientosBR.asp"-->
<%
'<!------------------------------------------------------------------->
'<!--                        INTRANET ACTISA                        -->
'<!--                                                               -->
'<!--               Programador: Ezequiel A. Bacarini               -->
'<!--               Fecha      : 07/04/2008                         -->
'<!--               Pagina     : BRExplorador.ASP                   -->
'<!--               Descripcion: Explorador de carpeta con imagenes -->
'<!------------------------------------------------------------------->
dim myContrato, myDoc, fso, cont, myCantFiles, MyImage, mySelected
dim MySubFolders0, MySubFolders1, MySubFolders2, MyFolder, MyFile
dim fsoImg, MyFileAux, dic
dim rs, sql
set dic = Server.CreateObject("Scripting.Dictionary")
myContrato = GF_Parametros7("pContrato","",6)
myDoc = GF_Parametros7("pDoc","",6)
myView = GF_Parametros7("pView","",6)
call CargarDocumentos



sub CargarDocumentos()
call GF_BD_Control_BR (rs, cn, "OPEN", "Select * from TipoDocumento")
	while not rs.EOF
		dic.Add trim(rs("cdTipoDocumento")), trim(rs("dsTipoDocumento"))
	rs.movenext
wend
call GF_BD_Control_BR (rs, cn, "CLOSE", "")
end sub

function getDescripcion(pCod)
if dic.Exists(pCod) then
	getDescripcion = dic(pCod)
else
	getDescripcion = pCod
end if
end function

set fsoImg = Server.CreateObject("Scripting.FileSystemObject")
'------------------------------------------------------------------------------------------------
%>
<table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
	<tr>
		<td>
			<div class="tabber" id="Tab3">
				<%  
				set fso = Server.CreateObject("Scripting.FileSystemObject")
				Set MySubFolders0 = fso.GetFolder(Server.MapPath("BR_Contracts\Scans"))
				for each MySubFolders1 in MySubFolders0.subFolders
					if MySubFolders1.name = myContrato then
						for each MySubFolders2 in MySubFolders1.subFolders
										%>
										<div class="tabbertab" title="<%=getDescripcion(trim(MySubFolders2.name))%>">
											<h2><%=getDescripcion(trim(MySubFolders2.name))%></h2>
											<table border=0 width="100%" cellpadding=0 cellspacing=0>
												<tr>
													<%
													'Response.Write "images/BR/" & MyFolder.name & ".ico (" & fsoImg.FileExists("images/BR/" & MyFolder.name & ".ico") & ")"
													if not fso.FileExists(Server.MapPath("images/BR_CONTRACTS/" & MySubFolders2.name & ".png")) then
														MyImage = "BR_Contracts/default.png"
													else
														MyImage = "BR_Contracts/" & MySubFolders2.name & ".png"
													end if	
													for each MyFile in MySubFolders2.files
														cont = cont + 1
														if myView = "I" then
														%>
														<td align="center" width="20%">
															<% 'Response.Write "openWin('BR_CONTRACTS/Scans/" & MySubFolders1.name & "/" & MySubFolders2.name & "/" & replace(MyFile.name,"'","\'") & "');" %> 
															<a title="<%=MyFile.name%>" style="cursor:pointer;" onclick="openWin('BR_CONTRACTS/Scans/<%=MySubFolders1.name%>/<%=MySubFolders2.name%>/<%=replace(MyFile.name,"'","\'")%>');">
																<img src="images/<%=MyImage%>">
																<br>
																<font class="Big">
																<b>
																<%	
																	myFileAux = myfile.name
																	if instr(myFileAux,".") > 0 then
																		myFileAux = mid(myFileAux, 1,len(myFileAux)-4)
																	end if
																	if len(myFileAux)>20 then
																		Response.write left(myFileAux,15) & "<br>"
																		Response.Write mid(myFileAux,16,15)
																	else
																		Response.write myFileAux & "<br><br>"
																	end if 
																%>
																</b>
																</font>
															</a>
														</td>
														<% 
														end if 
														if myView = "L" then
														%>
														<tr title="<%=MyFile.name%>" style="cursor:pointer;" onMouseOver="fcnResaltar(this)" onMouseOut="fcnNormal(this)" onclick="openWin('BR_CONTRACTS/Scans/<%=MySubFolders1.name%>/<%=MySubFolders2.name%>/<%=MyFile.name%>')">
															<td width="50%"><img src="images/<%=MyImage%>" width="13" height="13">&nbsp;<font class="small"><%=left(myfile.name,45)%></font></td>
															<td width="20%" colspan="1" align="center" valign="center"><font class="small"><%=MyFile.Type%></font></td>
															<td width="30%" colspan="2" valign="center"><font class="small"><%=MyFile.DateCreated%></font></td>
														</tr>
														<tr>
														<% 
														end if 
														'Response.Write "<br>CON(" & cont & ")"
														if cont mod 5 = 0 then Response.Write "</tr><tr>"
													next
													while not cont mod 5 = 0
														Response.Write "<td width='20%'>&nbsp;</td>"
														cont = cont + 1
													wend	
													cont=0
													%>
												</tr>
											</table>
										</div>
										<%
						next
					end if
				next		
				%>	
			</div>
		</td>
	</tr>
</table>	
