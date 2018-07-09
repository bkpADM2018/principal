<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosBR.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosseguridad.asp"-->
<%
'<!------------------------------------------------------------------->
'<!--                        INTRANET ACTISA                        -->
'<!--                                                               -->
'<!--               Programador: Ezequiel A. Bacarini               -->
'<!--               Fecha      : 07/04/2008                         -->
'<!--               Pagina     : BRExplorador.ASP                   -->
'<!--               Descripcion: Explorador de carpeta con imagenes -->
'<!------------------------------------------------------------------->
dim fso, MySubFolders0, MySubFolders1, MySubFolders2, MySubFolders3, MySubFolders4, MySubFolders5
dim cont, cont2, hubo, rs, cn, sql
dim myBuque, myBuqueCD, myAno, myMes, myDia, myDoc, myImagen, myWhere, cuantos, myBA, myHAM
dim impBuque, impBA, fso1, myImage
set fso = Server.CreateObject("Scripting.FileSystemObject")
myContrato = GF_Parametros7("pContrato","",6)
myOperacion = GF_Parametros7("pOperacion","",6)
myCompania = GF_Parametros7("pCompania","",6)
myAno = GF_Parametros7("pAno","",6)
myMes = GF_Parametros7("pMes","",6)
myDia = GF_Parametros7("pDia","",6)
myDoc = GF_Parametros7("pDoc","",6)
myCliente = GF_Parametros7("pCliente","",6)

if len(myContrato) > 0 then call armarWhere("NroContrato='" & myContrato & "'")
if len(myOperacion) > 0 then call armarWhere("cdOperacion='" & myOperacion & "'")
if len(myCompania) > 0 then call armarWhere("cdCompania='" & myCompania & "'")
if len(myCliente) > 0 then call armarWhere("cdCliente='" & myCliente & "'")
if isdate(myano & "/" & myMes & "/" & myDia) then call armarWhere("FechaCierre='" & myAno & "-" & myDia & "-" & myMes & " 00:00:00.0'")

if len(myWhere)>0 then sql = "Select * from Contratos " & myWhere
if len(sql)>0 then
	myContrato = "0"
	call GF_BD_Control_BR (rs, cn, "OPEN", sql & " order by convert(int,right(NroContrato,5)) desc")
	while not rs.EOF
		myContrato = myContrato & "/" & trim(rs("NroContrato"))
		rs.movenext
	wend
	call GF_BD_Control_BR (rs, cn, "CLOSE", sql)
	myContrato = myContrato & "/"
end if
'Response.Write sql & " order by convert(int,right(NroContrato,5)) desc"
'Response.Write "<hr>" & myContrato &  myDoc
'Response.End 
'--------------------------------------------------------------------------------------------
sub armarWhere(pCondicion)
	if len(myWhere) > 0 then 
		myWhere = myWhere & " and " & pCondicion
	else
		myWhere = " where "	& pCondicion 
	end if	 
end sub
'Response.Write myBA
%>
<html>
<head>
<meta http-equiv="Cache-Control" content="no-cache, mustrevalidate">
<meta http-equiv="Pragma" content="no-cache">
<title><%=GF_Traducir("Intranet ActiSA - Detalle del Buque")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head><BODY bgColor=LightSeaGreen>
	<%
	'dim fso
	'set fso = Server.CreateObject("Scripting.FileSystemObject")
	if not fso.FileExists(Server.MapPath("images/BR_CONTRACTS/" & myDoc & "48.png")) then
		MyImage = "BR_Contracts/default1_48.png"
	else
		MyImage = "BR_Contracts/" & myDoc & "48.png"
	end if	
	%>
	<table border="0" cellspacing="0" cellpadding="0" width="80%" align="center" rules=rows>
		<tr>
			<td width="6%">
				<img src="images/<%=MyImage%>">
			</td>
			<td valign="center" align="left">
				<font class="bigger"> <%= GF_Traducir("Archivos encontrados")%></font>
			</td>
		</tr>
	</table>		

	<table border="1" cellspacing="0" cellpadding="0" width="60%" align="center" rules=rows>
		<tr class="reg_header_navdos">
			<td align="left">	<font><b><%=GF_Traducir("Documento")%>		</b></font>
			<td align="center">	<font><b><%=GF_Traducir("Nro. Contrato")%>	</b></font>
		</tr>
		<%  
		if fso.FolderExists(Server.MapPath("BR_Contracts\Scans")) then
			Set MySubFolders0 = fso.GetFolder(Server.MapPath("BR_Contracts\Scans"))
				for each MySubFolders1 in MySubFolders0.subFolders '
							if (instr(myContrato, "/" & MySubFolders1.name & "/")>0 or myContrato="") then
								for each MySubFolders3 in MySubFolders1.subFolders 'Doc
									if (MySubFolders3.name = myDoc) then
										impContrato = MySubFolders1.name
										for each MySubFolders4 in MySubFolders3.files 'Files
												%>
												<tr style="cursor:pointer;" onMouseOver="fcnResaltar(this)" onMouseOut="fcnNormal(this)" title="<%=GF_Traducir("Abrir el documento")%>" onclick="openWin('BR_CONTRACTS/Scans/<%=MySubFolders1.name%>/<%=MySubFolders3.name%>/<%=MySubFolders4.name%>')">
													<td title="<%=MySubFolders4.name%>" valign="center">
															<font><%=left(MySubFolders4.name,30)%></font>
													</td>
													<td valign="center" align="center">
															<font><%=impContrato%></font>
													</td>
												</tr>
												<%
												hubo = true
										next
										cont = cont + MySubFolders3.files.count
									end if	
								next
							end if
				next
				if not hubo then
					%>
					<tr>
						<td colspan="4" align="center">
							<font color="red"><b><%=GF_Traducir("No se han encontrado resultados!")%></b></font>
						</td>
					</tr>	
					<%
				end if
		end if 
		%>
		</tr>
	</table>
</html>