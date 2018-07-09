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
dim myContrato, myBuqueCD, myAno, myMes, myDia, myDoc, myImagen, myWhere, cuantos, myBA, myHAM, myBALink
myContrato = GF_Parametros7("pContrato","",6)
myAno = GF_Parametros7("pAno","",6)
myMes = GF_Parametros7("pMes","",6)
myDia = GF_Parametros7("pDia","",6)
myDoc = GF_Parametros7("pDoc","",6)
myCompania = GF_Parametros7("pCompania","",6)
myOperacion = GF_Parametros7("pOperacion","",6)
myCliente = GF_Parametros7("pCliente","",6)
if myCliente = "undefined" then myCliente = ""
myProducto = GF_Parametros7("pProducto","",6)
'Response.Write "Buque(" & myBuque & "), BuqueCD(" & myBuqueCD & "), ANO(" & myAno & "), MES(" & myMes & "), DIA(" + myDia + "), Doc(" & myDoc & "), BA(" & myBA & "), HAM(" & myHAM & ")"
if len(myContrato) > 0 then 
	call armarWhere("NroContrato='" & myContrato & "'")
else
	if len(myOperacion) > 0 then call armarWhere("cdOperacion = '" & myOperacion & "'")
	if len(myCompania) > 0 then call armarWhere("cdCompania = '" & myCompania & "'")
	if isdate(myano & "/" & myMes & "/" & myDia) then  call armarWhere("FechaCierre = '" & myAno & "-" & myDia & "-" & myMes & " 00:00:00.0'")
	if len(myCliente) > 0 then  call armarWhere("dsCliente = '" & myCliente & "'")
	if len(myProducto) > 0 then  call armarWhere("cdProducto = '" & myProducto & "'")
end if	
if len(myWhere)>0 then sql = "Select * from Contratos Cnt inner join Clientes Cli on Cnt.cdCliente=Cli.cdCliente " & myWhere
'Response.Write "<hr>" & sql
'Response.End 
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


cont2 = 0
'--------------------------------------------------------------------------------------------
sub armarWhere(pCondicion)
	if len(myWhere) > 0 then 
		myWhere = myWhere & " and " & pCondicion
	else
		myWhere = " where "	& pCondicion 
	end if	 
end sub
'------------------------------------------------------------------------------------------------
%>
<html>
<head>
<meta http-equiv="Cache-Control" content="no-cache, mustrevalidate">
<meta http-equiv="Pragma" content="no-cache">
<title><%=GF_Traducir("Intranet ActiSA - Detalle del Contrato")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
	<table border="0" cellspacing="1" cellpadding="1" width="100%" align="center" rules=groups>
		<%  
		set fso = Server.CreateObject("Scripting.FileSystemObject")
		if fso.FolderExists(Server.MapPath("BR_CONTRACTS\Scans")) then
			Set MySubFolders0 = fso.GetFolder(Server.MapPath("BR_CONTRACTS\Scans"))
				for each MySubFolders1 in MySubFolders0.subFolders 'Buques
					if (instr(myContrato, "/" & MySubFolders1.name & "/")>0 or myContrato="") then				
						for each MySubFolders3 in MySubFolders1.subFolders 'Doc
							if (MySubFolders3.name = myDoc or myDoc = "") then
								'if myBA <> "" then myBALink = MySubFolders2.name
								cont = cont + MySubFolders3.files.count
							end if	
						next
					end if
					'if cont2 = 0 then 
					'	Response.Write "<tr>"
					'else	
					
					if cont2 mod 4 = 0 and cont2<>0 then	
						Response.Write "<tr><td><br></td></tr><tr>"
						cont2 = 0
					end if	
					'end if
					if cont > 0 then
						hubo = true
						cont2 = cont2 + 1
						%>
						<td width="25%" align="left">
							<a href="BRCExplorador.asp?pContrato=<%=MySubFolders1.name%>&pDoc=<%=myDoc%>">
								<img src="images/Contract16_7.gif">
								&nbsp;<font size="+2" color="#000000"><b><%=MySubFolders1.name%> (<%=cont%>)</b></font>
							</a>
						</td>
						<%
					end if
					cont = 0
					
				next
				if hubo then
					for i=(MySubFolders0.subFolders.count mod 4 + 1) to 4 'Completo las celdas que faltan
						Response.Write "<td width='25%'>&nbsp;</td>"
					next
				else
					%>
					<td width="100%" class="tderror" align="center">
						<%=GF_Traducir("No se han encontrado resultados!")%>
					</td>
					<%
				end if
		end if 
		%>
		</tr>
	</table>
</html>